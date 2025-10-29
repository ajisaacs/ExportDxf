using ExportDXF.Extensions;
using ExportDXF.ItemExtractors;
using ExportDXF.Models;
using ExportDXF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Environment = System.Environment;
using System.IO.Compression;

namespace ExportDXF.Services
{
    public interface IDxfExportService
    {
        /// <summary>
        /// Exports the document specified in the context to DXF format.
        /// </summary>
        /// <param name="context">The export context containing all necessary information.</param>
        void Export(ExportContext context);
    }

    /// <summary>
    /// Service responsible for orchestrating the export of SolidWorks documents to DXF format.
    /// </summary>
    public class DxfExportService : IDxfExportService
    {
        private readonly ISolidWorksService _solidWorksService;
        private readonly IBomExtractor _bomExtractor;
        private readonly IPartExporter _partExporter;
        private readonly IDrawingExporter _drawingExporter;
        private readonly ICutFabApiClient _apiClient;

        public DxfExportService(
            ISolidWorksService solidWorksService,
            IBomExtractor bomExtractor,
            IPartExporter partExporter,
            IDrawingExporter drawingExporter,
            ICutFabApiClient apiClient)
        {
            _solidWorksService = solidWorksService ?? throw new ArgumentNullException(nameof(solidWorksService));
            _bomExtractor = bomExtractor ?? throw new ArgumentNullException(nameof(bomExtractor));
            _partExporter = partExporter ?? throw new ArgumentNullException(nameof(partExporter));
            _drawingExporter = drawingExporter ?? throw new ArgumentNullException(nameof(drawingExporter));
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
        }

        /// <summary>
        /// Exports the document specified in the context to DXF format.
        /// </summary>
        public void Export(ExportContext context)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            ValidateContext(context);
            SetupExportContext(context);

            var startTime = DateTime.Now;

            try
            {
                _solidWorksService.EnableUserControl(false);

                switch (context.ActiveDocument.DocumentType)
                {
                    case DocumentType.Part:
                        ExportPart(context);
                        break;

                    case DocumentType.Assembly:
                        ExportAssembly(context);
                        break;

                    case DocumentType.Drawing:
                        ExportDrawing(context);
                        break;

                    default:
                        LogProgress(context, "Unknown document type.", Color.Red);
                        break;
                }
            }
            finally
            {
                CleanupExportContext(context);
                _solidWorksService.EnableUserControl(true);

                var duration = DateTime.Now - startTime;
                LogProgress(context, $"Run time: {duration.ToReadableFormat()}", null);
            }
        }

        #region Export Methods by Document Type

        private void ExportPart(ExportContext context)
        {
            LogProgress(context, "Active document is a Part", null);

            var part = context.ActiveDocument.NativeDocument as PartDoc;
            if (part == null)
            {
                LogProgress(context, "Failed to get part document.", Color.Red);
                return;
            }

            var tempDir = CreateTempWorkDir();
            _partExporter.ExportSinglePart(part, tempDir, context);
        }

        private void ExportAssembly(ExportContext context)
        {
            LogProgress(context, "Active document is an Assembly", null);
            LogProgress(context, "Fetching components...", null);

            var assembly = context.ActiveDocument.NativeDocument as AssemblyDoc;
            if (assembly == null)
            {
                LogProgress(context, "Failed to get assembly document.", Color.Red);
                return;
            }

            var items = ExtractItemsFromAssembly(assembly, context);

            if (items == null || items.Count == 0)
            {
                LogProgress(context, "No items found in assembly.", Color.DarkBlue);
                return;
            }

            LogProgress(context, $"Found {items.Count} item(s).", null);

            var tempDir = CreateTempWorkDir();
            ExportItems(items, tempDir, context, drawingId: null);
        }

        private void ExportDrawing(ExportContext context)
        {
            LogProgress(context, "Active document is a Drawing", null);
            LogProgress(context, "Finding BOM tables...", null);

            var drawing = context.ActiveDocument.NativeDocument as DrawingDoc;
            if (drawing == null)
            {
                LogProgress(context, "Failed to get drawing document.", Color.Red);
                return;
            }

            var items = _bomExtractor.ExtractFromDrawing(drawing, context.ProgressCallback);

            if (items == null || items.Count == 0)
            {
                LogProgress(context, "Error: Bill of materials not found.", Color.Red);
                return;
            }

            LogProgress(context, $"Found {items.Count} component(s)", null);

            var tempDir = CreateTempWorkDir();

            // Determine drawing number
            var drawingNumber = ParseDrawingNumber(context);
            if (string.IsNullOrWhiteSpace(drawingNumber))
            {
                LogProgress(context, "Warning: Could not determine drawing number for API upload.", Color.DarkBlue);
            }

            // Resolve drawing ID if possible
            int? drawingId = null;
            if (!string.IsNullOrWhiteSpace(drawingNumber))
            {
                drawingId = _apiClient.ResolveDrawingIdAsync(drawingNumber).GetAwaiter().GetResult();
                // Fallback: if resolve endpoint not available or failed, search equipment details
                if (drawingId == null && context.EquipmentId.HasValue)
                {
                    try
                    {
                        var drawings = _apiClient.GetDrawingsForEquipmentAsync(context.EquipmentId.Value).GetAwaiter().GetResult();
                        if (drawings != null)
                        {
                            // Match by exact DrawingNumber (case-insensitive, trimmed)
                            var match = drawings.FirstOrDefault(d => string.Equals(d.DrawingNumber?.Trim(), drawingNumber.Trim(), StringComparison.OrdinalIgnoreCase));
                            if (match != null) drawingId = match.ID;
                        }
                    }
                    catch { }
                }
                if (drawingId == null)
                {
                    // If equipment is provided, create the drawing on the API
                    if (context.EquipmentId.HasValue)
                    {
                        var create = _apiClient.CreateDrawingWithInfoAsync(context.EquipmentId.Value, drawingNumber).GetAwaiter().GetResult();
                        if (create != null && create.Success && create.Data.HasValue)
                        {
                            drawingId = create.Data;
                            LogProgress(context, "Created drawing '" + drawingNumber + "' (ID " + drawingId + ") for equipment " + context.EquipmentId, Color.Green);
                        }
                        else
                        {
                            var code = create != null ? create.StatusCode.ToString() : "?";
                            var err = create != null ? (create.Error ?? create.RawBody) : null;
                            if (!string.IsNullOrWhiteSpace(err) && err.Length > 180) err = err.Substring(0, 180) + "...";
                            LogProgress(context, "Warning: Could not create drawing '" + drawingNumber + "' on API (status " + code + "). " + (err ?? string.Empty), Color.DarkBlue);
                        }
                    }
                    else
                    {
                        LogProgress(context, $"Warning: Drawing '{drawingNumber}' not found in API; uploads will be skipped.", Color.DarkBlue);
                    }
                }

                // Export drawing to PDF first
                _drawingExporter.ExportToPdf(drawing, tempDir, context);

                // Upload PDF if we have a drawing number
                try
                {
                    if (!string.IsNullOrWhiteSpace(drawingNumber))
                    {
                        var pdfs = Directory.GetFiles(tempDir, "*.pdf");
                        var pdfName = pdfs.Length > 0 ? pdfs[0] : null;
                        if (pdfName != null)
                        {
                            var uploadedBy = Environment.UserName;
                            var ok = _apiClient.UploadDrawingPdfAsync(drawingNumber, pdfName, uploadedBy, null).GetAwaiter().GetResult();
                            LogProgress(context, ok ? $"Uploaded PDF for '{drawingNumber}'" : $"Failed to upload PDF for '{drawingNumber}'", ok ? Color.Green : Color.Red);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogProgress(context, $"PDF upload error: {ex.Message}", Color.Red);
                }

                // If we still don't have an ID, resolve again after PDF upload (server may create on upload)
                if (!drawingId.HasValue && !string.IsNullOrWhiteSpace(drawingNumber))
                {
                    var resolved = _apiClient.ResolveDrawingIdAsync(drawingNumber).GetAwaiter().GetResult();
                    if (!resolved.HasValue && context.EquipmentId.HasValue)
                    {
                        try
                        {
                            var drawings = _apiClient.GetDrawingsForEquipmentAsync(context.EquipmentId.Value).GetAwaiter().GetResult();
                            if (drawings != null)
                            {
                                var match = drawings.FirstOrDefault(d => string.Equals(d.DrawingNumber?.Trim(), drawingNumber.Trim(), StringComparison.OrdinalIgnoreCase));
                                if (match != null) resolved = match.ID;
                            }
                        }
                        catch { }
                    }
                    if (resolved.HasValue)
                    {
                        drawingId = resolved;
                        LogProgress(context, $"Resolved drawing ID after PDF upload: {drawingId}", Color.Green);
                    }
                }

                // Then export parts to DXF and upload per-file
                ExportItems(items, tempDir, context, drawingId);

                // Attempt to auto-link templates at the end
                try
                {
                    if (drawingId.HasValue)
                    {
                        _apiClient.AutoLinkTemplatesAsync(drawingId.Value).GetAwaiter().GetResult();
                    }
                }
                catch { }
            }
        }

        #endregion

        #region Context Management

        private void SetupExportContext(ExportContext context)
        {
            // Set up SolidWorks application reference
            context.SolidWorksApp = _solidWorksService.GetNativeSldWorks();

            if (context.SolidWorksApp == null)
            {
                throw new InvalidOperationException("SolidWorks service is not connected.");
            }

            // Set up drawing template path
            context.TemplateDrawing = null;

            LogProgress(context, "Export context initialized", null);
        }

        private void CleanupExportContext(ExportContext context)
        {
            try
            {
                // Clean up template drawing if it was created
                context.CleanupTemplateDrawing();
            }
            catch (Exception ex)
            {
                LogProgress(context, $"Warning: Failed to cleanup template drawing: {ex.Message}", Color.DarkBlue);
                // Don't throw - this is cleanup code
            }
        }

        #endregion

        #region Item Processing

        private List<Item> ExtractItemsFromAssembly(AssemblyDoc assembly, ExportContext context)
        {
            try
            {
                var extractor = new AssemblyItemExtractor(assembly)
                {
                    TopLevelOnly = false
                };

                return extractor.GetItems();
            }
            catch (Exception ex)
            {
                LogProgress(context, $"Failed to extract items from assembly: {ex.Message}", Color.Red);
                return new List<Item>();
            }
        }

        private void ExportItems(List<Item> items, string saveDirectory, ExportContext context, int? drawingId)
        {
            LogProgress(context, "", null);

            int successCount = 0;
            int failureCount = 0;

            foreach (var item in items)
            {
                if (context.CancellationToken.IsCancellationRequested)
                {
                    LogProgress(context, "Export canceled by user.", Color.DarkBlue);
                    return;
                }

                try
                {
                    // PartExporter will handle template drawing creation through context
                    _partExporter.ExportItem(item, saveDirectory, context);

                    // Always create BOM item first if we have a drawing
                    if (drawingId.HasValue)
                    {
                        var dto = new
                        {
                            DrawingID = drawingId.Value,
                            ItemNo = item.ItemNo,
                            PartNo = !string.IsNullOrEmpty(item.FileName) ? item.FileName : item.PartName,
                            Qty = (int?)item.Quantity,
                            Description = string.IsNullOrWhiteSpace(item.Description) ? null : item.Description,
                            PartName = string.IsNullOrWhiteSpace(item.PartName) ? null : item.PartName,
                            ConfigurationName = string.IsNullOrWhiteSpace(item.Configuration) ? null : item.Configuration,
                            Material = string.IsNullOrWhiteSpace(item.Material) ? null : item.Material,
                            SortOrder = 0,
                            CutTemplateID = (int?)null,
                            FormProgramID = (int?)null
                        };
                        var bomId = _apiClient.CreateBomItemAsync(dto).GetAwaiter().GetResult();
                        LogProgress(context, bomId.HasValue ? $"Created BOM item for {item.ItemNo ?? item.PartName}" : $"Failed to create BOM item for {item.ItemNo ?? item.PartName}", bomId.HasValue ? Color.Green : Color.Red);
                    }

                    if (!string.IsNullOrEmpty(item.FileName))
                    {
                        successCount++;

                        // If we know the drawing, upload DXF
                        if (drawingId.HasValue)
                        {
                            var dxfPath = Path.Combine(saveDirectory, item.FileName + ".dxf");
                            if (File.Exists(dxfPath))
                            {
                                // Zip just this file
                                string zipPath = CreateZipWithSingleFile(dxfPath);
                                try
                                {
                                    // Pass thickness and kfactor from the item
                                    double? thickness = item.Thickness > 0 ? item.Thickness : (double?)null;
                                    double? kfactor = item.KFactor > 0 ? item.KFactor : (double?)null;
                                    var okZip = _apiClient.UploadDxfZipAsync(drawingId.Value, zipPath, thickness, kfactor).GetAwaiter().GetResult();
                                    LogProgress(context, okZip ? $"Uploaded DXF: {Path.GetFileName(dxfPath)}" : $"Failed to upload DXF: {Path.GetFileName(dxfPath)}", okZip ? Color.Green : Color.Red);
                                }
                                finally
                                {
                                    try { if (File.Exists(zipPath)) File.Delete(zipPath); } catch { }
                                }
                            }
                        }
                    }
                    else
                    {
                        failureCount++;
                    }
                }
                catch (Exception ex)
                {
                    LogProgress(context, $"Error exporting item {item.ItemNo}: {ex.Message}", Color.Red);
                    failureCount++;
                }

                LogProgress(context, "", null);
            }

            LogProgress(context, $"Export complete: {successCount} succeeded, {failureCount} failed",
                failureCount > 0 ? Color.DarkBlue : Color.Green);
        }

        #endregion

        #region Temp + Upload Helpers

        private string CreateTempWorkDir()
        {
            var path = Path.Combine(Path.GetTempPath(), "ExportDXF-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }

        private string ParseDrawingNumber(ExportContext context)
        {
            // Prefer prefix (e.g., "5007 A02 PT"), fallback to active document title
            var candidate = context?.FilePrefix;
            var info = string.IsNullOrWhiteSpace(candidate) ? null : DrawingInfo.Parse(candidate);
            if (info == null)
            {
                var title = context?.ActiveDocument?.Title;
                info = string.IsNullOrWhiteSpace(title) ? null : DrawingInfo.Parse(title);
            }
            return info != null ? ($"{info.EquipmentNo} {info.DrawingNo}") : null;
        }

        private string CreateZipWithSingleFile(string filePath)
        {
            var zipPath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".zip");
            if (File.Exists(zipPath)) File.Delete(zipPath);
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, System.IO.Compression.ZipArchiveMode.Create))
            {
                zip.CreateEntryFromFile(filePath, Path.GetFileName(filePath));
            }
            return zipPath;
        }

        #endregion

        #region Helper Methods

        private void ValidateContext(ExportContext context)
        {
            if (context.ActiveDocument == null)
                throw new ArgumentException("ActiveDocument cannot be null.", nameof(context));

            if (context.ProgressCallback == null)
                throw new ArgumentException("ProgressCallback cannot be null.", nameof(context));
        }

        private void LogProgress(ExportContext context, string message, Color? color)
        {
            context.ProgressCallback?.Invoke(message, color);
        }

        

        #endregion
    }
}


