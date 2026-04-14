using ExportDXF.ApiClient;
using ExportDXF.Extensions;
using ExportDXF.ItemExtractors;
using ExportDXF.Models;
using ExportDXF.Utilities;
using ExportDXF;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExportDXF.Services
{
    public interface IDxfExportService
    {
        /// <summary>
        /// Exports the document specified in the context to DXF format.
        /// </summary>
        /// <param name="context">The export context containing all necessary information.</param>
        Task ExportAsync(ExportContext context);
    }

    /// <summary>
    /// Service responsible for orchestrating the export of SolidWorks documents to DXF format.
    /// Files are generated locally in a temp directory, then uploaded to the API for storage and versioning.
    /// </summary>
    public class DxfExportService : IDxfExportService
    {
        private readonly ISolidWorksService _solidWorksService;
        private readonly IBomExtractor _bomExtractor;
        private readonly IPartExporter _partExporter;
        private readonly IDrawingExporter _drawingExporter;
        private readonly IFabWorksApiClient _apiClient;

        public DxfExportService(
            ISolidWorksService solidWorksService,
            IBomExtractor bomExtractor,
            IPartExporter partExporter,
            IDrawingExporter drawingExporter,
            IFabWorksApiClient apiClient)
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
        public async Task ExportAsync(ExportContext context)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            ValidateContext(context);
            SetupExportContext(context);

            var startTime = DateTime.Now;
            var tempDir = CreateTempWorkDir();

            try
            {
                _solidWorksService.EnableUserControl(false);

                var drawingNumber = ParseDrawingNumber(context);

                switch (context.ActiveDocument.DocumentType)
                {
                    case DocumentType.Part:
                        await ExportPartAsync(context, tempDir, drawingNumber);
                        break;

                    case DocumentType.Assembly:
                        await ExportAssemblyAsync(context, tempDir, drawingNumber);
                        break;

                    case DocumentType.Drawing:
                        await ExportDrawingAsync(context, drawingNumber, tempDir);
                        break;

                    default:
                        LogProgress(context, "Unknown document type.", LogLevel.Error);
                        break;
                }
            }
            finally
            {
                CleanupExportContext(context);
                _solidWorksService.EnableUserControl(true);
                CleanupTempDir(tempDir);

                var duration = DateTime.Now - startTime;
                LogProgress(context, $"Run time: {duration.ToReadableFormat()}");
            }
        }

        #region Export Methods by Document Type

        private async Task ExportPartAsync(ExportContext context, string tempDir, string drawingNumber)
        {
            LogProgress(context, "Active document is a Part");

            var part = context.ActiveDocument.NativeDocument as PartDoc;
            if (part == null)
            {
                LogProgress(context, "Failed to get part document.", LogLevel.Error);
                return;
            }

            var exportRecord = await CreateExportRecordAsync(context, drawingNumber);
            var item = _partExporter.ExportSinglePart(part, tempDir, context);

            if (item != null)
            {
                // Check if this part+config already has a BOM item for this drawing
                var existingItemNo = await FindExistingItemNoAsync(exportRecord?.Id, item.PartName, item.Configuration);
                item.ItemNo = existingItemNo ?? await GetNextItemNumberAsync(drawingNumber);

                var bomItem = new BomItem
                {
                    ExportRecordId = exportRecord?.Id ?? 0,
                    ItemNo = item.ItemNo,
                    PartNo = item.FileName ?? item.PartName ?? "",
                    SortOrder = 0,
                    Qty = item.Quantity,
                    TotalQty = item.Quantity,
                    Description = item.Description ?? "",
                    PartName = item.PartName ?? "",
                    ConfigurationName = item.Configuration ?? "",
                    Material = item.Material ?? ""
                };

                // Upload DXF to API and get stored path
                if (!string.IsNullOrEmpty(item.LocalTempPath))
                {
                    var uploadResult = await UploadDxfAsync(item, context);
                    if (uploadResult != null)
                    {
                        bomItem.CutTemplate = new CutTemplate
                        {
                            DxfFilePath = uploadResult.StoredFilePath,
                            ContentHash = item.ContentHash,
                            Thickness = item.Thickness > 0 ? item.Thickness : null,
                            KFactor = item.KFactor > 0 ? item.KFactor : null,
                            DefaultBendRadius = item.BendRadius > 0 ? item.BendRadius : null
                        };
                    }
                }

                context.BomItemCallback?.Invoke(bomItem);

                if (exportRecord != null)
                    await SaveBomItemAsync(exportRecord.Id, bomItem, context);
            }
        }

        private async Task ExportAssemblyAsync(ExportContext context, string tempDir, string drawingNumber)
        {
            LogProgress(context, "Active document is an Assembly");
            LogProgress(context, "Fetching components...");

            var assembly = context.ActiveDocument.NativeDocument as AssemblyDoc;
            if (assembly == null)
            {
                LogProgress(context, "Failed to get assembly document.", LogLevel.Error);
                return;
            }

            var items = ExtractItemsFromAssembly(assembly, context);

            if (items == null || items.Count == 0)
            {
                LogProgress(context, "No items found in assembly.", LogLevel.Warning);
                return;
            }

            LogProgress(context, $"Found {items.Count} item(s).");

            var exportRecord = await CreateExportRecordAsync(context, drawingNumber);

            // Check existing BOM items and reuse item numbers, or assign new ones
            var nextNum = int.Parse(await GetNextItemNumberAsync(drawingNumber));
            foreach (var item in items)
            {
                if (string.IsNullOrWhiteSpace(item.ItemNo))
                {
                    var existingItemNo = await FindExistingItemNoAsync(exportRecord?.Id, item.PartName, item.Configuration);
                    if (existingItemNo != null)
                    {
                        item.ItemNo = existingItemNo;
                    }
                    else
                    {
                        item.ItemNo = nextNum.ToString();
                        nextNum++;
                    }
                }
            }

            await ExportItemsAsync(items, tempDir, context, exportRecord?.Id);
        }

        private async Task ExportDrawingAsync(ExportContext context, string drawingNumber, string tempDir)
        {
            LogProgress(context, "Active document is a Drawing");
            LogProgress(context, "Finding BOM tables...");

            var drawing = context.ActiveDocument.NativeDocument as DrawingDoc;
            if (drawing == null)
            {
                LogProgress(context, "Failed to get drawing document.", LogLevel.Error);
                return;
            }

            var items = _bomExtractor.ExtractFromDrawing(drawing, context.ProgressCallback);

            if (items == null || items.Count == 0)
            {
                LogProgress(context, "Error: Bill of materials not found.", LogLevel.Error);
                return;
            }

            LogProgress(context, $"Found {items.Count} component(s)");

            // Export drawing to PDF in temp dir
            _drawingExporter.ExportToPdf(drawing, tempDir, context);

            // Create export record via API
            var exportRecord = await CreateExportRecordAsync(context, drawingNumber);

            // Upload PDF to API with versioning
            try
            {
                var pdfs = Directory.GetFiles(tempDir, "*.pdf");
                if (pdfs.Length > 0)
                {
                    var pdfTempPath = pdfs[0];
                    var pdfHash = ContentHasher.ComputeFileHash(pdfTempPath);

                    var uploadResult = await _apiClient.UploadPdfAsync(
                        pdfTempPath,
                        context.Equipment,
                        context.DrawingNo,
                        pdfHash,
                        exportRecord?.Id);

                    if (uploadResult != null)
                    {
                        if (uploadResult.WasUnchanged)
                            LogProgress(context, $"PDF unchanged: {uploadResult.FileName}", LogLevel.Info);
                        else if (uploadResult.IsNewFile)
                            LogProgress(context, $"Saved PDF: {uploadResult.FileName}", LogLevel.Info);
                        else
                            LogProgress(context, $"PDF updated: {uploadResult.FileName}", LogLevel.Info);
                    }
                }
            }
            catch (Exception ex)
            {
                LogProgress(context, $"PDF upload error: {ex.Message}", LogLevel.Error);
            }

            // Export parts to DXF and save BOM items
            await ExportItemsAsync(items, tempDir, context, exportRecord?.Id);
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

            LogProgress(context, "Export context initialized");
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
                LogProgress(context, $"Warning: Failed to cleanup template drawing: {ex.Message}", LogLevel.Warning);
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
                LogProgress(context, $"Failed to extract items from assembly: {ex.Message}", LogLevel.Error);
                return new List<Item>();
            }
        }

        private async Task ExportItemsAsync(List<Item> items, string tempDir, ExportContext context, int? exportRecordId = null)
        {
            int successCount = 0;
            int skippedCount = 0;
            int failureCount = 0;
            int sortOrder = 0;

            foreach (var item in items)
            {
                if (context.CancellationToken.IsCancellationRequested)
                {
                    LogProgress(context, "Export canceled by user.", LogLevel.Warning);
                    return;
                }

                try
                {
                    // PartExporter will handle template drawing creation through context
                    _partExporter.ExportItem(item, tempDir, context);

                    // Always create BomItem for every item (sheet metal or not)
                    var bomItem = new BomItem
                    {
                        ExportRecordId = exportRecordId ?? 0,
                        ItemNo = item.ItemNo ?? "",
                        PartNo = item.FileName ?? item.PartName ?? "",
                        SortOrder = sortOrder++,
                        Qty = item.Quantity,
                        TotalQty = item.Quantity,
                        Description = item.Description ?? "",
                        PartName = item.PartName ?? "",
                        ConfigurationName = item.Configuration ?? "",
                        Material = item.Material ?? ""
                    };

                    // Only upload and create CutTemplate if DXF was exported successfully
                    if (!string.IsNullOrEmpty(item.LocalTempPath))
                    {
                        successCount++;

                        var uploadResult = await UploadDxfAsync(item, context);
                        if (uploadResult != null)
                        {
                            bomItem.CutTemplate = new CutTemplate
                            {
                                DxfFilePath = uploadResult.StoredFilePath,
                                ContentHash = item.ContentHash,
                                Thickness = item.Thickness > 0 ? item.Thickness : null,
                                KFactor = item.KFactor > 0 ? item.KFactor : null,
                                DefaultBendRadius = item.BendRadius > 0 ? item.BendRadius : null
                            };
                        }
                    }
                    else
                    {
                        skippedCount++;
                    }

                    // Add to UI
                    context.BomItemCallback?.Invoke(bomItem);

                    // Save BOM item via API if we have an export record
                    if (exportRecordId.HasValue)
                    {
                        await SaveBomItemAsync(exportRecordId.Value, bomItem, context);
                    }
                }
                catch (Exception ex)
                {
                    LogProgress(context, $"Error exporting item {item.ItemNo}: {ex.Message}", LogLevel.Error);
                    failureCount++;
                }
            }

            var summary = $"Export complete: {successCount} exported, {skippedCount} skipped";
            if (failureCount > 0)
                summary += $", {failureCount} failed";
            LogProgress(context, summary, failureCount > 0 ? LogLevel.Warning : LogLevel.Info);

            if (exportRecordId.HasValue)
            {
                LogProgress(context, $"BOM items saved (ExportRecord ID: {exportRecordId.Value})", LogLevel.Info);
            }
        }

        #endregion

        #region File Upload

        private async Task<ApiFileUploadResponse> UploadDxfAsync(Item item, ExportContext context)
        {
            try
            {
                var result = await _apiClient.UploadDxfAsync(
                    item.LocalTempPath,
                    context.Equipment,
                    context.DrawingNo,
                    item.ItemNo,
                    item.ContentHash);

                if (result.WasUnchanged)
                    LogProgress(context, $"DXF unchanged: {result.FileName}", LogLevel.Info);
                else if (result.IsNewFile)
                    LogProgress(context, $"Exported: {result.FileName}", LogLevel.Info);
                else
                    LogProgress(context, $"DXF updated: {result.FileName}", LogLevel.Info);

                return result;
            }
            catch (Exception ex)
            {
                LogProgress(context, $"DXF upload failed for {item.FileName}: {ex.Message}", LogLevel.Warning);
                return null;
            }
        }

        #endregion

        #region API Helpers

        private async Task<ExportRecord> CreateExportRecordAsync(ExportContext context, string drawingNumber)
        {
            try
            {
                var dto = await _apiClient.CreateExportAsync(
                    drawingNumber ?? context.ActiveDocument.Title,
                    context.Equipment ?? "",
                    context.DrawingNo ?? "",
                    context.ActiveDocument.FilePath,
                    "", // Output folder is now managed by the API
                    context.Title);

                var record = new ExportRecord
                {
                    Id = dto.Id,
                    DrawingNumber = dto.DrawingNumber,
                    EquipmentNo = dto.EquipmentNo,
                    DrawingNo = dto.DrawingNo,
                    SourceFilePath = dto.SourceFilePath,
                    OutputFolder = dto.OutputFolder,
                    ExportedAt = dto.ExportedAt,
                    ExportedBy = dto.ExportedBy
                };

                LogProgress(context, $"Created export record (ID: {record.Id})", LogLevel.Info);
                return record;
            }
            catch (Exception ex)
            {
                LogProgress(context, $"API error creating export record: {ex.Message}", LogLevel.Error);
                return null;
            }
        }

        private async Task<string> FindExistingItemNoAsync(int? exportRecordId, string partName, string configurationName)
        {
            if (!exportRecordId.HasValue)
                return null;

            try
            {
                var existing = await _apiClient.FindExistingBomItemAsync(exportRecordId.Value, partName, configurationName);
                return existing?.ItemNo;
            }
            catch
            {
                return null;
            }
        }

        private async Task<string> GetNextItemNumberAsync(string drawingNumber)
        {
            if (string.IsNullOrEmpty(drawingNumber))
                return "1";

            try
            {
                return await _apiClient.GetNextItemNumberAsync(drawingNumber);
            }
            catch
            {
                return "1";
            }
        }

        private async Task SaveBomItemAsync(int exportRecordId, BomItem bomItem, ExportContext context)
        {
            try
            {
                var apiBomItem = new ApiBomItem
                {
                    ItemNo = bomItem.ItemNo,
                    PartNo = bomItem.PartNo,
                    SortOrder = bomItem.SortOrder,
                    Qty = bomItem.Qty,
                    TotalQty = bomItem.TotalQty,
                    Description = bomItem.Description,
                    PartName = bomItem.PartName,
                    ConfigurationName = bomItem.ConfigurationName,
                    Material = bomItem.Material
                };

                if (bomItem.CutTemplate != null)
                {
                    apiBomItem.CutTemplate = new ApiCutTemplate
                    {
                        DxfFilePath = bomItem.CutTemplate.DxfFilePath,
                        ContentHash = bomItem.CutTemplate.ContentHash,
                        Thickness = bomItem.CutTemplate.Thickness,
                        KFactor = bomItem.CutTemplate.KFactor,
                        DefaultBendRadius = bomItem.CutTemplate.DefaultBendRadius
                    };
                }

                await _apiClient.CreateBomItemAsync(exportRecordId, apiBomItem);
            }
            catch (Exception ex)
            {
                LogProgress(context, $"API error saving BOM item: {ex.Message}", LogLevel.Error);
            }
        }

        #endregion

        #region Helper Methods

        private string CreateTempWorkDir()
        {
            var path = Path.Combine(Path.GetTempPath(), "ExportDXF-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }

        private void CleanupTempDir(string tempDir)
        {
            try
            {
                if (Directory.Exists(tempDir))
                    Directory.Delete(tempDir, recursive: true);
            }
            catch
            {
                // Best-effort cleanup
            }
        }

        private string ParseDrawingNumber(ExportContext context)
        {
            // Use explicit Equipment/DrawingNo from the UI when available
            if (!string.IsNullOrWhiteSpace(context?.Equipment))
            {
                return !string.IsNullOrWhiteSpace(context?.DrawingNo)
                    ? $"{context.Equipment} {context.DrawingNo}"
                    : context.Equipment;
            }

            // Fallback: parse from prefix or document title
            var candidate = context?.FilePrefix;
            var info = string.IsNullOrWhiteSpace(candidate) ? null : DrawingInfo.Parse(candidate);
            if (info == null)
            {
                var title = context?.ActiveDocument?.Title;
                info = string.IsNullOrWhiteSpace(title) ? null : DrawingInfo.Parse(title);
            }
            return info?.ToString();
        }

        private void ValidateContext(ExportContext context)
        {
            if (context.ActiveDocument == null)
                throw new ArgumentException("ActiveDocument cannot be null.", nameof(context));

            if (context.ProgressCallback == null)
                throw new ArgumentException("ProgressCallback cannot be null.", nameof(context));
        }

        private void LogProgress(ExportContext context, string message, LogLevel level = LogLevel.Info, string file = null)
        {
            context.ProgressCallback?.Invoke(message, level, file);
        }

        #endregion
    }
}
