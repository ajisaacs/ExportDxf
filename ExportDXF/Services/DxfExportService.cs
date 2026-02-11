using ExportDXF.Data;
using ExportDXF.Extensions;
using ExportDXF.ItemExtractors;
using ExportDXF.Models;
using ExportDXF;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.IO;

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
        private readonly IFileExportService _fileExportService;
        private readonly Func<ExportDxfDbContext> _dbContextFactory;

        public DxfExportService(
            ISolidWorksService solidWorksService,
            IBomExtractor bomExtractor,
            IPartExporter partExporter,
            IDrawingExporter drawingExporter,
            IFileExportService fileExportService,
            Func<ExportDxfDbContext> dbContextFactory = null)
        {
            _solidWorksService = solidWorksService ?? throw new ArgumentNullException(nameof(solidWorksService));
            _bomExtractor = bomExtractor ?? throw new ArgumentNullException(nameof(bomExtractor));
            _partExporter = partExporter ?? throw new ArgumentNullException(nameof(partExporter));
            _drawingExporter = drawingExporter ?? throw new ArgumentNullException(nameof(drawingExporter));
            _fileExportService = fileExportService ?? throw new ArgumentNullException(nameof(fileExportService));
            _dbContextFactory = dbContextFactory ?? (() => new ExportDxfDbContext());
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
                        LogProgress(context, "Unknown document type.", LogLevel.Error);
                        break;
                }
            }
            finally
            {
                CleanupExportContext(context);
                _solidWorksService.EnableUserControl(true);

                var duration = DateTime.Now - startTime;
                LogProgress(context, $"Run time: {duration.ToReadableFormat()}");
            }
        }

        #region Export Methods by Document Type

        private void ExportPart(ExportContext context)
        {
            LogProgress(context, "Active document is a Part");

            var part = context.ActiveDocument.NativeDocument as PartDoc;
            if (part == null)
            {
                LogProgress(context, "Failed to get part document.", LogLevel.Error);
                return;
            }

            // Export directly to the output folder
            _partExporter.ExportSinglePart(part, _fileExportService.OutputFolder, context);
        }

        private void ExportAssembly(ExportContext context)
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

            // Export directly to the output folder
            ExportItems(items, _fileExportService.OutputFolder, context);
        }

        private void ExportDrawing(ExportContext context)
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

            // Determine drawing number for file naming
            var drawingNumber = ParseDrawingNumber(context);

            // Export drawing to PDF
            var tempDir = CreateTempWorkDir();
            _drawingExporter.ExportToPdf(drawing, tempDir, context);

            // Copy PDF to output folder
            try
            {
                var pdfs = Directory.GetFiles(tempDir, "*.pdf");
                if (pdfs.Length > 0)
                {
                    var savedPath = _fileExportService.SavePdfFile(pdfs[0], drawingNumber);
                    LogProgress(context, $"Saved PDF: {Path.GetFileName(savedPath)}", LogLevel.Info);
                }
            }
            catch (Exception ex)
            {
                LogProgress(context, $"PDF save error: {ex.Message}", LogLevel.Error);
            }

            // Create export record in database
            ExportRecord exportRecord = null;
            try
            {
                using (var db = _dbContextFactory())
                {
                    db.Database.EnsureCreated();
                    exportRecord = new ExportRecord
                    {
                        DrawingNumber = drawingNumber ?? context.ActiveDocument.Title,
                        SourceFilePath = context.ActiveDocument.FilePath,
                        OutputFolder = _fileExportService.OutputFolder,
                        ExportedAt = DateTime.Now,
                        ExportedBy = System.Environment.UserName
                    };
                    db.ExportRecords.Add(exportRecord);
                    db.SaveChanges();
                    LogProgress(context, $"Created export record (ID: {exportRecord.Id})", LogLevel.Info);
                }
            }
            catch (Exception ex)
            {
                LogProgress(context, $"Database error creating export record: {ex.Message}", LogLevel.Error);
            }

            // Export parts to DXF (directly to output folder) and save BOM items
            ExportItems(items, _fileExportService.OutputFolder, context, exportRecord?.Id);
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

        private void ExportItems(List<Item> items, string saveDirectory, ExportContext context, int? exportRecordId = null)
        {
            int successCount = 0;
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
                    _partExporter.ExportItem(item, saveDirectory, context);

                    if (!string.IsNullOrEmpty(item.FileName))
                    {
                        successCount++;
                        LogProgress(context, $"Exported: {item.FileName}.dxf", LogLevel.Info);

                        // Create BOM item
                        var dxfPath = Path.Combine(saveDirectory, item.FileName + ".dxf");
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
                            Material = item.Material ?? "",
                            Thickness = item.Thickness > 0 ? item.Thickness : null,
                            KFactor = item.KFactor > 0 ? item.KFactor : null,
                            DefaultBendRadius = item.BendRadius > 0 ? item.BendRadius : null,
                            DxfFilePath = dxfPath
                        };

                        // Add to UI
                        context.BomItemCallback?.Invoke(bomItem);

                        // Save BOM item to database if we have an export record
                        if (exportRecordId.HasValue)
                        {
                            try
                            {
                                using (var db = _dbContextFactory())
                                {
                                    db.BomItems.Add(bomItem);
                                    db.SaveChanges();
                                }
                            }
                            catch (Exception dbEx)
                            {
                                LogProgress(context, $"Database error saving BOM item: {dbEx.Message}", LogLevel.Error);
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
                    LogProgress(context, $"Error exporting item {item.ItemNo}: {ex.Message}", LogLevel.Error);
                    failureCount++;
                }
            }

            LogProgress(context, $"Export complete: {successCount} succeeded, {failureCount} failed",
                failureCount > 0 ? LogLevel.Warning : LogLevel.Info);

            if (exportRecordId.HasValue)
            {
                LogProgress(context, $"BOM items saved to database (ExportRecord ID: {exportRecordId.Value})", LogLevel.Info);
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
