using ExportDXF.Data;
using ExportDXF.Extensions;
using ExportDXF.ItemExtractors;
using ExportDXF.Models;
using ExportDXF.Utilities;
using ExportDXF;
using Microsoft.EntityFrameworkCore;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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

            var drawingNumber = ParseDrawingNumber(context);
            var outputFolder = _fileExportService.GetDrawingOutputFolder(context.Equipment, context.DrawingNo);

            try
            {
                _solidWorksService.EnableUserControl(false);

                switch (context.ActiveDocument.DocumentType)
                {
                    case DocumentType.Part:
                        ExportPart(context, outputFolder, drawingNumber);
                        break;

                    case DocumentType.Assembly:
                        ExportAssembly(context, outputFolder, drawingNumber);
                        break;

                    case DocumentType.Drawing:
                        ExportDrawing(context, drawingNumber, outputFolder);
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

        private void ExportPart(ExportContext context, string outputFolder, string drawingNumber)
        {
            LogProgress(context, "Active document is a Part");

            var part = context.ActiveDocument.NativeDocument as PartDoc;
            if (part == null)
            {
                LogProgress(context, "Failed to get part document.", LogLevel.Error);
                return;
            }

            var exportRecord = CreateExportRecord(context, drawingNumber, outputFolder);
            var item = _partExporter.ExportSinglePart(part, outputFolder, context);

            if (item != null)
            {
                // Assign auto-incremented item number
                var nextItemNo = GetNextItemNumber(drawingNumber);
                item.ItemNo = nextItemNo;

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

                if (!string.IsNullOrEmpty(item.FileName))
                {
                    var dxfPath = Path.Combine(outputFolder, item.FileName + ".dxf");
                    bomItem.CutTemplate = new CutTemplate
                    {
                        DxfFilePath = dxfPath,
                        ContentHash = item.ContentHash,
                        Thickness = item.Thickness > 0 ? item.Thickness : null,
                        KFactor = item.KFactor > 0 ? item.KFactor : null,
                        DefaultBendRadius = item.BendRadius > 0 ? item.BendRadius : null
                    };

                    HandleDxfVersioning(item, dxfPath, context);
                }

                context.BomItemCallback?.Invoke(bomItem);

                if (exportRecord != null)
                    SaveBomItem(bomItem, context);
            }
        }

        private void ExportAssembly(ExportContext context, string outputFolder, string drawingNumber)
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

            var exportRecord = CreateExportRecord(context, drawingNumber, outputFolder);

            // Auto-assign item numbers for items that don't have one
            var nextNum = int.Parse(GetNextItemNumber(drawingNumber));
            foreach (var item in items)
            {
                if (string.IsNullOrWhiteSpace(item.ItemNo))
                {
                    item.ItemNo = nextNum.ToString();
                    nextNum++;
                }
            }

            ExportItems(items, outputFolder, context, exportRecord?.Id);
        }

        private void ExportDrawing(ExportContext context, string drawingNumber, string drawingOutputFolder)
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

            // Export drawing to PDF
            var tempDir = CreateTempWorkDir();
            _drawingExporter.ExportToPdf(drawing, tempDir, context);

            // Copy PDF to output folder with versioning
            string pdfStashPath = null;
            string savedPdfPath = null;
            try
            {
                var pdfs = Directory.GetFiles(tempDir, "*.pdf");
                if (pdfs.Length > 0)
                {
                    // Determine the destination path to stash the existing file
                    var pdfFileName = !string.IsNullOrEmpty(drawingNumber)
                        ? $"{drawingNumber}.pdf"
                        : Path.GetFileName(pdfs[0]);
                    var pdfDestPath = Path.Combine(drawingOutputFolder, pdfFileName);

                    pdfStashPath = _fileExportService.StashFile(pdfDestPath);
                    savedPdfPath = _fileExportService.SavePdfFile(pdfs[0], drawingNumber, drawingOutputFolder);
                }
            }
            catch (Exception ex)
            {
                LogProgress(context, $"PDF save error: {ex.Message}", LogLevel.Error);
            }

            // Create export record in database
            var exportRecord = CreateExportRecord(context, drawingNumber, drawingOutputFolder);

            // Handle PDF versioning and update export record with hash
            if (exportRecord != null && savedPdfPath != null)
            {
                try
                {
                    HandlePdfVersioning(savedPdfPath, exportRecord.DrawingNumber, exportRecord, context);

                    // Archive or discard old PDF based on hash comparison
                    if (pdfStashPath != null)
                    {
                        using (var db = _dbContextFactory())
                        {
                            var previousRecord = db.ExportRecords
                                .Where(r => r.DrawingNumber == exportRecord.DrawingNumber
                                         && r.PdfContentHash != null
                                         && r.Id != exportRecord.Id)
                                .OrderByDescending(r => r.Id)
                                .FirstOrDefault();

                            if (previousRecord != null && previousRecord.PdfContentHash == exportRecord.PdfContentHash)
                            {
                                _fileExportService.DiscardStash(pdfStashPath);
                            }
                            else
                            {
                                _fileExportService.ArchiveFile(pdfStashPath, savedPdfPath);
                            }
                        }
                    }

                    // Update the record with the PDF hash
                    using (var db = _dbContextFactory())
                    {
                        db.ExportRecords.Attach(exportRecord);
                        db.Entry(exportRecord).Property(r => r.PdfContentHash).IsModified = true;
                        db.SaveChanges();
                    }
                }
                catch (Exception ex)
                {
                    _fileExportService.DiscardStash(pdfStashPath);
                    LogProgress(context, $"PDF versioning error: {ex.Message}", LogLevel.Error);
                }
            }
            else if (pdfStashPath != null)
            {
                // No export record - discard stash
                _fileExportService.DiscardStash(pdfStashPath);
            }

            // Export parts to DXF and save BOM items
            ExportItems(items, drawingOutputFolder, context, exportRecord?.Id);
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
                    _partExporter.ExportItem(item, saveDirectory, context);

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

                    // Only create CutTemplate if DXF was exported successfully
                    if (!string.IsNullOrEmpty(item.FileName))
                    {
                        successCount++;

                        var dxfPath = Path.Combine(saveDirectory, item.FileName + ".dxf");
                        bomItem.CutTemplate = new CutTemplate
                        {
                            DxfFilePath = dxfPath,
                            ContentHash = item.ContentHash,
                            Thickness = item.Thickness > 0 ? item.Thickness : null,
                            KFactor = item.KFactor > 0 ? item.KFactor : null,
                            DefaultBendRadius = item.BendRadius > 0 ? item.BendRadius : null
                        };

                        // Compare hash with previous export to decide archive/discard
                        HandleDxfVersioning(item, dxfPath, context);
                    }
                    else
                    {
                        skippedCount++;
                    }

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
                LogProgress(context, $"BOM items saved to database (ExportRecord ID: {exportRecordId.Value})", LogLevel.Info);
            }
        }

        #endregion

        #region Versioning

        private void HandleDxfVersioning(Item item, string dxfPath, ExportContext context)
        {
            if (string.IsNullOrEmpty(item.ContentHash))
                return;

            try
            {
                using (var db = _dbContextFactory())
                {
                    var previousCutTemplate = db.CutTemplates
                        .Where(ct => ct.DxfFilePath == dxfPath && ct.ContentHash != null)
                        .OrderByDescending(ct => ct.Id)
                        .FirstOrDefault();

                    if (previousCutTemplate != null && previousCutTemplate.ContentHash == item.ContentHash)
                    {
                        // Content unchanged - discard the stashed file
                        _fileExportService.DiscardStash(item.StashedFilePath);
                        LogProgress(context, $"DXF unchanged: {item.FileName}.dxf", LogLevel.Info);
                    }
                    else
                    {
                        // Content changed or first export - archive the old file
                        if (!string.IsNullOrEmpty(item.StashedFilePath))
                        {
                            _fileExportService.ArchiveFile(item.StashedFilePath, dxfPath);
                            LogProgress(context, $"DXF updated, previous version archived: {item.FileName}.dxf", LogLevel.Info);
                        }
                        else
                        {
                            LogProgress(context, $"Exported: {item.FileName}.dxf", LogLevel.Info);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Don't fail the export if versioning fails - just discard the stash
                _fileExportService.DiscardStash(item.StashedFilePath);
                LogProgress(context, $"Versioning check failed for {item.FileName}: {ex.Message}", LogLevel.Warning);
            }
        }

        private void HandlePdfVersioning(string pdfPath, string drawingNumber, ExportRecord exportRecord, ExportContext context)
        {
            try
            {
                var newHash = ContentHasher.ComputeFileHash(pdfPath);

                using (var db = _dbContextFactory())
                {
                    var previousRecord = db.ExportRecords
                        .Where(r => r.DrawingNumber == drawingNumber && r.PdfContentHash != null)
                        .OrderByDescending(r => r.Id)
                        .FirstOrDefault();

                    if (previousRecord != null && previousRecord.PdfContentHash == newHash)
                    {
                        LogProgress(context, $"PDF unchanged: {Path.GetFileName(pdfPath)}", LogLevel.Info);
                    }
                    else
                    {
                        LogProgress(context, $"Saved PDF: {Path.GetFileName(pdfPath)}", LogLevel.Info);
                    }
                }

                if (exportRecord != null)
                    exportRecord.PdfContentHash = newHash;
            }
            catch (Exception ex)
            {
                LogProgress(context, $"PDF versioning check failed: {ex.Message}", LogLevel.Warning);
            }
        }

        #endregion

        #region Database Helpers

        private ExportRecord CreateExportRecord(ExportContext context, string drawingNumber, string outputFolder)
        {
            try
            {
                using (var db = _dbContextFactory())
                {
                    db.Database.Migrate();
                    var record = new ExportRecord
                    {
                        DrawingNumber = drawingNumber ?? context.ActiveDocument.Title,
                        SourceFilePath = context.ActiveDocument.FilePath,
                        OutputFolder = outputFolder,
                        ExportedAt = DateTime.Now,
                        ExportedBy = System.Environment.UserName
                    };

                    db.ExportRecords.Add(record);
                    db.SaveChanges();
                    LogProgress(context, $"Created export record (ID: {record.Id})", LogLevel.Info);
                    return record;
                }
            }
            catch (Exception ex)
            {
                LogProgress(context, $"Database error creating export record: {ex.Message}", LogLevel.Error);
                return null;
            }
        }

        private string GetNextItemNumber(string drawingNumber)
        {
            if (string.IsNullOrEmpty(drawingNumber))
                return "1";

            try
            {
                using (var db = _dbContextFactory())
                {
                    var existingItems = db.ExportRecords
                        .Where(r => r.DrawingNumber == drawingNumber)
                        .SelectMany(r => r.BomItems)
                        .Select(b => b.ItemNo)
                        .ToList();

                    int maxNum = 0;
                    foreach (var itemNo in existingItems)
                    {
                        if (int.TryParse(itemNo, out var num) && num > maxNum)
                            maxNum = num;
                    }

                    return (maxNum + 1).ToString();
                }
            }
            catch
            {
                return "1";
            }
        }

        private void SaveBomItem(BomItem bomItem, ExportContext context)
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
            // Use explicit Equipment/DrawingNo from the UI when available
            if (!string.IsNullOrWhiteSpace(context?.Equipment) && !string.IsNullOrWhiteSpace(context?.DrawingNo))
                return $"{context.Equipment} {context.DrawingNo}";

            // Fallback: parse from prefix or document title
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
