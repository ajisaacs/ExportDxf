using ExportDXF.Extensions;
using ExportDXF.ItemExtractors;
using ExportDXF.Models;
using ExportDXF.Utilities;
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
        Task ExportAsync(ExportContext context);
    }

    public class DxfExportService : IDxfExportService
    {
        private readonly ISolidWorksService _solidWorksService;
        private readonly IBomExtractor _bomExtractor;
        private readonly IPartExporter _partExporter;
        private readonly IDrawingExporter _drawingExporter;
        private readonly ExcelExportService _excelExportService;
        private readonly LogFileService _logFileService;

        public DxfExportService(
            ISolidWorksService solidWorksService,
            IBomExtractor bomExtractor,
            IPartExporter partExporter,
            IDrawingExporter drawingExporter,
            ExcelExportService excelExportService,
            LogFileService logFileService)
        {
            _solidWorksService = solidWorksService ?? throw new ArgumentNullException(nameof(solidWorksService));
            _bomExtractor = bomExtractor ?? throw new ArgumentNullException(nameof(bomExtractor));
            _partExporter = partExporter ?? throw new ArgumentNullException(nameof(partExporter));
            _drawingExporter = drawingExporter ?? throw new ArgumentNullException(nameof(drawingExporter));
            _excelExportService = excelExportService ?? throw new ArgumentNullException(nameof(excelExportService));
            _logFileService = logFileService ?? throw new ArgumentNullException(nameof(logFileService));
        }

        public async Task ExportAsync(ExportContext context)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            ValidateContext(context);
            SetupExportContext(context);

            var outputFolder = context.OutputFolder;
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var prefix = FilenameTemplateParser.GetPrefix(
                context.FilenameTemplate,
                context.ActiveDocument.Title);

            var xlsxPath = Path.Combine(outputFolder, $"{prefix}.xlsx");
            var logPath = Path.Combine(outputFolder, $"{prefix}.log");

            _logFileService.StartExportLog(logPath);
            _logFileService.LogInfo($"Export started: {context.ActiveDocument.FilePath}");
            _logFileService.LogInfo($"Template: {context.FilenameTemplate}");
            _logFileService.LogInfo($"Output: {outputFolder}");

            // Read existing cut templates for revision comparison
            var existingTemplates = _excelExportService.ReadExistingCutTemplates(xlsxPath);

            var bomItems = new List<BomItem>();
            List<Dictionary<string, string>> rawBomTable = null;

            var startTime = DateTime.Now;
            var tempDir = CreateTempWorkDir();

            try
            {
                _solidWorksService.EnableUserControl(false);

                switch (context.ActiveDocument.DocumentType)
                {
                    case DocumentType.Part:
                        await ExportPartAsync(context, tempDir, outputFolder, existingTemplates, bomItems);
                        break;

                    case DocumentType.Assembly:
                        await ExportAssemblyAsync(context, tempDir, outputFolder, existingTemplates, bomItems);
                        break;

                    case DocumentType.Drawing:
                        rawBomTable = await ExportDrawingAsync(context, tempDir, outputFolder, existingTemplates, bomItems);
                        break;

                    default:
                        LogProgress(context, "Unknown document type.", LogLevel.Error);
                        break;
                }

                // Write Excel file
                _excelExportService.Write(xlsxPath, rawBomTable, bomItems);
                _logFileService.LogInfo($"Wrote {Path.GetFileName(xlsxPath)}");
                LogProgress(context, $"Saved {Path.GetFileName(xlsxPath)}");
            }
            catch (OperationCanceledException)
            {
                _logFileService.LogWarning("Export cancelled by user");
                throw;
            }
            catch (Exception ex)
            {
                _logFileService.LogError($"Export failed: {ex.Message}");
                throw;
            }
            finally
            {
                CleanupExportContext(context);
                _solidWorksService.EnableUserControl(true);
                CleanupTempDir(tempDir);

                var duration = DateTime.Now - startTime;
                _logFileService.LogInfo($"Run time: {duration.ToReadableFormat()}");
                LogProgress(context, $"Run time: {duration.ToReadableFormat()}");
            }
        }

        #region Export Methods by Document Type

        private async Task ExportPartAsync(
            ExportContext context,
            string tempDir,
            string outputFolder,
            Dictionary<string, (string ContentHash, int Revision, string FileName)> existingTemplates,
            List<BomItem> bomItems)
        {
            LogProgress(context, "Active document is a Part");

            var part = context.ActiveDocument.NativeDocument as PartDoc;
            if (part == null)
            {
                LogProgress(context, "Failed to get part document.", LogLevel.Error);
                return;
            }

            var item = _partExporter.ExportSinglePart(part, tempDir, context);

            if (item != null)
            {
                item.ItemNo = "1";

                var bomItem = CreateBomItem(item);
                PlaceDxfFile(item, context, outputFolder, existingTemplates, bomItem);
                bomItems.Add(bomItem);
                context.BomItemCallback?.Invoke(bomItem);
            }
        }

        private async Task ExportAssemblyAsync(
            ExportContext context,
            string tempDir,
            string outputFolder,
            Dictionary<string, (string ContentHash, int Revision, string FileName)> existingTemplates,
            List<BomItem> bomItems)
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

            // Assign item numbers
            int nextNum = 1;
            foreach (var item in items)
            {
                if (string.IsNullOrWhiteSpace(item.ItemNo))
                {
                    item.ItemNo = nextNum.ToString();
                    nextNum++;
                }
            }

            await ExportItemsAsync(items, tempDir, outputFolder, context, existingTemplates, bomItems);
        }

        private async Task<List<Dictionary<string, string>>> ExportDrawingAsync(
            ExportContext context,
            string tempDir,
            string outputFolder,
            Dictionary<string, (string ContentHash, int Revision, string FileName)> existingTemplates,
            List<BomItem> bomItems)
        {
            LogProgress(context, "Active document is a Drawing");
            LogProgress(context, "Finding BOM tables...");

            var drawing = context.ActiveDocument.NativeDocument as DrawingDoc;
            if (drawing == null)
            {
                LogProgress(context, "Failed to get drawing document.", LogLevel.Error);
                return null;
            }

            // Read raw BOM table for Excel output
            var rawBomTable = new List<Dictionary<string, string>>();
            var bomTables = drawing.GetBomTables();
            foreach (var table in bomTables)
            {
                var rows = RawBomTableReader.Read(table);
                rawBomTable.AddRange(rows);
            }

            // Extract items for DXF export
            var items = _bomExtractor.ExtractFromDrawing(drawing, context.ProgressCallback);

            if (items == null || items.Count == 0)
            {
                LogProgress(context, "Error: Bill of materials not found.", LogLevel.Error);
                return rawBomTable;
            }

            LogProgress(context, $"Found {items.Count} component(s)");

            // Export drawing to PDF in output folder
            _drawingExporter.ExportToPdf(drawing, outputFolder, context);

            await ExportItemsAsync(items, tempDir, outputFolder, context, existingTemplates, bomItems);

            return rawBomTable;
        }

        #endregion

        #region Context Management

        private void SetupExportContext(ExportContext context)
        {
            context.SolidWorksApp = _solidWorksService.GetNativeSldWorks();

            if (context.SolidWorksApp == null)
                throw new InvalidOperationException("SolidWorks service is not connected.");

            context.TemplateDrawing = null;
            LogProgress(context, "Export context initialized");
        }

        private void CleanupExportContext(ExportContext context)
        {
            try
            {
                context.CleanupTemplateDrawing();
            }
            catch (Exception ex)
            {
                LogProgress(context, $"Warning: Failed to cleanup template drawing: {ex.Message}", LogLevel.Warning);
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

        private async Task ExportItemsAsync(
            List<Item> items,
            string tempDir,
            string outputFolder,
            ExportContext context,
            Dictionary<string, (string ContentHash, int Revision, string FileName)> existingTemplates,
            List<BomItem> bomItems)
        {
            int successCount = 0;
            int skippedCount = 0;
            int unchangedCount = 0;
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
                    _partExporter.ExportItem(item, tempDir, context);

                    var bomItem = CreateBomItem(item);
                    bomItem.SortOrder = sortOrder++;

                    if (!string.IsNullOrEmpty(item.LocalTempPath))
                    {
                        var wasPlaced = PlaceDxfFile(item, context, outputFolder, existingTemplates, bomItem);
                        if (wasPlaced)
                            successCount++;
                        else
                            unchangedCount++;
                    }
                    else
                    {
                        skippedCount++;
                    }

                    bomItems.Add(bomItem);
                    context.BomItemCallback?.Invoke(bomItem);
                }
                catch (Exception ex)
                {
                    LogProgress(context, $"Error exporting item {item.ItemNo}: {ex.Message}", LogLevel.Error);
                    _logFileService.LogError($"Item {item.ItemNo}: {ex.Message}");
                    failureCount++;
                }
            }

            var summary = $"Export complete: {successCount} exported";
            if (unchangedCount > 0)
                summary += $", {unchangedCount} unchanged";
            if (skippedCount > 0)
                summary += $", {skippedCount} skipped (non-sheet-metal)";
            if (failureCount > 0)
                summary += $", {failureCount} failed";

            LogProgress(context, summary, failureCount > 0 ? LogLevel.Warning : LogLevel.Info);
            _logFileService.LogInfo(summary);
        }

        /// <summary>
        /// Places a DXF file in the output folder with revision tracking.
        /// Returns true if a new file was written, false if unchanged.
        /// </summary>
        private bool PlaceDxfFile(
            Item item,
            ExportContext context,
            string outputFolder,
            Dictionary<string, (string ContentHash, int Revision, string FileName)> existingTemplates,
            BomItem bomItem)
        {
            var baseName = FilenameTemplateParser.Evaluate(context.FilenameTemplate, item);
            var contentHash = item.ContentHash;

            int revision = 1;
            string dxfFileName;

            // Check existing templates for revision comparison
            if (existingTemplates.TryGetValue(item.ItemNo, out var existing))
            {
                if (existing.ContentHash == contentHash)
                {
                    // Unchanged — skip file write, keep existing
                    dxfFileName = existing.FileName;
                    revision = existing.Revision;

                    LogProgress(context, $"Unchanged: {dxfFileName}", LogLevel.Info, item.PartName);
                    _logFileService.LogInfo($"Unchanged: {dxfFileName}");

                    bomItem.CutTemplate = new CutTemplate
                    {
                        DxfFilePath = dxfFileName,
                        ContentHash = contentHash,
                        Revision = revision,
                        Thickness = item.Thickness > 0 ? item.Thickness : (double?)null,
                        KFactor = item.KFactor > 0 ? item.KFactor : (double?)null,
                        DefaultBendRadius = item.BendRadius > 0 ? item.BendRadius : (double?)null
                    };

                    return false;
                }
                else
                {
                    // Changed — increment revision
                    revision = existing.Revision + 1;
                    dxfFileName = GetRevisionFileName(baseName, revision);

                    LogProgress(context, $"Updated: {dxfFileName} (Rev{revision})", LogLevel.Info, item.PartName);
                    _logFileService.LogInfo($"Updated: {dxfFileName} (was {existing.FileName})");
                }
            }
            else
            {
                // New item
                dxfFileName = $"{baseName}.dxf";
                LogProgress(context, $"Exported: {dxfFileName}", LogLevel.Info, item.PartName);
                _logFileService.LogInfo($"Exported: {dxfFileName}");
            }

            // Copy from temp to output
            var destPath = Path.Combine(outputFolder, dxfFileName);
            File.Copy(item.LocalTempPath, destPath, overwrite: true);

            bomItem.CutTemplate = new CutTemplate
            {
                DxfFilePath = Path.GetFileNameWithoutExtension(dxfFileName),
                ContentHash = contentHash,
                Revision = revision,
                Thickness = item.Thickness > 0 ? item.Thickness : (double?)null,
                KFactor = item.KFactor > 0 ? item.KFactor : (double?)null,
                DefaultBendRadius = item.BendRadius > 0 ? item.BendRadius : (double?)null
            };

            return true;
        }

        private BomItem CreateBomItem(Item item)
        {
            return new BomItem
            {
                ItemNo = item.ItemNo ?? "",
                PartNo = item.FileName ?? item.PartName ?? "",
                SortOrder = 0,
                Qty = item.Quantity,
                TotalQty = item.Quantity,
                Description = item.Description ?? "",
                PartName = item.PartName ?? "",
                ConfigurationName = item.Configuration ?? "",
                Material = item.Material ?? ""
            };
        }

        #endregion

        #region Helper Methods

        private string GetRevisionFileName(string baseName, int revision)
        {
            if (revision <= 1)
                return $"{baseName}.dxf";
            return $"{baseName} Rev{revision}.dxf";
        }

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

        private void ValidateContext(ExportContext context)
        {
            if (context.ActiveDocument == null)
                throw new ArgumentException("ActiveDocument cannot be null.", nameof(context));

            if (context.ProgressCallback == null)
                throw new ArgumentException("ProgressCallback cannot be null.", nameof(context));

            if (string.IsNullOrWhiteSpace(context.FilenameTemplate))
                throw new ArgumentException("FilenameTemplate cannot be null or empty.", nameof(context));

            if (string.IsNullOrWhiteSpace(context.OutputFolder))
                throw new ArgumentException("OutputFolder cannot be null or empty.", nameof(context));
        }

        private void LogProgress(ExportContext context, string message, LogLevel level = LogLevel.Info, string file = null)
        {
            context.ProgressCallback?.Invoke(message, level, file);
        }

        #endregion
    }
}
