using ExportDXF.Extensions;
using ExportDXF.ItemExtractors;
using ExportDXF.Models;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

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
        private const string DRAWING_TEMPLATE_FOLDER = "Templates";
        private const string DRAWING_TEMPLATE_FILE = "Blank.drwdot";

        private readonly ISolidWorksService _solidWorksService;
        private readonly IBomExtractor _bomExtractor;
        private readonly IPartExporter _partExporter;
        private readonly IDrawingExporter _drawingExporter;
        private readonly IBomExcelExporter _bomExcelExporter;

        public DxfExportService(
            ISolidWorksService solidWorksService,
            IBomExtractor bomExtractor,
            IPartExporter partExporter,
            IDrawingExporter drawingExporter,
            IBomExcelExporter bomExcelExporter)
        {
            _solidWorksService = solidWorksService ?? throw new ArgumentNullException(nameof(solidWorksService));
            _bomExtractor = bomExtractor ?? throw new ArgumentNullException(nameof(bomExtractor));
            _partExporter = partExporter ?? throw new ArgumentNullException(nameof(partExporter));
            _drawingExporter = drawingExporter ?? throw new ArgumentNullException(nameof(drawingExporter));
            _bomExcelExporter = bomExcelExporter ?? throw new ArgumentNullException(nameof(bomExcelExporter));
        }

        /// <summary>
        /// Exports the document specified in the context to DXF format.
        /// </summary>
        public void Export(ExportContext context)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            ValidateContext(context);

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

            var saveDirectory = PromptUserForDirectory(context);
            if (saveDirectory == null)
            {
                LogProgress(context, "Canceled", Color.Red);
                return;
            }

            _partExporter.ExportSinglePart(part, saveDirectory, context);
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
                LogProgress(context, "No items found in assembly.", Color.Yellow);
                return;
            }

            LogProgress(context, $"Found {items.Count} item(s).", null);

            var saveDirectory = PromptUserForDirectory(context);
            if (saveDirectory == null)
            {
                LogProgress(context, "Canceled", Color.Red);
                return;
            }

            ExportItems(items, saveDirectory, context);
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

            var saveDirectory = PromptUserForDirectory(context);
            if (saveDirectory == null)
            {
                LogProgress(context, "Canceled", Color.Red);
                return;
            }

            // Export drawing to PDF first
            _drawingExporter.ExportToPdf(drawing, saveDirectory, context.ProgressCallback);

            // Then export parts to DXF
            ExportItems(items, saveDirectory, context);
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

        private void ExportItems(List<Item> items, string saveDirectory, ExportContext context)
        {
            DrawingDoc templateDrawing = null;

            try
            {
                templateDrawing = CreateTemplateDrawing();

                LogProgress(context, "", null);

                int successCount = 0;
                int failureCount = 0;

                foreach (var item in items)
                {
                    if (context.CancellationToken.IsCancellationRequested)
                    {
                        LogProgress(context, "Export canceled by user.", Color.Yellow);
                        return;
                    }

                    try
                    {
                        _partExporter.ExportItem(item, templateDrawing, saveDirectory, context);

                        if (!string.IsNullOrEmpty(item.FileName))
                            successCount++;
                        else
                            failureCount++;
                    }
                    catch (Exception ex)
                    {
                        LogProgress(context, $"Error exporting item {item.ItemNo}: {ex.Message}", Color.Red);
                        failureCount++;
                    }

                    LogProgress(context, "", null);
                }

                LogProgress(context, $"Export complete: {successCount} succeeded, {failureCount} failed",
                    failureCount > 0 ? Color.Yellow : Color.Green);

                // Create BOM Excel file
                CreateBomExcelFile(items, saveDirectory, context);
            }
            finally
            {
                CloseTemplateDrawing(templateDrawing);
            }
        }

        #endregion

        #region Template Drawing Management

        private DrawingDoc CreateTemplateDrawing()
        {
            var sldWorks = _solidWorksService.GetNativeSldWorks();

            if (sldWorks == null)
                throw new InvalidOperationException("SolidWorks service is not connected.");

            var templatePath = GetDrawingTemplatePath();

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException(
                    $"Drawing template not found at: {templatePath}",
                    templatePath);
            }

            var drawing = sldWorks.NewDocument(
                templatePath,
                (int)swDwgPaperSizes_e.swDwgPaperDsize,
                1,
                1) as DrawingDoc;

            if (drawing == null)
            {
                throw new InvalidOperationException(
                    "Failed to create drawing from template. Please check the template file.");
            }

            return drawing;
        }

        private void CloseTemplateDrawing(DrawingDoc drawing)
        {
            if (drawing == null)
                return;

            try
            {
                var model = drawing as ModelDoc2;
                var title = model?.GetTitle();

                if (!string.IsNullOrEmpty(title))
                {
                    _solidWorksService.CloseDocument(title);
                }
            }
            catch
            {
                // Ignore errors during cleanup
            }
        }

        private string GetDrawingTemplatePath()
        {
            return Path.Combine(
                Application.StartupPath,
                DRAWING_TEMPLATE_FOLDER,
                DRAWING_TEMPLATE_FILE);
        }

        #endregion

        #region BOM Excel Creation

        private void CreateBomExcelFile(List<Item> items, string saveDirectory, ExportContext context)
        {
            try
            {
                var bomFileName = GetBomFileName(context.FilePrefix);
                var bomFilePath = Path.Combine(saveDirectory, bomFileName + ".xlsx");

                _bomExcelExporter.CreateBOMExcelFile(bomFilePath, items);

                LogProgress(context, $"Created BOM Excel file: {bomFileName}.xlsx", Color.Green);
            }
            catch (Exception ex)
            {
                LogProgress(context, $"Failed to create BOM Excel: {ex.Message}", Color.Red);
            }
        }

        private string GetBomFileName(string prefix)
        {
            if (string.IsNullOrWhiteSpace(prefix))
                return "BOM";

            var drawingInfo = DrawingInfo.Parse(prefix);

            if (drawingInfo != null)
            {
                return $"{drawingInfo.JobNo} {drawingInfo.DrawingNo} BOM";
            }

            return prefix.Trim() + " BOM";
        }

        #endregion

        #region User Interaction

        private string PromptUserForDirectory(ExportContext context)
        {
            // Check if already canceled
            if (context.CancellationToken.IsCancellationRequested)
                return null;

            string selectedPath = null;

            // Must run on STA thread for FolderBrowserDialog
            var thread = new System.Threading.Thread(() =>
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "Where do you want to save the DXF files?";
                    dialog.ShowNewFolderButton = true;

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        selectedPath = dialog.SelectedPath;
                    }
                }
            });

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();

            return selectedPath;
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
