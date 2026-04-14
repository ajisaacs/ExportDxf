using ExportDXF.Models;
using ExportDXF.ViewFlipDeciders;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExportDXF.Services
{
    /// <summary>
    /// Context object containing all information needed for an export operation.
    /// </summary>
    public class ExportContext
    {
        private const string DRAWING_TEMPLATE_FOLDER = "Templates";
        private const string DRAWING_TEMPLATE_FILE = "Blank.drwdot";

        /// <summary>
        /// The document to be exported.
        /// </summary>
        public SolidWorksDocument ActiveDocument { get; set; }

        /// <summary>
        /// The view flip decider to determine if views should be flipped.
        /// </summary>
        public IViewFlipDecider ViewFlipDecider { get; set; }

        /// <summary>
        /// Filename template with placeholders (e.g., "4321 A01 PT{item_no:2}").
        /// </summary>
        public string FilenameTemplate { get; set; }

        /// <summary>
        /// Output folder for DXF files and Excel workbook.
        /// </summary>
        public string OutputFolder { get; set; }

        /// <summary>
        /// Cancellation token for canceling the export operation.
        /// </summary>
        public CancellationToken CancellationToken { get; set; }

        /// <summary>
        /// Callback for reporting progress and status messages.
        /// Parameters: message, level, file (optional)
        /// </summary>
        public Action<string, LogLevel, string> ProgressCallback { get; set; }

        /// <summary>
        /// Callback for adding BOM items to the UI.
        /// </summary>
        public Action<BomItem> BomItemCallback { get; set; }

        public void LogProgress(string message, LogLevel level = LogLevel.Info, string file = null)
        {
            ProgressCallback?.Invoke(message, level, file);
        }

        public SldWorks SolidWorksApp { get; set; }

        public DrawingDoc TemplateDrawing { get; set; }

        private string DrawingTemplatePath
        {
            get
            {
                return Path.Combine(
                    Application.StartupPath,
                    DRAWING_TEMPLATE_FOLDER,
                    DRAWING_TEMPLATE_FILE);
            }
        }

        public DrawingDoc GetOrCreateTemplateDrawing()
        {
            if (TemplateDrawing != null)
                return TemplateDrawing;

            TemplateDrawing = SolidWorksApp.NewDocument(
                DrawingTemplatePath,
                (int)swDwgPaperSizes_e.swDwgPaperDsize,
                1,
                1) as DrawingDoc;

            return TemplateDrawing;
        }

        public void CleanupTemplateDrawing()
        {
            try
            {
                if (TemplateDrawing == null)
                    return;

                if (SolidWorksApp == null)
                {
                    ProgressCallback?.Invoke("Warning: Cannot cleanup template drawing - SolidWorks app not available", LogLevel.Warning, null);
                    TemplateDrawing = null;
                    return;
                }

                var model = TemplateDrawing as ModelDoc2;
                if (model != null)
                {
                    var title = model.GetTitle();

                    if (!string.IsNullOrEmpty(title))
                    {
                        SolidWorksApp.CloseDoc(title);
                        ProgressCallback?.Invoke("Closed template drawing", LogLevel.Info, null);
                    }
                }

                TemplateDrawing = null;
            }
            catch (Exception ex)
            {
                ProgressCallback?.Invoke($"Failed to close template drawing: {ex.Message}", LogLevel.Error, null);
                TemplateDrawing = null;
            }
        }

        public void CloseDocument(string title)
        {
            SolidWorksApp?.CloseDoc(title);
        }

        public ModelDoc2 CreateDocument(string templatePath, int paperSize, double width, double height)
        {
            return SolidWorksApp?.NewDocument(templatePath, paperSize, width, height) as ModelDoc2;
        }
    }
}
