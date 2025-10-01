using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Drawing;
using System.IO;

namespace ExportDXF.Services
{
    /// <summary>
    /// Service for exporting drawing documents to PDF format.
    /// </summary>
    public interface IDrawingExporter
    {
        /// <summary>
        /// Exports a drawing document to PDF format.
        /// </summary>
        /// <param name="drawing">The drawing document to export.</param>
        /// <param name="saveDirectory">The directory where the PDF file will be saved.</param>
        /// <param name="context">The export context containing SolidWorks app and callbacks.</param>
        void ExportToPdf(DrawingDoc drawing, string saveDirectory, ExportContext context);
    }

    public class DrawingExporter : IDrawingExporter
    {
        public void ExportToPdf(DrawingDoc drawing, string saveDirectory, ExportContext context)
        {
            if (drawing == null)
                throw new ArgumentNullException(nameof(drawing));

            if (string.IsNullOrWhiteSpace(saveDirectory))
                throw new ArgumentException("Save directory cannot be null or empty.", nameof(saveDirectory));

            if (context == null)
                throw new ArgumentNullException(nameof(context));

            if (context.SolidWorksApp == null)
                throw new ArgumentException("SolidWorksApp cannot be null in context.", nameof(context));

            try
            {
                var pdfFileName = GetPdfFileName(drawing);
                var pdfPath = Path.Combine(saveDirectory, pdfFileName);

                var model = drawing as ModelDoc2;
                var sldWorks = context.SolidWorksApp;

                var exportData = sldWorks.GetExportFileData(
                    (int)swExportDataFileType_e.swExportPdfData) as ExportPdfData;

                if (exportData == null)
                {
                    throw new InvalidOperationException("Failed to get PDF export data from SolidWorks.");
                }

                exportData.ViewPdfAfterSaving = false;
                exportData.SetSheets(
                    (int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets,
                    drawing);

                context.ProgressCallback?.Invoke($"Exporting drawing to PDF: \"{pdfFileName}\"", null);

                int errors = 0;
                int warnings = 0;

                var modelExtension = model.Extension;
                var success = modelExtension.SaveAs(
                    pdfPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    exportData,
                    ref errors,
                    ref warnings);

                if (success && errors == 0)
                {
                    context.ProgressCallback?.Invoke($"Saved drawing to PDF: \"{pdfFileName}\"", Color.Green);
                }
                else if (success && warnings > 0)
                {
                    context.ProgressCallback?.Invoke(
                        $"PDF export completed with warnings: {warnings}",
                        Color.DarkBlue);
                }
                else
                {
                    context.ProgressCallback?.Invoke(
                        $"PDF export failed. Errors: {errors}, Warnings: {warnings}",
                        Color.Red);
                    throw new InvalidOperationException($"PDF export failed with {errors} errors and {warnings} warnings.");
                }
            }
            catch (Exception ex)
            {
                var errorMessage = $"Failed to export PDF: {ex.Message}";
                context.ProgressCallback?.Invoke(errorMessage, Color.Red);
                throw new InvalidOperationException(errorMessage, ex);
            }
        }

        private string GetPdfFileName(DrawingDoc drawing)
        {
            var model = drawing as ModelDoc2;
            var modelFilePath = model.GetPathName();

            if (string.IsNullOrEmpty(modelFilePath))
            {
                // Handle unsaved documents
                var title = model.GetTitle();
                return string.IsNullOrEmpty(title) ? "Untitled.pdf" : Path.GetFileNameWithoutExtension(title) + ".pdf";
            }

            return Path.GetFileNameWithoutExtension(modelFilePath) + ".pdf";
        }
    }
}