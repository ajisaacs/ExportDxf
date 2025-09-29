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
        /// <param name="progressCallback">Optional callback for progress updates.</param>
        void ExportToPdf(DrawingDoc drawing, string saveDirectory, Action<string, Color?> progressCallback);
    }


    public class DrawingExporter : IDrawingExporter
    {
        public void ExportToPdf(DrawingDoc drawing, string saveDirectory, Action<string, Color?> progressCallback)
        {
            if (drawing == null)
                throw new ArgumentNullException(nameof(drawing));

            if (string.IsNullOrWhiteSpace(saveDirectory))
                throw new ArgumentException("Save directory cannot be null or empty.", nameof(saveDirectory));

            try
            {
                var pdfFileName = GetPdfFileName(drawing);
                var pdfPath = Path.Combine(saveDirectory, pdfFileName);

                var model = drawing as ModelDoc2;
                var sldWorks = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));

                var exportData = sldWorks.GetExportFileData(
                    (int)swExportDataFileType_e.swExportPdfData) as ExportPdfData;

                exportData.ViewPdfAfterSaving = false;
                exportData.SetSheets(
                    (int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets,
                    drawing);

                int errors = 0;
                int warnings = 0;

                var modelExtension = model.Extension;
                modelExtension.SaveAs(
                    pdfPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    exportData,
                    ref errors,
                    ref warnings);

                if (errors == 0)
                {
                    progressCallback?.Invoke($"Saved drawing to PDF: \"{pdfFileName}\"", Color.Green);
                }
                else
                {
                    progressCallback?.Invoke($"PDF export completed with errors: {errors}", Color.Yellow);
                }
            }
            catch (Exception ex)
            {
                progressCallback?.Invoke($"Failed to export PDF: {ex.Message}", Color.Red);
                throw;
            }
        }

        private string GetPdfFileName(DrawingDoc drawing)
        {
            var model = drawing as ModelDoc2;
            var modelFilePath = model.GetPathName();
            return Path.GetFileNameWithoutExtension(modelFilePath) + ".pdf";
        }
    }
}