using System;
using System.IO;

namespace ExportDXF.Services
{
    public interface IFileExportService
    {
        string OutputFolder { get; }
        string SaveDxfFile(string sourcePath, string drawingNumber, string itemNo);
        string SavePdfFile(string sourcePath, string drawingNumber);
        void EnsureOutputFolderExists();
    }

    public class FileExportService : IFileExportService
    {
        public string OutputFolder { get; }

        public FileExportService(string outputFolder)
        {
            OutputFolder = outputFolder ?? throw new ArgumentNullException(nameof(outputFolder));
            EnsureOutputFolderExists();
        }

        public void EnsureOutputFolderExists()
        {
            if (!Directory.Exists(OutputFolder))
            {
                Directory.CreateDirectory(OutputFolder);
            }
        }

        public string SaveDxfFile(string sourcePath, string drawingNumber, string itemNo)
        {
            if (string.IsNullOrEmpty(sourcePath))
                throw new ArgumentNullException(nameof(sourcePath));

            var fileName = !string.IsNullOrEmpty(drawingNumber) && !string.IsNullOrEmpty(itemNo)
                ? $"{drawingNumber} PT{itemNo}.dxf"
                : Path.GetFileName(sourcePath);

            var destPath = Path.Combine(OutputFolder, fileName);

            // If source and dest are the same, skip copy
            if (!string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourcePath, destPath, overwrite: true);
            }

            return destPath;
        }

        public string SavePdfFile(string sourcePath, string drawingNumber)
        {
            if (string.IsNullOrEmpty(sourcePath))
                throw new ArgumentNullException(nameof(sourcePath));

            var fileName = !string.IsNullOrEmpty(drawingNumber)
                ? $"{drawingNumber}.pdf"
                : Path.GetFileName(sourcePath);

            var destPath = Path.Combine(OutputFolder, fileName);

            // If source and dest are the same, skip copy
            if (!string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourcePath, destPath, overwrite: true);
            }

            return destPath;
        }
    }
}
