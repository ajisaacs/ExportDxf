using System;
using System.IO;

namespace ExportDXF.Services
{
    public interface IFileExportService
    {
        string OutputFolder { get; }
        string GetDrawingOutputFolder(string equipment, string drawingNo);
        string SaveDxfFile(string sourcePath, string drawingNumber, string itemNo);
        string SavePdfFile(string sourcePath, string drawingNumber, string outputFolder = null);
        void EnsureOutputFolderExists();
        string StashFile(string filePath);
        void ArchiveFile(string stashPath, string originalPath);
        void DiscardStash(string stashPath);
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

        public string GetDrawingOutputFolder(string equipment, string drawingNo)
        {
            if (string.IsNullOrEmpty(equipment) || string.IsNullOrEmpty(drawingNo))
                return OutputFolder;

            var folder = Path.Combine(OutputFolder, equipment, drawingNo);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            return folder;
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

        public string SavePdfFile(string sourcePath, string drawingNumber, string outputFolder = null)
        {
            if (string.IsNullOrEmpty(sourcePath))
                throw new ArgumentNullException(nameof(sourcePath));

            var folder = outputFolder ?? OutputFolder;
            var fileName = !string.IsNullOrEmpty(drawingNumber)
                ? $"{drawingNumber}.pdf"
                : Path.GetFileName(sourcePath);

            var destPath = Path.Combine(folder, fileName);

            // If source and dest are the same, skip copy
            if (!string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourcePath, destPath, overwrite: true);
            }

            return destPath;
        }

        public string StashFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return null;

            var stashPath = filePath + ".bak";
            File.Move(filePath, stashPath, overwrite: true);
            return stashPath;
        }

        public void ArchiveFile(string stashPath, string originalPath)
        {
            if (string.IsNullOrEmpty(stashPath) || !File.Exists(stashPath))
                return;

            var fileDir = Path.GetDirectoryName(originalPath) ?? OutputFolder;
            var archiveDir = Path.Combine(fileDir, "_archive");
            if (!Directory.Exists(archiveDir))
                Directory.CreateDirectory(archiveDir);

            var originalName = Path.GetFileNameWithoutExtension(originalPath);
            var ext = Path.GetExtension(originalPath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var archiveName = $"{originalName} [{timestamp}]{ext}";
            var archivePath = Path.Combine(archiveDir, archiveName);

            File.Move(stashPath, archivePath, overwrite: true);
        }

        public void DiscardStash(string stashPath)
        {
            if (!string.IsNullOrEmpty(stashPath) && File.Exists(stashPath))
                File.Delete(stashPath);
        }
    }
}
