using FabWorks.Api.Configuration;
using FabWorks.Core.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;

namespace FabWorks.Api.Services
{
    public class FileUploadResult
    {
        public string ContentHash { get; set; }
        public string FileName { get; set; }
        public bool WasUnchanged { get; set; }
        public bool IsNewFile { get; set; }
    }

    public interface IFileStorageService
    {
        string OutputFolder { get; }
        Task<FileUploadResult> StoreDxfAsync(Stream stream, string equipment, string drawingNo, string itemNo, string contentHash);
        Task<FileUploadResult> StorePdfAsync(Stream stream, string equipment, string drawingNo, string contentHash, int? exportRecordId = null);
        Stream OpenBlob(string contentHash, string extension);
        bool BlobExists(string contentHash, string extension);
    }

    public class FileStorageService : IFileStorageService
    {
        private readonly FileStorageOptions _options;
        private readonly FabWorksDbContext _db;

        public string OutputFolder => _options.OutputFolder;

        public FileStorageService(IOptions<FileStorageOptions> options, FabWorksDbContext db)
        {
            _options = options.Value;
            _db = db;

            var blobRoot = Path.Combine(_options.OutputFolder, "blobs");
            if (!Directory.Exists(blobRoot))
                Directory.CreateDirectory(blobRoot);
        }

        public async Task<FileUploadResult> StoreDxfAsync(Stream stream, string equipment, string drawingNo, string itemNo, string contentHash)
        {
            var fileName = BuildDxfFileName(drawingNo, equipment, itemNo);

            // Look up previous hash by drawing number + item number
            var drawingNumber = BuildDrawingNumber(equipment, drawingNo);
            var previousHash = await _db.CutTemplates
                .Where(c => c.BomItem.ExportRecord.DrawingNumber == drawingNumber
                         && c.BomItem.ItemNo == itemNo
                         && c.ContentHash != null)
                .OrderByDescending(c => c.Id)
                .Select(c => c.ContentHash)
                .FirstOrDefaultAsync();

            var wasUnchanged = previousHash != null && previousHash == contentHash;
            var isNewFile = await StoreBlobAsync(stream, contentHash, "dxf");

            return new FileUploadResult
            {
                ContentHash = contentHash,
                FileName = fileName,
                WasUnchanged = wasUnchanged,
                IsNewFile = isNewFile
            };
        }

        public async Task<FileUploadResult> StorePdfAsync(Stream stream, string equipment, string drawingNo, string contentHash, int? exportRecordId = null)
        {
            var drawingNumber = BuildDrawingNumber(equipment, drawingNo);
            var fileName = $"{drawingNumber}.pdf";

            // Look up previous PDF hash
            var previousHash = await _db.ExportRecords
                .Where(r => r.DrawingNumber == drawingNumber
                         && r.PdfContentHash != null
                         && (exportRecordId == null || r.Id != exportRecordId))
                .OrderByDescending(r => r.Id)
                .Select(r => r.PdfContentHash)
                .FirstOrDefaultAsync();

            var wasUnchanged = previousHash != null && previousHash == contentHash;
            var isNewFile = await StoreBlobAsync(stream, contentHash, "pdf");

            return new FileUploadResult
            {
                ContentHash = contentHash,
                FileName = fileName,
                WasUnchanged = wasUnchanged,
                IsNewFile = isNewFile
            };
        }

        public Stream OpenBlob(string contentHash, string extension)
        {
            var path = GetBlobPath(contentHash, extension);
            if (!File.Exists(path))
                return null;
            return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        }

        public bool BlobExists(string contentHash, string extension)
        {
            return File.Exists(GetBlobPath(contentHash, extension));
        }

        private async Task<bool> StoreBlobAsync(Stream stream, string contentHash, string extension)
        {
            var blobPath = GetBlobPath(contentHash, extension);

            if (File.Exists(blobPath))
                return false; // blob already exists (dedup)

            var dir = Path.GetDirectoryName(blobPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using var fileStream = new FileStream(blobPath, FileMode.Create, FileAccess.Write);
            await stream.CopyToAsync(fileStream);
            return true; // new blob written
        }

        private string GetBlobPath(string contentHash, string extension)
        {
            var prefix1 = contentHash[..2];
            var prefix2 = contentHash[2..4];
            return Path.Combine(_options.OutputFolder, "blobs", prefix1, prefix2, $"{contentHash}.{extension}");
        }

        private static string BuildDrawingNumber(string equipment, string drawingNo)
        {
            if (!string.IsNullOrEmpty(equipment) && !string.IsNullOrEmpty(drawingNo))
                return $"{equipment} {drawingNo}";
            if (!string.IsNullOrEmpty(equipment))
                return equipment;
            return drawingNo ?? "";
        }

        private static string BuildDxfFileName(string drawingNo, string equipment, string itemNo)
        {
            var drawingNumber = BuildDrawingNumber(equipment, drawingNo);
            var paddedItem = (itemNo ?? "").PadLeft(2, '0');
            if (!string.IsNullOrEmpty(drawingNumber) && !string.IsNullOrEmpty(itemNo))
                return $"{drawingNumber} PT{paddedItem}.dxf";
            return $"PT{paddedItem}.dxf";
        }
    }
}
