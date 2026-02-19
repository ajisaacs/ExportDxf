using FabWorks.Api.Services;
using FabWorks.Core.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.EntityFrameworkCore;

namespace FabWorks.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FileBrowserController : ControllerBase
    {
        private readonly IFileStorageService _fileStorage;
        private readonly FabWorksDbContext _db;
        private readonly FileExtensionContentTypeProvider _contentTypeProvider = new();

        public FileBrowserController(IFileStorageService fileStorage, FabWorksDbContext db)
        {
            _fileStorage = fileStorage;
            _db = db;
        }

        [HttpGet("files")]
        public async Task<ActionResult<FileListResult>> ListFiles(
            [FromQuery] string search = null,
            [FromQuery] string type = null)
        {
            var files = new List<StoredFileEntry>();

            // Query DXF files from CutTemplates
            if (type == null || type.Equals("dxf", StringComparison.OrdinalIgnoreCase))
            {
                var dxfQuery = _db.CutTemplates
                    .Where(c => c.ContentHash != null)
                    .Select(c => new
                    {
                        c.Id,
                        c.DxfFilePath,
                        c.ContentHash,
                        c.Thickness,
                        DrawingNumber = c.BomItem.ExportRecord.DrawingNumber,
                        CreatedAt = c.BomItem.ExportRecord.ExportedAt
                    });

                if (!string.IsNullOrWhiteSpace(search))
                {
                    var term = search.Trim().ToLower();
                    dxfQuery = dxfQuery.Where(c =>
                        c.DxfFilePath.ToLower().Contains(term) ||
                        c.DrawingNumber.ToLower().Contains(term));
                }

                var dxfResults = await dxfQuery
                    .OrderByDescending(c => c.CreatedAt)
                    .Take(500)
                    .ToListAsync();

                // Deduplicate by content hash (keep latest)
                var seenDxf = new HashSet<string>();
                foreach (var c in dxfResults)
                {
                    if (seenDxf.Contains(c.ContentHash)) continue;
                    seenDxf.Add(c.ContentHash);

                    var fileName = c.DxfFilePath?.Split(new[] { '/', '\\' }).LastOrDefault() ?? c.DxfFilePath;
                    files.Add(new StoredFileEntry
                    {
                        FileName = fileName,
                        ContentHash = c.ContentHash,
                        FileType = "dxf",
                        DrawingNumber = c.DrawingNumber,
                        Thickness = c.Thickness,
                        CreatedAt = c.CreatedAt
                    });
                }
            }

            // Query PDF files from ExportRecords
            if (type == null || type.Equals("pdf", StringComparison.OrdinalIgnoreCase))
            {
                var pdfQuery = _db.ExportRecords
                    .Where(r => r.PdfContentHash != null)
                    .Select(r => new
                    {
                        r.Id,
                        r.DrawingNumber,
                        r.PdfContentHash,
                        r.ExportedAt
                    });

                if (!string.IsNullOrWhiteSpace(search))
                {
                    var term = search.Trim().ToLower();
                    pdfQuery = pdfQuery.Where(r =>
                        r.DrawingNumber.ToLower().Contains(term));
                }

                var pdfResults = await pdfQuery
                    .OrderByDescending(r => r.ExportedAt)
                    .Take(500)
                    .ToListAsync();

                // Deduplicate by content hash
                var seenPdf = new HashSet<string>();
                foreach (var r in pdfResults)
                {
                    if (seenPdf.Contains(r.PdfContentHash)) continue;
                    seenPdf.Add(r.PdfContentHash);

                    files.Add(new StoredFileEntry
                    {
                        FileName = $"{r.DrawingNumber}.pdf",
                        ContentHash = r.PdfContentHash,
                        FileType = "pdf",
                        DrawingNumber = r.DrawingNumber,
                        CreatedAt = r.ExportedAt
                    });
                }
            }

            return new FileListResult
            {
                Total = files.Count,
                Files = files.OrderByDescending(f => f.CreatedAt).ToList()
            };
        }

        [HttpGet("preview")]
        public IActionResult PreviewFile([FromQuery] string hash, [FromQuery] string ext = "dxf")
        {
            if (string.IsNullOrEmpty(hash) || hash.Length < 4)
                return BadRequest("Invalid hash.");

            if (!_fileStorage.BlobExists(hash, ext))
                return NotFound("File not found.");

            var stream = _fileStorage.OpenBlob(hash, ext);
            if (stream == null)
                return NotFound("File not found.");

            var virtualName = $"file.{ext}";
            if (!_contentTypeProvider.TryGetContentType(virtualName, out var contentType))
                contentType = "application/octet-stream";

            return File(stream, contentType);
        }

        [HttpGet("download")]
        public IActionResult DownloadFile([FromQuery] string hash, [FromQuery] string ext = "dxf", [FromQuery] string name = null)
        {
            if (string.IsNullOrEmpty(hash) || hash.Length < 4)
                return BadRequest("Invalid hash.");

            if (!_fileStorage.BlobExists(hash, ext))
                return NotFound("File not found.");

            var stream = _fileStorage.OpenBlob(hash, ext);
            if (stream == null)
                return NotFound("File not found.");

            var fileName = name ?? $"{hash[..8]}.{ext}";
            if (!_contentTypeProvider.TryGetContentType(fileName, out var contentType))
                contentType = "application/octet-stream";

            return File(stream, contentType, fileName);
        }
    }

    public class FileListResult
    {
        public int Total { get; set; }
        public List<StoredFileEntry> Files { get; set; }
    }

    public class StoredFileEntry
    {
        public string FileName { get; set; }
        public string ContentHash { get; set; }
        public string FileType { get; set; }
        public string DrawingNumber { get; set; }
        public double? Thickness { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}
