using FabWorks.Api.DTOs;
using FabWorks.Api.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;

namespace FabWorks.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FilesController : ControllerBase
    {
        private readonly IFileStorageService _fileStorage;
        private readonly FileExtensionContentTypeProvider _contentTypeProvider = new();

        public FilesController(IFileStorageService fileStorage)
        {
            _fileStorage = fileStorage;
        }

        [HttpPost("dxf")]
        [RequestSizeLimit(50_000_000)] // 50 MB
        public async Task<ActionResult<FileUploadResponse>> UploadDxf(
            IFormFile file,
            [FromForm] string equipment,
            [FromForm] string drawingNo,
            [FromForm] string itemNo,
            [FromForm] string contentHash)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            using var stream = file.OpenReadStream();
            var result = await _fileStorage.StoreDxfAsync(stream, equipment, drawingNo, itemNo, contentHash);

            return Ok(new FileUploadResponse
            {
                StoredFilePath = result.FileName,
                ContentHash = result.ContentHash,
                FileName = result.FileName,
                WasUnchanged = result.WasUnchanged,
                IsNewFile = result.IsNewFile
            });
        }

        [HttpPost("pdf")]
        [RequestSizeLimit(100_000_000)] // 100 MB
        public async Task<ActionResult<FileUploadResponse>> UploadPdf(
            IFormFile file,
            [FromForm] string equipment,
            [FromForm] string drawingNo,
            [FromForm] string contentHash,
            [FromForm] int? exportRecordId = null)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            using var stream = file.OpenReadStream();
            var result = await _fileStorage.StorePdfAsync(stream, equipment, drawingNo, contentHash, exportRecordId);

            return Ok(new FileUploadResponse
            {
                StoredFilePath = result.FileName,
                ContentHash = result.ContentHash,
                FileName = result.FileName,
                WasUnchanged = result.WasUnchanged,
                IsNewFile = result.IsNewFile
            });
        }

        [HttpGet("blob/{hash}")]
        public IActionResult GetBlob(string hash, [FromQuery] string ext = "dxf", [FromQuery] bool download = false, [FromQuery] string name = null)
        {
            if (string.IsNullOrEmpty(hash) || hash.Length < 4)
                return BadRequest("Invalid hash.");

            if (!_fileStorage.BlobExists(hash, ext))
                return NotFound("Blob not found.");

            var stream = _fileStorage.OpenBlob(hash, ext);
            if (stream == null)
                return NotFound("Blob not found.");

            var fileName = !string.IsNullOrEmpty(name) ? name : $"{hash[..8]}.{ext}";
            if (!_contentTypeProvider.TryGetContentType(fileName, out var contentType))
                contentType = "application/octet-stream";

            if (download)
                return File(stream, contentType, fileName);

            return File(stream, contentType);
        }
    }
}
