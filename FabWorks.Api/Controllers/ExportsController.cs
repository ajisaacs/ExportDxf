using System.IO.Compression;
using FabWorks.Api.DTOs;
using FabWorks.Api.Services;
using FabWorks.Core.Data;
using FabWorks.Core.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace FabWorks.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExportsController : ControllerBase
    {
        private readonly FabWorksDbContext _db;
        private readonly IFileStorageService _fileStorage;

        public ExportsController(FabWorksDbContext db, IFileStorageService fileStorage)
        {
            _db = db;
            _fileStorage = fileStorage;
        }

        [HttpGet]
        public async Task<ActionResult<object>> List(
            [FromQuery] string search = null,
            [FromQuery] int skip = 0,
            [FromQuery] int take = 50)
        {
            var query = _db.ExportRecords
                .Include(r => r.BomItems)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(search))
            {
                var term = search.Trim().ToLower();
                query = query.Where(r =>
                    r.DrawingNumber.ToLower().Contains(term) ||
                    (r.Title != null && r.Title.ToLower().Contains(term)) ||
                    r.ExportedBy.ToLower().Contains(term) ||
                    r.BomItems.Any(b => b.PartName.ToLower().Contains(term) ||
                                        b.Description.ToLower().Contains(term)));
            }

            var total = await query.CountAsync();

            var records = await query
                .OrderByDescending(r => r.ExportedAt)
                .Skip(skip)
                .Take(take)
                .ToListAsync();

            return new
            {
                total,
                items = records.Select(r => new
                {
                    r.Id,
                    r.DrawingNumber,
                    r.Title,
                    r.SourceFilePath,
                    r.ExportedAt,
                    r.ExportedBy,
                    BomItemCount = r.BomItems?.Count ?? 0
                })
            };
        }

        [HttpPost]
        public async Task<ActionResult<ExportDetailDto>> Create(CreateExportRequest request)
        {
            var record = new ExportRecord
            {
                DrawingNumber = request.DrawingNumber,
                Title = request.Title,
                EquipmentNo = request.EquipmentNo,
                DrawingNo = request.DrawingNo,
                SourceFilePath = request.SourceFilePath,
                OutputFolder = request.OutputFolder,
                ExportedAt = DateTime.Now,
                ExportedBy = Environment.UserName
            };

            _db.ExportRecords.Add(record);
            await _db.SaveChangesAsync();

            return CreatedAtAction(nameof(GetById), new { id = record.Id }, MapToDto(record));
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<ExportDetailDto>> GetById(int id)
        {
            var record = await _db.ExportRecords
                .Include(r => r.BomItems).ThenInclude(b => b.CutTemplate)
                .Include(r => r.BomItems).ThenInclude(b => b.FormProgram)
                .FirstOrDefaultAsync(r => r.Id == id);

            if (record == null) return NotFound();
            return MapToDto(record);
        }

        [HttpGet("by-source")]
        public async Task<ActionResult<ExportDetailDto>> GetBySourceFile([FromQuery] string path)
        {
            var record = await _db.ExportRecords
                .Where(r => r.SourceFilePath.ToLower() == path.ToLower()
                         && !string.IsNullOrEmpty(r.DrawingNumber))
                .OrderByDescending(r => r.Id)
                .FirstOrDefaultAsync();

            if (record == null) return NotFound();
            return MapToDto(record);
        }

        [HttpGet("by-drawing")]
        public async Task<ActionResult<List<ExportDetailDto>>> GetByDrawing([FromQuery] string drawingNumber)
        {
            var records = await _db.ExportRecords
                .Include(r => r.BomItems).ThenInclude(b => b.CutTemplate)
                .Include(r => r.BomItems).ThenInclude(b => b.FormProgram)
                .Where(r => r.DrawingNumber == drawingNumber)
                .OrderByDescending(r => r.ExportedAt)
                .ToListAsync();

            return records.Select(MapToDto).ToList();
        }

        [HttpGet("next-item-number")]
        public async Task<ActionResult<string>> GetNextItemNumber([FromQuery] string drawingNumber)
        {
            if (string.IsNullOrEmpty(drawingNumber)) return "1";

            var existingItems = await _db.ExportRecords
                .Where(r => r.DrawingNumber == drawingNumber)
                .SelectMany(r => r.BomItems)
                .Select(b => b.ItemNo)
                .ToListAsync();

            int maxNum = 0;
            foreach (var itemNo in existingItems)
            {
                if (int.TryParse(itemNo, out var num) && num > maxNum)
                    maxNum = num;
            }
            return (maxNum + 1).ToString();
        }

        [HttpGet("drawing-numbers")]
        public async Task<ActionResult<List<string>>> GetDrawingNumbers()
        {
            var numbers = await _db.ExportRecords
                .Select(r => r.DrawingNumber)
                .Where(d => !string.IsNullOrEmpty(d))
                .Distinct()
                .ToListAsync();

            return numbers;
        }

        [HttpGet("equipment-numbers")]
        public async Task<ActionResult<List<string>>> GetEquipmentNumbers()
        {
            var numbers = await _db.ExportRecords
                .Select(r => r.EquipmentNo)
                .Where(e => !string.IsNullOrEmpty(e))
                .Distinct()
                .OrderBy(e => e)
                .ToListAsync();

            return numbers;
        }

        [HttpGet("drawing-numbers-by-equipment")]
        public async Task<ActionResult<List<string>>> GetDrawingNumbersByEquipment([FromQuery] string equipmentNo)
        {
            var query = _db.ExportRecords
                .Where(r => !string.IsNullOrEmpty(r.DrawingNo));

            if (!string.IsNullOrEmpty(equipmentNo))
                query = query.Where(r => r.EquipmentNo == equipmentNo);

            var numbers = await query
                .Select(r => r.DrawingNo)
                .Distinct()
                .OrderBy(d => d)
                .ToListAsync();

            return numbers;
        }

        [HttpGet("previous-pdf-hash")]
        public async Task<ActionResult<string>> GetPreviousPdfHash(
            [FromQuery] string drawingNumber,
            [FromQuery] int? excludeId = null)
        {
            var hash = await _db.ExportRecords
                .Where(r => r.DrawingNumber == drawingNumber
                         && r.PdfContentHash != null
                         && (excludeId == null || r.Id != excludeId))
                .OrderByDescending(r => r.Id)
                .Select(r => r.PdfContentHash)
                .FirstOrDefaultAsync();

            if (hash == null) return NotFound();
            return hash;
        }

        [HttpPatch("{id}/pdf-hash")]
        public async Task<IActionResult> UpdatePdfHash(int id, [FromBody] UpdatePdfHashRequest request)
        {
            var record = await _db.ExportRecords.FindAsync(id);
            if (record == null) return NotFound();

            record.PdfContentHash = request.PdfContentHash;
            await _db.SaveChangesAsync();

            return NoContent();
        }

        [HttpGet("previous-cut-template")]
        public async Task<ActionResult<CutTemplateDto>> GetPreviousCutTemplate(
            [FromQuery] string drawingNumber,
            [FromQuery] string itemNo)
        {
            if (string.IsNullOrEmpty(drawingNumber) || string.IsNullOrEmpty(itemNo))
                return BadRequest("drawingNumber and itemNo are required.");

            var ct = await _db.CutTemplates
                .Where(c => c.BomItem.ExportRecord.DrawingNumber == drawingNumber
                         && c.BomItem.ItemNo == itemNo
                         && c.ContentHash != null)
                .OrderByDescending(c => c.Id)
                .FirstOrDefaultAsync();

            if (ct == null) return NotFound();

            return new CutTemplateDto
            {
                Id = ct.Id,
                DxfFilePath = ct.DxfFilePath,
                ContentHash = ct.ContentHash,
                Thickness = ct.Thickness,
                KFactor = ct.KFactor,
                DefaultBendRadius = ct.DefaultBendRadius
            };
        }

        [HttpGet("{id}/download-dxfs")]
        public async Task<IActionResult> DownloadAllDxfs(int id)
        {
            var record = await _db.ExportRecords
                .Include(r => r.BomItems).ThenInclude(b => b.CutTemplate)
                .FirstOrDefaultAsync(r => r.Id == id);

            if (record == null) return NotFound();

            var dxfItems = record.BomItems
                .Where(b => b.CutTemplate?.ContentHash != null)
                .ToList();

            if (dxfItems.Count == 0) return NotFound("No DXF files for this export.");

            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var b in dxfItems)
                {
                    var ct = b.CutTemplate;
                    var fileName = ct.DxfFilePath?.Split(new[] { '/', '\\' }).LastOrDefault()
                                   ?? $"PT{(b.ItemNo ?? "").PadLeft(2, '0')}.dxf";

                    // Ensure unique names in zip
                    if (!usedNames.Add(fileName))
                    {
                        var baseName = Path.GetFileNameWithoutExtension(fileName);
                        var ext = Path.GetExtension(fileName);
                        var counter = 2;
                        do { fileName = $"{baseName}_{counter++}{ext}"; }
                        while (!usedNames.Add(fileName));
                    }

                    var blobStream = _fileStorage.OpenBlob(ct.ContentHash, "dxf");
                    if (blobStream == null) continue;

                    var entry = zip.CreateEntry(fileName, CompressionLevel.Fastest);
                    using var entryStream = entry.Open();
                    await blobStream.CopyToAsync(entryStream);
                    blobStream.Dispose();
                }
            }

            ms.Position = 0;
            var zipName = $"{record.DrawingNumber ?? $"Export-{id}"} DXFs.zip";
            return File(ms, "application/zip", zipName);
        }

        private static ExportDetailDto MapToDto(ExportRecord r) => new()
        {
            Id = r.Id,
            DrawingNumber = r.DrawingNumber,
            Title = r.Title,
            EquipmentNo = r.EquipmentNo,
            DrawingNo = r.DrawingNo,
            SourceFilePath = r.SourceFilePath,
            OutputFolder = r.OutputFolder,
            ExportedAt = r.ExportedAt,
            ExportedBy = r.ExportedBy,
            PdfContentHash = r.PdfContentHash,
            BomItems = r.BomItems?.Select(b => new BomItemDto
            {
                ID = b.ID,
                ItemNo = b.ItemNo,
                PartNo = b.PartNo,
                SortOrder = b.SortOrder,
                Qty = b.Qty,
                TotalQty = b.TotalQty,
                Description = b.Description,
                PartName = b.PartName,
                ConfigurationName = b.ConfigurationName,
                Material = b.Material,
                CutTemplate = b.CutTemplate == null ? null : new CutTemplateDto
                {
                    Id = b.CutTemplate.Id,
                    DxfFilePath = b.CutTemplate.DxfFilePath,
                    ContentHash = b.CutTemplate.ContentHash,
                    Thickness = b.CutTemplate.Thickness,
                    KFactor = b.CutTemplate.KFactor,
                    DefaultBendRadius = b.CutTemplate.DefaultBendRadius
                },
                FormProgram = b.FormProgram == null ? null : new FormProgramDto
                {
                    Id = b.FormProgram.Id,
                    ProgramFilePath = b.FormProgram.ProgramFilePath,
                    ContentHash = b.FormProgram.ContentHash,
                    ProgramName = b.FormProgram.ProgramName,
                    Thickness = b.FormProgram.Thickness,
                    MaterialType = b.FormProgram.MaterialType,
                    KFactor = b.FormProgram.KFactor,
                    BendCount = b.FormProgram.BendCount,
                    UpperToolNames = b.FormProgram.UpperToolNames,
                    LowerToolNames = b.FormProgram.LowerToolNames,
                    SetupNotes = b.FormProgram.SetupNotes
                }
            }).ToList() ?? new()
        };
    }
}
