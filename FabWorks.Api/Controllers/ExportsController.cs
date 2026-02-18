using FabWorks.Api.DTOs;
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

        public ExportsController(FabWorksDbContext db) => _db = db;

        [HttpPost]
        public async Task<ActionResult<ExportDetailDto>> Create(CreateExportRequest request)
        {
            var record = new ExportRecord
            {
                DrawingNumber = request.DrawingNumber,
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

        private static ExportDetailDto MapToDto(ExportRecord r) => new()
        {
            Id = r.Id,
            DrawingNumber = r.DrawingNumber,
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
