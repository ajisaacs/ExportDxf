using FabWorks.Api.DTOs;
using FabWorks.Api.Services;
using FabWorks.Core.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace FabWorks.Api.Controllers
{
    [ApiController]
    [Route("api/form-programs")]
    public class FormProgramsController : ControllerBase
    {
        private readonly FabWorksDbContext _db;
        private readonly FormProgramService _formService;

        public FormProgramsController(FabWorksDbContext db, FormProgramService formService)
        {
            _db = db;
            _formService = formService;
        }

        [HttpGet("by-drawing")]
        public async Task<ActionResult<List<FormProgramDto>>> GetByDrawing([FromQuery] string drawingNumber)
        {
            var programs = await _db.FormPrograms
                .Include(fp => fp.BomItem)
                .ThenInclude(b => b.ExportRecord)
                .Where(fp => fp.BomItem.ExportRecord.DrawingNumber == drawingNumber)
                .ToListAsync();

            return programs.Select(fp => new FormProgramDto
            {
                Id = fp.Id,
                ProgramFilePath = fp.ProgramFilePath,
                ContentHash = fp.ContentHash,
                ProgramName = fp.ProgramName,
                Thickness = fp.Thickness,
                MaterialType = fp.MaterialType,
                KFactor = fp.KFactor,
                BendCount = fp.BendCount,
                UpperToolNames = fp.UpperToolNames,
                LowerToolNames = fp.LowerToolNames,
                SetupNotes = fp.SetupNotes
            }).ToList();
        }

        [HttpPost("parse")]
        public ActionResult<FormProgramDto> Parse([FromQuery] string filePath)
        {
            if (!System.IO.File.Exists(filePath))
                return NotFound($"File not found: {filePath}");

            var fp = _formService.ParseFromFile(filePath);
            return new FormProgramDto
            {
                ProgramFilePath = fp.ProgramFilePath,
                ContentHash = fp.ContentHash,
                ProgramName = fp.ProgramName,
                Thickness = fp.Thickness,
                MaterialType = fp.MaterialType,
                KFactor = fp.KFactor,
                BendCount = fp.BendCount,
                UpperToolNames = fp.UpperToolNames,
                LowerToolNames = fp.LowerToolNames,
                SetupNotes = fp.SetupNotes
            };
        }

        [HttpPost("{bomItemId}")]
        public async Task<ActionResult<FormProgramDto>> AttachToItem(int bomItemId, [FromQuery] string filePath)
        {
            var bomItem = await _db.BomItems
                .Include(b => b.FormProgram)
                .FirstOrDefaultAsync(b => b.ID == bomItemId);

            if (bomItem == null) return NotFound("BOM item not found");

            if (!System.IO.File.Exists(filePath))
                return NotFound($"File not found: {filePath}");

            var fp = _formService.ParseFromFile(filePath);
            fp.BomItemId = bomItemId;

            if (bomItem.FormProgram != null)
                _db.FormPrograms.Remove(bomItem.FormProgram);

            bomItem.FormProgram = fp;
            await _db.SaveChangesAsync();

            return new FormProgramDto
            {
                Id = fp.Id,
                ProgramFilePath = fp.ProgramFilePath,
                ContentHash = fp.ContentHash,
                ProgramName = fp.ProgramName,
                Thickness = fp.Thickness,
                MaterialType = fp.MaterialType,
                KFactor = fp.KFactor,
                BendCount = fp.BendCount,
                UpperToolNames = fp.UpperToolNames,
                LowerToolNames = fp.LowerToolNames,
                SetupNotes = fp.SetupNotes
            };
        }
    }
}
