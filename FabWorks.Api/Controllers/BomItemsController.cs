using FabWorks.Api.DTOs;
using FabWorks.Core.Data;
using FabWorks.Core.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace FabWorks.Api.Controllers
{
    [ApiController]
    [Route("api/exports/{exportId}/bom-items")]
    public class BomItemsController : ControllerBase
    {
        private readonly FabWorksDbContext _db;

        public BomItemsController(FabWorksDbContext db) => _db = db;

        [HttpGet("find")]
        public async Task<ActionResult<BomItemDto>> FindExisting(int exportId, [FromQuery] string partName, [FromQuery] string configurationName)
        {
            var export = await _db.ExportRecords.FindAsync(exportId);
            if (export == null) return NotFound();

            var existing = await _db.BomItems
                .Include(b => b.CutTemplate)
                .Include(b => b.FormProgram)
                .Include(b => b.ExportRecord)
                .Where(b => b.ExportRecord.DrawingNumber == export.DrawingNumber
                         && b.PartName == (partName ?? "")
                         && b.ConfigurationName == (configurationName ?? ""))
                .OrderByDescending(b => b.ID)
                .FirstOrDefaultAsync();

            if (existing == null) return NotFound();
            return MapToDto(existing);
        }

        [HttpGet]
        public async Task<ActionResult<List<BomItemDto>>> GetByExport(int exportId)
        {
            var items = await _db.BomItems
                .Include(b => b.CutTemplate)
                .Include(b => b.FormProgram)
                .Where(b => b.ExportRecordId == exportId)
                .OrderBy(b => b.SortOrder)
                .ToListAsync();

            return items.Select(MapToDto).ToList();
        }

        [HttpPost]
        public async Task<ActionResult<BomItemDto>> Create(int exportId, BomItemDto dto)
        {
            var export = await _db.ExportRecords.FindAsync(exportId);
            if (export == null) return NotFound("Export record not found");

            // Look for existing BomItem with same PartName + ConfigurationName under the same drawing
            var existing = await _db.BomItems
                .Include(b => b.CutTemplate)
                .Include(b => b.FormProgram)
                .Include(b => b.ExportRecord)
                .Where(b => b.ExportRecord.DrawingNumber == export.DrawingNumber
                         && b.PartName == (dto.PartName ?? "")
                         && b.ConfigurationName == (dto.ConfigurationName ?? ""))
                .OrderByDescending(b => b.ID)
                .FirstOrDefaultAsync();

            if (existing != null)
            {
                // Update existing: move to new export record and refresh fields
                existing.ExportRecordId = exportId;
                existing.PartNo = dto.PartNo ?? "";
                existing.SortOrder = dto.SortOrder;
                existing.Qty = dto.Qty;
                existing.TotalQty = dto.TotalQty;
                existing.Description = dto.Description ?? "";
                existing.Material = dto.Material ?? "";

                if (dto.CutTemplate != null)
                {
                    if (existing.CutTemplate != null)
                    {
                        existing.CutTemplate.DxfFilePath = dto.CutTemplate.DxfFilePath ?? "";
                        existing.CutTemplate.ContentHash = dto.CutTemplate.ContentHash;
                        existing.CutTemplate.Thickness = dto.CutTemplate.Thickness;
                        existing.CutTemplate.KFactor = dto.CutTemplate.KFactor;
                        existing.CutTemplate.DefaultBendRadius = dto.CutTemplate.DefaultBendRadius;
                    }
                    else
                    {
                        existing.CutTemplate = new CutTemplate
                        {
                            DxfFilePath = dto.CutTemplate.DxfFilePath ?? "",
                            ContentHash = dto.CutTemplate.ContentHash,
                            Thickness = dto.CutTemplate.Thickness,
                            KFactor = dto.CutTemplate.KFactor,
                            DefaultBendRadius = dto.CutTemplate.DefaultBendRadius
                        };
                    }
                }

                if (dto.FormProgram != null)
                {
                    if (existing.FormProgram != null)
                    {
                        existing.FormProgram.ProgramFilePath = dto.FormProgram.ProgramFilePath ?? "";
                        existing.FormProgram.ContentHash = dto.FormProgram.ContentHash;
                        existing.FormProgram.ProgramName = dto.FormProgram.ProgramName ?? "";
                        existing.FormProgram.Thickness = dto.FormProgram.Thickness;
                        existing.FormProgram.MaterialType = dto.FormProgram.MaterialType ?? "";
                        existing.FormProgram.KFactor = dto.FormProgram.KFactor;
                        existing.FormProgram.BendCount = dto.FormProgram.BendCount;
                        existing.FormProgram.UpperToolNames = dto.FormProgram.UpperToolNames ?? "";
                        existing.FormProgram.LowerToolNames = dto.FormProgram.LowerToolNames ?? "";
                        existing.FormProgram.SetupNotes = dto.FormProgram.SetupNotes ?? "";
                    }
                    else
                    {
                        existing.FormProgram = new FormProgram
                        {
                            ProgramFilePath = dto.FormProgram.ProgramFilePath ?? "",
                            ContentHash = dto.FormProgram.ContentHash,
                            ProgramName = dto.FormProgram.ProgramName ?? "",
                            Thickness = dto.FormProgram.Thickness,
                            MaterialType = dto.FormProgram.MaterialType ?? "",
                            KFactor = dto.FormProgram.KFactor,
                            BendCount = dto.FormProgram.BendCount,
                            UpperToolNames = dto.FormProgram.UpperToolNames ?? "",
                            LowerToolNames = dto.FormProgram.LowerToolNames ?? "",
                            SetupNotes = dto.FormProgram.SetupNotes ?? ""
                        };
                    }
                }

                await _db.SaveChangesAsync();
                return Ok(MapToDto(existing));
            }

            // No existing match — create new
            var item = new BomItem
            {
                ExportRecordId = exportId,
                ItemNo = dto.ItemNo ?? "",
                PartNo = dto.PartNo ?? "",
                SortOrder = dto.SortOrder,
                Qty = dto.Qty,
                TotalQty = dto.TotalQty,
                Description = dto.Description ?? "",
                PartName = dto.PartName ?? "",
                ConfigurationName = dto.ConfigurationName ?? "",
                Material = dto.Material ?? ""
            };

            if (dto.CutTemplate != null)
            {
                item.CutTemplate = new CutTemplate
                {
                    DxfFilePath = dto.CutTemplate.DxfFilePath ?? "",
                    ContentHash = dto.CutTemplate.ContentHash,
                    Thickness = dto.CutTemplate.Thickness,
                    KFactor = dto.CutTemplate.KFactor,
                    DefaultBendRadius = dto.CutTemplate.DefaultBendRadius
                };
            }

            if (dto.FormProgram != null)
            {
                item.FormProgram = new FormProgram
                {
                    ProgramFilePath = dto.FormProgram.ProgramFilePath ?? "",
                    ContentHash = dto.FormProgram.ContentHash,
                    ProgramName = dto.FormProgram.ProgramName ?? "",
                    Thickness = dto.FormProgram.Thickness,
                    MaterialType = dto.FormProgram.MaterialType ?? "",
                    KFactor = dto.FormProgram.KFactor,
                    BendCount = dto.FormProgram.BendCount,
                    UpperToolNames = dto.FormProgram.UpperToolNames ?? "",
                    LowerToolNames = dto.FormProgram.LowerToolNames ?? "",
                    SetupNotes = dto.FormProgram.SetupNotes ?? ""
                };
            }

            _db.BomItems.Add(item);
            await _db.SaveChangesAsync();

            return CreatedAtAction(nameof(GetByExport), new { exportId }, MapToDto(item));
        }

        private static BomItemDto MapToDto(BomItem b) => new()
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
        };
    }
}
