using FabWorks.Core.Models;
using Microsoft.EntityFrameworkCore;

namespace FabWorks.Core.Data
{
    public class FabWorksDbContext : DbContext
    {
        public DbSet<ExportRecord> ExportRecords { get; set; }
        public DbSet<BomItem> BomItems { get; set; }
        public DbSet<CutTemplate> CutTemplates { get; set; }
        public DbSet<FormProgram> FormPrograms { get; set; }

        public FabWorksDbContext(DbContextOptions<FabWorksDbContext> options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<ExportRecord>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.DrawingNumber).HasMaxLength(100);
                entity.Property(e => e.SourceFilePath).HasMaxLength(500);
                entity.Property(e => e.OutputFolder).HasMaxLength(500);
                entity.Property(e => e.ExportedBy).HasMaxLength(100);
                entity.Property(e => e.PdfContentHash).HasMaxLength(64);

                entity.HasMany(e => e.BomItems)
                      .WithOne(b => b.ExportRecord)
                      .HasForeignKey(b => b.ExportRecordId)
                      .OnDelete(DeleteBehavior.Cascade);
            });

            modelBuilder.Entity<BomItem>(entity =>
            {
                entity.HasKey(e => e.ID);
                entity.Property(e => e.ItemNo).HasMaxLength(50);
                entity.Property(e => e.PartNo).HasMaxLength(100);
                entity.Property(e => e.Description).HasMaxLength(500);
                entity.Property(e => e.PartName).HasMaxLength(200);
                entity.Property(e => e.ConfigurationName).HasMaxLength(100);
                entity.Property(e => e.Material).HasMaxLength(100);

                entity.HasOne(e => e.CutTemplate)
                      .WithOne(ct => ct.BomItem)
                      .HasForeignKey<CutTemplate>(ct => ct.BomItemId)
                      .OnDelete(DeleteBehavior.Cascade);

                entity.HasOne(e => e.FormProgram)
                      .WithOne(fp => fp.BomItem)
                      .HasForeignKey<FormProgram>(fp => fp.BomItemId)
                      .OnDelete(DeleteBehavior.Cascade);
            });

            modelBuilder.Entity<CutTemplate>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.DxfFilePath).HasMaxLength(500);
                entity.Property(e => e.CutTemplateName).HasMaxLength(100);
                entity.Property(e => e.ContentHash).HasMaxLength(64);
            });

            modelBuilder.Entity<FormProgram>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.ProgramFilePath).HasMaxLength(500);
                entity.Property(e => e.ContentHash).HasMaxLength(64);
                entity.Property(e => e.ProgramName).HasMaxLength(200);
                entity.Property(e => e.MaterialType).HasMaxLength(50);
                entity.Property(e => e.UpperToolNames).HasMaxLength(500);
                entity.Property(e => e.LowerToolNames).HasMaxLength(500);
                entity.Property(e => e.SetupNotes).HasMaxLength(2000);
            });
        }
    }
}
