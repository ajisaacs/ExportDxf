using ExportDXF.Models;
using Microsoft.EntityFrameworkCore;
using System.Configuration;

namespace ExportDXF.Data
{
    public class ExportDxfDbContext : DbContext
    {
        public DbSet<ExportRecord> ExportRecords { get; set; }
        public DbSet<BomItem> BomItems { get; set; }
        public DbSet<CutTemplate> CutTemplates { get; set; }

        public ExportDxfDbContext() : base()
        {
        }

        public ExportDxfDbContext(DbContextOptions<ExportDxfDbContext> options) : base(options)
        {
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                var connectionString = ConfigurationManager.ConnectionStrings["ExportDxfDb"]?.ConnectionString
                    ?? "Server=localhost;Database=ExportDxfDb;Trusted_Connection=True;TrustServerCertificate=True;";
                optionsBuilder.UseSqlServer(connectionString);
            }
        }

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
            });

            modelBuilder.Entity<CutTemplate>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.CutTemplateName).HasMaxLength(100);
                entity.Property(e => e.ContentHash).HasMaxLength(64);
            });
        }
    }
}
