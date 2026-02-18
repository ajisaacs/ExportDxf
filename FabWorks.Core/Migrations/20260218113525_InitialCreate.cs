using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace FabWorks.Core.Migrations
{
    /// <inheritdoc />
    public partial class InitialCreate : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            // NOTE: ExportRecords, BomItems, and CutTemplates tables already exist
            // in the database (created by ExportDXF's ExportDxfDbContext migrations).
            // This migration only adds the new FormPrograms table.

            migrationBuilder.CreateTable(
                name: "FormPrograms",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProgramFilePath = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    ContentHash = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: true),
                    ProgramName = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    Thickness = table.Column<double>(type: "float", nullable: true),
                    MaterialType = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    KFactor = table.Column<double>(type: "float", nullable: true),
                    BendCount = table.Column<int>(type: "int", nullable: false),
                    UpperToolNames = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    LowerToolNames = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    SetupNotes = table.Column<string>(type: "nvarchar(2000)", maxLength: 2000, nullable: true),
                    BomItemId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_FormPrograms", x => x.Id);
                    table.ForeignKey(
                        name: "FK_FormPrograms_BomItems_BomItemId",
                        column: x => x.BomItemId,
                        principalTable: "BomItems",
                        principalColumn: "ID",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_FormPrograms_BomItemId",
                table: "FormPrograms",
                column: "BomItemId",
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "FormPrograms");
        }
    }
}
