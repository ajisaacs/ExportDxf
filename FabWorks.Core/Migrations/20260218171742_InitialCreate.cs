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
            migrationBuilder.CreateTable(
                name: "ExportRecords",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DrawingNumber = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Title = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    EquipmentNo = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    DrawingNo = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    SourceFilePath = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    OutputFolder = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    ExportedAt = table.Column<DateTime>(type: "datetime2", nullable: false),
                    ExportedBy = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    PdfContentHash = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ExportRecords", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BomItems",
                columns: table => new
                {
                    ID = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ItemNo = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    PartNo = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    SortOrder = table.Column<int>(type: "int", nullable: false),
                    Qty = table.Column<int>(type: "int", nullable: true),
                    TotalQty = table.Column<int>(type: "int", nullable: true),
                    Description = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    PartName = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    ConfigurationName = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Material = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    ExportRecordId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BomItems", x => x.ID);
                    table.ForeignKey(
                        name: "FK_BomItems_ExportRecords_ExportRecordId",
                        column: x => x.ExportRecordId,
                        principalTable: "ExportRecords",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "CutTemplates",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DxfFilePath = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    ContentHash = table.Column<string>(type: "nvarchar(64)", maxLength: 64, nullable: true),
                    CutTemplateName = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Thickness = table.Column<double>(type: "float", nullable: true),
                    KFactor = table.Column<double>(type: "float", nullable: true),
                    DefaultBendRadius = table.Column<double>(type: "float", nullable: true),
                    BomItemId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_CutTemplates", x => x.Id);
                    table.ForeignKey(
                        name: "FK_CutTemplates_BomItems_BomItemId",
                        column: x => x.BomItemId,
                        principalTable: "BomItems",
                        principalColumn: "ID",
                        onDelete: ReferentialAction.Cascade);
                });

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
                name: "IX_BomItems_ExportRecordId",
                table: "BomItems",
                column: "ExportRecordId");

            migrationBuilder.CreateIndex(
                name: "IX_CutTemplates_BomItemId",
                table: "CutTemplates",
                column: "BomItemId",
                unique: true);

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
                name: "CutTemplates");

            migrationBuilder.DropTable(
                name: "FormPrograms");

            migrationBuilder.DropTable(
                name: "BomItems");

            migrationBuilder.DropTable(
                name: "ExportRecords");
        }
    }
}
