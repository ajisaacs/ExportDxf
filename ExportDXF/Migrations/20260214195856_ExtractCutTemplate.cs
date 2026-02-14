using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ExportDXF.Migrations
{
    /// <inheritdoc />
    public partial class ExtractCutTemplate : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "ContentHash",
                table: "BomItems");

            migrationBuilder.DropColumn(
                name: "CutTemplateName",
                table: "BomItems");

            migrationBuilder.DropColumn(
                name: "DefaultBendRadius",
                table: "BomItems");

            migrationBuilder.DropColumn(
                name: "DxfFilePath",
                table: "BomItems");

            migrationBuilder.DropColumn(
                name: "KFactor",
                table: "BomItems");

            migrationBuilder.DropColumn(
                name: "Thickness",
                table: "BomItems");

            migrationBuilder.CreateTable(
                name: "CutTemplates",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DxfFilePath = table.Column<string>(type: "nvarchar(max)", nullable: true),
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

            migrationBuilder.CreateIndex(
                name: "IX_CutTemplates_BomItemId",
                table: "CutTemplates",
                column: "BomItemId",
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "CutTemplates");

            migrationBuilder.AddColumn<string>(
                name: "ContentHash",
                table: "BomItems",
                type: "nvarchar(64)",
                maxLength: 64,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "CutTemplateName",
                table: "BomItems",
                type: "nvarchar(100)",
                maxLength: 100,
                nullable: true);

            migrationBuilder.AddColumn<double>(
                name: "DefaultBendRadius",
                table: "BomItems",
                type: "float",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "DxfFilePath",
                table: "BomItems",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<double>(
                name: "KFactor",
                table: "BomItems",
                type: "float",
                nullable: true);

            migrationBuilder.AddColumn<double>(
                name: "Thickness",
                table: "BomItems",
                type: "float",
                nullable: true);
        }
    }
}
