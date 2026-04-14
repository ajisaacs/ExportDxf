using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace FabWorks.Core.Migrations
{
    /// <inheritdoc />
    public partial class AddCutTemplateRevision : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "Revision",
                table: "CutTemplates",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Revision",
                table: "CutTemplates");
        }
    }
}
