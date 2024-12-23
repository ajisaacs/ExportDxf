using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExportDXF
{
    public class BomToExcel
    {
        private const string DefaultTemplatePath = "Templates/BomTemplate.xlsx";
        private const string DefaultSheetName = "Parts";

        public string TemplatePath { get; set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DefaultTemplatePath);

        public void CreateBOMExcelFile(string filepath, IList<Item> items)
        {
            try
            {
                CopyTemplate(filepath);

                using (var pkg = new ExcelPackage(new FileInfo(filepath)))
                {
                    var partsSheet = pkg.Workbook.Worksheets[DefaultSheetName] ?? throw new Exception($"Worksheet '{DefaultSheetName}' not found in template.");

                    PopulateSheet(partsSheet, items);
                    AutoFitColumns(partsSheet);

                    pkg.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create BOM Excel file: {ex.Message}");
                throw;
            }
        }

        private void CopyTemplate(string filepath)
        {
            if (!File.Exists(TemplatePath))
                throw new FileNotFoundException("Template file not found.", TemplatePath);

            File.Copy(TemplatePath, filepath, true);
        }

        private void PopulateSheet(ExcelWorksheet sheet, IList<Item> items)
        {
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var row = i + 2; // Assuming row 1 is the header
                var col = 1;

                sheet.Cells[row, col++].Value = item.ItemNo;
                sheet.Cells[row, col++].Value = item.FileName;
                sheet.Cells[row, col++].Value = item.Quantity;
                sheet.Cells[row, col++].Value = item.Description;
                sheet.Cells[row, col++].Value = item.PartName;
                sheet.Cells[row, col++].Value = item.Configuration;
                sheet.Cells[row, col++].Value = item.Thickness > 0 ? (object)item.Thickness : null;
                sheet.Cells[row, col++].Value = item.Material;
                sheet.Cells[row, col++].Value = item.KFactor > 0 ? (object)item.KFactor : null;
                sheet.Cells[row, col++].Value = item.BendRadius > 0 ? (object)item.BendRadius : null;
            }
        }

        private void AutoFitColumns(ExcelWorksheet sheet)
        {
            for (int i = 1; i <= sheet.Dimension.Columns; i++)
            {
                var column = sheet.Column(i);
                column.AutoFit();
                column.Width += 1; // Adding padding for better visibility
            }
        }
    }
}
