using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExportDXF
{
    public class ExportBomToExcel
    {
        public string TemplatePath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "BomTemplate.xlsx"); }
        }

        public void CreateBOMExcelFile(string filepath, IList<Item> items)
        {
            File.Copy(TemplatePath, filepath, true);

            var newFile = new FileInfo(filepath);

            using (var pkg = new ExcelPackage(newFile))
            {
                var workbook = pkg.Workbook;
                var partsSheet = workbook.Worksheets["Parts"];

                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    var row = i + 2;
                    var col = 1;

                    partsSheet.Cells[row, col++].Value = item.ItemNo;
                    partsSheet.Cells[row, col++].Value = item.FileName;
                    partsSheet.Cells[row, col++].Value = item.Quantity;
                    partsSheet.Cells[row, col++].Value = item.Description;
                    partsSheet.Cells[row, col++].Value = item.PartName;
                    partsSheet.Cells[row, col++].Value = item.Configuration;

                    if (item.Thickness > 0)
                        partsSheet.Cells[row, col].Value = item.Thickness;
                    col++;

                    partsSheet.Cells[row, col++].Value = item.Material;

                    if (item.KFactor > 0)
                        partsSheet.Cells[row, col].Value = item.KFactor;
                    col++;

                    if (item.BendRadius > 0)
                        partsSheet.Cells[row, col].Value = item.BendRadius;
                }

                for (int i = 1; i <= 8; i++)
                {
                    var column = partsSheet.Column(i);

                    if (column.Style.WrapText)
                        continue;

                    column.AutoFit();
                    column.Width += 1;
                }

                workbook.Calculate();
                pkg.Save();
            }
        }
    }
}
