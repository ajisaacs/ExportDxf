using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExportDXF.Models;

namespace ExportDXF.Services
{
    public class ExcelExportService
    {
        /// <summary>
        /// Reads existing Cut Templates from an xlsx file to compare content hashes.
        /// Returns empty dictionary if file doesn't exist or has no Cut Templates sheet.
        /// Key = Item #, Value = (ContentHash, Revision, FileName)
        /// </summary>
        public Dictionary<string, (string ContentHash, int Revision, string FileName)> ReadExistingCutTemplates(string xlsxPath)
        {
            var result = new Dictionary<string, (string, int, string)>();

            if (!File.Exists(xlsxPath))
                return result;

            using (var workbook = new XLWorkbook(xlsxPath))
            {
                if (!workbook.TryGetWorksheet("Cut Templates", out var ws))
                    return result;

                var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

                if (lastCol == 0 || lastRow <= 1)
                    return result;

                var headers = new Dictionary<string, int>();
                for (int col = 1; col <= lastCol; col++)
                {
                    var header = ws.Cell(1, col).GetString();
                    if (!string.IsNullOrEmpty(header))
                        headers[header] = col;
                }

                if (!headers.ContainsKey("Item #") || !headers.ContainsKey("Content Hash"))
                    return result;

                for (int row = 2; row <= lastRow; row++)
                {
                    var itemNo = ws.Cell(row, headers["Item #"]).GetString();
                    var hash = ws.Cell(row, headers["Content Hash"]).GetString();
                    var revision = headers.ContainsKey("Revision")
                        ? ws.Cell(row, headers["Revision"]).GetValue<int>()
                        : 1;
                    var fileName = headers.ContainsKey("File Name")
                        ? ws.Cell(row, headers["File Name"]).GetString()
                        : "";

                    if (!string.IsNullOrEmpty(itemNo))
                        result[itemNo] = (hash, revision, fileName);
                }
            }

            return result;
        }

        /// <summary>
        /// Writes or updates the xlsx file with BOM and Cut Templates sheets.
        /// rawBomTable: list of rows where each row is a dictionary of column name → value.
        /// If null or empty, the BOM sheet is not written (Part/Assembly exports).
        /// </summary>
        public void Write(
            string xlsxPath,
            List<Dictionary<string, string>> rawBomTable,
            List<BomItem> bomItems)
        {
            using (var workbook = File.Exists(xlsxPath)
                ? new XLWorkbook(xlsxPath)
                : new XLWorkbook())
            {
                WriteBomSheet(workbook, rawBomTable);
                WriteCutTemplatesSheet(workbook, bomItems);
                workbook.SaveAs(xlsxPath);
            }
        }

        private void WriteBomSheet(XLWorkbook workbook, List<Dictionary<string, string>> rawBomTable)
        {
            if (rawBomTable == null || rawBomTable.Count == 0)
                return;

            if (workbook.TryGetWorksheet("BOM", out _))
                workbook.Worksheets.Delete("BOM");

            var sheet = workbook.Worksheets.Add("BOM");
            var columns = rawBomTable[0].Keys.ToList();

            for (int col = 0; col < columns.Count; col++)
                sheet.Cell(1, col + 1).Value = columns[col];

            for (int row = 0; row < rawBomTable.Count; row++)
            {
                for (int col = 0; col < columns.Count; col++)
                {
                    string value;
                    rawBomTable[row].TryGetValue(columns[col], out value);
                    sheet.Cell(row + 2, col + 1).Value = value ?? "";
                }
            }

            sheet.Columns().AdjustToContents();
        }

        private void WriteCutTemplatesSheet(XLWorkbook workbook, List<BomItem> bomItems)
        {
            if (workbook.TryGetWorksheet("Cut Templates", out _))
                workbook.Worksheets.Delete("Cut Templates");

            var sheet = workbook.Worksheets.Add("Cut Templates");

            var headers = new[] { "Item #", "File Name", "Revision", "Thickness", "K-Factor", "Bend Radius", "Content Hash" };
            for (int col = 0; col < headers.Length; col++)
                sheet.Cell(1, col + 1).Value = headers[col];

            int row = 2;
            foreach (var item in bomItems.Where(b => b.CutTemplate != null).OrderBy(b => b.ItemNo))
            {
                var ct = item.CutTemplate;
                sheet.Cell(row, 1).Value = item.ItemNo;
                sheet.Cell(row, 2).Value = ct.DxfFilePath;
                sheet.Cell(row, 3).Value = ct.Revision;
                sheet.Cell(row, 4).Value = ct.Thickness ?? 0;
                sheet.Cell(row, 5).Value = ct.KFactor ?? 0;
                sheet.Cell(row, 6).Value = ct.DefaultBendRadius ?? 0;
                sheet.Cell(row, 7).Value = ct.ContentHash ?? "";
                row++;
            }

            sheet.Columns().AdjustToContents();
        }
    }
}
