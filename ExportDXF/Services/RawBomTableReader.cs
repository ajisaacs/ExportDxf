using System.Collections.Generic;
using SolidWorks.Interop.sldworks;

namespace ExportDXF.Services
{
    /// <summary>
    /// Reads all visible columns and rows from a SolidWorks BOM table annotation
    /// as raw string data for direct copy into an Excel sheet.
    /// </summary>
    public static class RawBomTableReader
    {
        public static List<Dictionary<string, string>> Read(BomTableAnnotation bomTable)
        {
            var table = (TableAnnotation)bomTable;
            var rows = new List<Dictionary<string, string>>();

            int colCount = table.ColumnCount;
            int rowCount = table.RowCount;

            // Build visible column headers
            var columns = new List<(int Index, string Header)>();
            for (int col = 0; col < colCount; col++)
            {
                if (table.ColumnHidden[col])
                    continue;

                var header = table.get_Text(0, col)?.Trim() ?? $"Column{col}";
                columns.Add((col, header));
            }

            // Read data rows (skip header row 0, skip hidden rows)
            for (int row = 1; row < rowCount; row++)
            {
                if (table.RowHidden[row])
                    continue;

                var rowData = new Dictionary<string, string>();
                foreach (var (colIdx, header) in columns)
                {
                    rowData[header] = table.get_Text(row, colIdx)?.Trim() ?? "";
                }
                rows.Add(rowData);
            }

            return rows;
        }
    }
}
