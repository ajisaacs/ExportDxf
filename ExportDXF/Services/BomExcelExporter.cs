using ExportDXF.Utilities;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ExportDXF.Services
{
    /// <summary>
    /// Service for creating BOM Excel files.
    /// </summary>
    public interface IBomExcelExporter
    {
        /// <summary>
        /// Creates an Excel file containing the Bill of Materials.
        /// </summary>
        /// <param name="filepath">The full path where the Excel file will be saved.</param>
        /// <param name="items">The list of items to include in the BOM.</param>
        void CreateBOMExcelFile(string filepath, IList<Item> items);
    }

    /// <summary>
    /// Service for creating Excel files containing Bill of Materials data.
    /// </summary>
    public class BomExcelExporter : IBomExcelExporter
    {
        private const string DEFAULT_TEMPLATE_FILENAME = "BomTemplate.xlsx";
        private const string DEFAULT_SHEET_NAME = "Parts";
        private const int HEADER_ROW = 1;
        private const int DATA_START_ROW = 2;
        private const double COLUMN_PADDING = 2.0;
        private const int MAX_COLUMN_WIDTH = 50;

        private readonly string _templatePath;
        private readonly BomExcelSettings _settings;

        public BomExcelExporter() : this(GetDefaultTemplatePath(), new BomExcelSettings())
        {
        }

        public BomExcelExporter(string templatePath) : this(templatePath, new BomExcelSettings())
        {
        }

        public BomExcelExporter(string templatePath, BomExcelSettings settings)
        {
            if (string.IsNullOrWhiteSpace(templatePath))
                throw new ArgumentException("Template path cannot be null or empty.", nameof(templatePath));

            _templatePath = templatePath;
            _settings = settings ?? new BomExcelSettings();

            // Set EPPlus license context (required for newer versions)
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Creates an Excel file containing the Bill of Materials.
        /// </summary>
        public void CreateBOMExcelFile(string filepath, IList<Item> items)
        {
            if (string.IsNullOrWhiteSpace(filepath))
                throw new ArgumentException("Filepath cannot be null or empty.", nameof(filepath));

            if (items == null)
                throw new ArgumentNullException(nameof(items));

            ValidateAndPrepareFile(filepath);

            using (var package = new ExcelPackage(new FileInfo(filepath)))
            {
                var worksheet = GetOrCreateWorksheet(package);

                PopulateWorksheet(worksheet, items);
                FormatWorksheet(worksheet);

                package.Save();
            }
        }

        #region Worksheet Management

        private ExcelWorksheet GetOrCreateWorksheet(ExcelPackage package)
        {
            var worksheet = package.Workbook.Worksheets[DEFAULT_SHEET_NAME];

            if (worksheet == null)
            {
                if (_settings.CreateWorksheetIfMissing)
                {
                    worksheet = package.Workbook.Worksheets.Add(DEFAULT_SHEET_NAME);
                    CreateDefaultHeaders(worksheet);
                }
                else
                {
                    var availableSheets = string.Join(", ",
                        package.Workbook.Worksheets.Select(ws => ws.Name));

                    throw new InvalidOperationException(
                        $"Worksheet '{DEFAULT_SHEET_NAME}' not found in template. " +
                        $"Available sheets: {availableSheets}");
                }
            }

            return worksheet;
        }

        private void CreateDefaultHeaders(ExcelWorksheet worksheet)
        {
            var headers = new[]
            {
            "Item No", "File Name", "Qty", "Description", "Part Name",
            "Configuration", "Thickness", "Material", "K-Factor", "Bend Radius"
        };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cells[HEADER_ROW, i + 1];
                cell.Value = headers[i];

                if (_settings.FormatHeaders)
                {
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
            }
        }

        #endregion

        #region Data Population

        private void PopulateWorksheet(ExcelWorksheet worksheet, IList<Item> items)
        {
            var filteredItems = _settings.SkipItemsWithoutFiles
                ? items.Where(i => !string.IsNullOrEmpty(i.FileName)).ToList()
                : items.ToList();

            for (int i = 0; i < filteredItems.Count; i++)
            {
                var item = filteredItems[i];
                var row = DATA_START_ROW + i;

                WriteItemToRow(worksheet, row, item);
            }

            // Add summary information if enabled
            if (_settings.AddSummary && filteredItems.Count > 0)
            {
                AddSummarySection(worksheet, filteredItems, DATA_START_ROW + filteredItems.Count + 2);
            }
        }

        private void WriteItemToRow(ExcelWorksheet worksheet, int row, Item item)
        {
            int col = 1;

            // Item No
            worksheet.Cells[row, col++].Value = FormatItemNumber(item.ItemNo);

            // File Name
            worksheet.Cells[row, col++].Value = item.FileName;

            // Quantity
            worksheet.Cells[row, col++].Value = item.Quantity;

            // Description
            worksheet.Cells[row, col++].Value = CleanDescription(item.Description);

            // Part Name
            worksheet.Cells[row, col++].Value = item.PartName;

            // Configuration
            worksheet.Cells[row, col++].Value = item.Configuration;

            // Thickness (only if > 0)
            worksheet.Cells[row, col++].Value = GetFormattedNumericValue(item.Thickness, _settings.ThicknessDecimalPlaces);

            // Material
            worksheet.Cells[row, col++].Value = item.Material;

            // K-Factor (only if > 0)
            worksheet.Cells[row, col++].Value = GetFormattedNumericValue(item.KFactor, _settings.KFactorDecimalPlaces);

            // Bend Radius (only if > 0)
            worksheet.Cells[row, col++].Value = GetFormattedNumericValue(item.BendRadius, _settings.BendRadiusDecimalPlaces);

            // Apply row formatting
            if (_settings.AlternateRowColors && row % 2 == 0)
            {
                var rowRange = worksheet.Cells[row, 1, row, col - 1];
                rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rowRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 240, 240));
            }
        }

        private void AddSummarySection(ExcelWorksheet worksheet, IList<Item> items, int startRow)
        {
            worksheet.Cells[startRow, 1].Value = "SUMMARY";
            worksheet.Cells[startRow, 1].Style.Font.Bold = true;
            worksheet.Cells[startRow, 1].Style.Font.Size = 12;

            startRow += 2;

            // Total items
            worksheet.Cells[startRow, 1].Value = "Total Items:";
            worksheet.Cells[startRow, 2].Value = items.Count;

            // Total quantity
            var totalQty = items.Sum(i => i.Quantity);
            worksheet.Cells[startRow + 1, 1].Value = "Total Quantity:";
            worksheet.Cells[startRow + 1, 2].Value = totalQty;

            // Unique materials
            var uniqueMaterials = items.Where(i => !string.IsNullOrEmpty(i.Material))
                                        .Select(i => i.Material)
                                        .Distinct()
                                        .Count();

            worksheet.Cells[startRow + 2, 1].Value = "Unique Materials:";
            worksheet.Cells[startRow + 2, 2].Value = uniqueMaterials;

            // Thickness range
            var thicknesses = items.Where(i => i.Thickness > 0).Select(i => i.Thickness).ToList();
            if (thicknesses.Any())
            {
                worksheet.Cells[startRow + 3, 1].Value = "Thickness Range:";
                worksheet.Cells[startRow + 3, 2].Value = $"{thicknesses.Min():F1} - {thicknesses.Max():F1} mm";
            }
        }

        #endregion

        #region Data Formatting

        private string FormatItemNumber(string itemNo)
        {
            if (string.IsNullOrWhiteSpace(itemNo))
                return string.Empty;

            // Try to parse as number and pad with zeros if it's a simple number
            if (int.TryParse(itemNo, out int number))
            {
                return number.ToString().PadLeft(_settings.ItemNumberPadding, '0');
            }

            return itemNo;
        }

        private string CleanDescription(string description)
        {
            if (string.IsNullOrWhiteSpace(description))
                return string.Empty;

            // Remove any remaining XML tags that might have been missed
            description = TextHelper.RemoveXmlTags(description);

            // Trim and normalize whitespace
            return System.Text.RegularExpressions.Regex.Replace(description.Trim(), @"\s+", " ");
        }

        private object GetFormattedNumericValue(double value, int decimalPlaces)
        {
            if (value <= 0)
                return null;

            return Math.Round(value, decimalPlaces);
        }

        #endregion

        #region Worksheet Formatting

        private void FormatWorksheet(ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null)
                return;

            // Auto-fit columns with limits
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                worksheet.Column(col).AutoFit();

                // Special handling for Description column (assuming it's column 4)
                if (col == 4) // Description column
                {
                    worksheet.Column(col).Width = Math.Max(worksheet.Column(col).Width, 30); // Minimum 30 units

                    if (worksheet.Column(col).Width > 50) // Maximum 50 units
                        worksheet.Column(col).Width = 50;
                }
                else if (worksheet.Column(col).Width > MAX_COLUMN_WIDTH)
                {
                    worksheet.Column(col).Width = MAX_COLUMN_WIDTH;
                }
                else
                {
                    worksheet.Column(col).Width += COLUMN_PADDING;
                }
            }

            // Add borders if enabled
            if (_settings.AddBorders)
            {
                var dataRange = worksheet.Cells[HEADER_ROW, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            // Format numeric columns
            FormatNumericColumns(worksheet);

            // Freeze header row if enabled
            if (_settings.FreezeHeaderRow)
            {
                worksheet.View.FreezePanes(DATA_START_ROW, 1);
            }
        }

        private void FormatNumericColumns(ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null)
                return;

            var lastRow = worksheet.Dimension.End.Row;

            // Format quantity column (assuming it's column 3)
            if (worksheet.Dimension.Columns >= 3)
            {
                var qtyRange = worksheet.Cells[DATA_START_ROW, 3, lastRow, 3];
                qtyRange.Style.Numberformat.Format = "0";
                qtyRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            // Format thickness column (assuming it's column 7)
            if (worksheet.Dimension.Columns >= 7)
            {
                var thicknessRange = worksheet.Cells[DATA_START_ROW, 7, lastRow, 7];
                thicknessRange.Style.Numberformat.Format = $"0.{new string('0', _settings.ThicknessDecimalPlaces)}";
                thicknessRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }
        }

        #endregion

        #region File Management

        private void ValidateAndPrepareFile(string filepath)
        {
            var directory = Path.GetDirectoryName(filepath);

            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (UseTemplate())
            {
                CopyTemplateToDestination(filepath);
            }
            else if (File.Exists(filepath))
            {
                File.Delete(filepath); // Remove existing file to start fresh
            }
        }

        private bool UseTemplate()
        {
            return !string.IsNullOrEmpty(_templatePath) && File.Exists(_templatePath);
        }

        private void CopyTemplateToDestination(string filepath)
        {
            if (!File.Exists(_templatePath))
            {
                throw new FileNotFoundException(
                    $"BOM template file not found at: {_templatePath}. " +
                    "Either provide a valid template or enable CreateWorksheetIfMissing.",
                    _templatePath);
            }

            try
            {
                File.Copy(_templatePath, filepath, overwrite: true);
            }
            catch (Exception ex)
            {
                throw new IOException($"Failed to copy template to {filepath}: {ex.Message}", ex);
            }
        }

        private static string GetDefaultTemplatePath()
        {
            return Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Templates",
                DEFAULT_TEMPLATE_FILENAME);
        }

        #endregion
    }
}
