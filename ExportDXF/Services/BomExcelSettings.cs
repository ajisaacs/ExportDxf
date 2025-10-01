namespace ExportDXF.Services
{

    /// <summary>
    /// Configuration settings for BOM Excel export.
    /// </summary>
    public class BomExcelSettings
    {
        /// <summary>
        /// Create the worksheet if it doesn't exist in the template.
        /// </summary>
        public bool CreateWorksheetIfMissing { get; set; } = true;

        /// <summary>
        /// Skip items that don't have an associated DXF file.
        /// </summary>
        public bool SkipItemsWithoutFiles { get; set; } = false;

        /// <summary>
        /// Add summary section with totals and statistics.
        /// </summary>
        public bool AddSummary { get; set; } = false;

        /// <summary>
        /// Format header row with bold text and background color.
        /// </summary>
        public bool FormatHeaders { get; set; } = false;

        /// <summary>
        /// Add borders around all cells.
        /// </summary>
        public bool AddBorders { get; set; } = false;

        /// <summary>
        /// Freeze the header row for easier scrolling.
        /// </summary>
        public bool FreezeHeaderRow { get; set; } = true;

        /// <summary>
        /// Alternate row background colors for easier reading.
        /// </summary>
        public bool AlternateRowColors { get; set; } = true;

        /// <summary>
        /// Number of decimal places for thickness values.
        /// </summary>
        public int ThicknessDecimalPlaces { get; set; } = 4;

        /// <summary>
        /// Number of decimal places for K-Factor values.
        /// </summary>
        public int KFactorDecimalPlaces { get; set; } = 3;

        /// <summary>
        /// Number of decimal places for bend radius values.
        /// </summary>
        public int BendRadiusDecimalPlaces { get; set; } = 3;

        /// <summary>
        /// Minimum number of digits for item numbers (pad with zeros).
        /// </summary>
        public int ItemNumberPadding { get; set; } = 2;
    }
}
