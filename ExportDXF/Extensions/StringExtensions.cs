namespace ExportDXF.Extensions
{
    /// <summary>
    /// Extension methods for string manipulation.
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Extension method to remove XML tags from a string.
        /// </summary>
        public static string RemoveXmlTags(this string input)
        {
            return ExportDXF.Utilities.TextHelper.RemoveXmlTags(input);
        }

        /// <summary>
        /// Extension method to clean text (remove XML and normalize whitespace).
        /// </summary>
        public static string CleanText(this string input)
        {
            return ExportDXF.Utilities.TextHelper.CleanText(input);
        }

        /// <summary>
        /// Extension method to sanitize a filename.
        /// </summary>
        public static string SanitizeFileName(this string input)
        {
            return ExportDXF.Utilities.TextHelper.SanitizeFileName(input);
        }

        /// <summary>
        /// Extension method to truncate a string.
        /// </summary>
        public static string Truncate(this string input, int maxLength, bool useEllipsis = true)
        {
            return ExportDXF.Utilities.TextHelper.Truncate(input, maxLength, useEllipsis);
        }

        /// <summary>
        /// Extension method to convert to title case.
        /// </summary>
        public static string ToTitleCase(this string input)
        {
            return ExportDXF.Utilities.TextHelper.ToTitleCase(input);
        }
    }
}
