using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExportDXF.Utilities
{
    /// <summary>
    /// Utility class for text processing and string manipulation operations.
    /// </summary>
    public static class TextHelper
    {
        private static readonly Regex XmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.Compiled);
        private static readonly Regex WhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);
        private static readonly Regex FontTagRegex = new Regex(@"<FONT.*?>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Removes all XML tags from the input string.
        /// </summary>
        /// <param name="input">The string containing XML tags to remove.</param>
        /// <returns>The string with all XML tags removed, or the original input if null/empty.</returns>
        public static string RemoveXmlTags(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            return XmlTagRegex.Replace(input, string.Empty);
        }

        /// <summary>
        /// Removes specific SolidWorks font XML tags from the input string.
        /// This is more targeted than RemoveXmlTags and handles SolidWorks-specific formatting.
        /// </summary>
        /// <param name="input">The string containing font tags to remove.</param>
        /// <returns>The string with font tags removed.</returns>
        public static string RemoveFontXmlTags(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            var result = input;
            var matches = FontTagRegex.Matches(result);

            // Process matches in reverse order to maintain indices
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                var match = matches[i];
                result = result.Remove(match.Index, match.Length);
            }

            return result;
        }

        /// <summary>
        /// Normalizes whitespace in a string by replacing multiple consecutive whitespace characters with a single space.
        /// </summary>
        /// <param name="input">The string to normalize.</param>
        /// <returns>The string with normalized whitespace.</returns>
        public static string NormalizeWhitespace(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            return WhitespaceRegex.Replace(input.Trim(), " ");
        }

        /// <summary>
        /// Cleans text by removing XML tags and normalizing whitespace.
        /// This is a common operation for processing text from SolidWorks.
        /// </summary>
        /// <param name="input">The text to clean.</param>
        /// <returns>Cleaned text with XML tags removed and whitespace normalized.</returns>
        public static string CleanText(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            var cleaned = RemoveXmlTags(input);
            return NormalizeWhitespace(cleaned);
        }

        /// <summary>
        /// Returns a number with its ordinal suffix (1st, 2nd, 3rd, 4th, etc.).
        /// </summary>
        /// <param name="number">The number to format.</param>
        /// <returns>The number with appropriate ordinal suffix.</returns>
        public static string GetOrdinalSuffix(int number)
        {
            if (number <= 0)
                return number.ToString();

            // Special cases for 11th, 12th, 13th
            if (number >= 11 && number <= 13)
                return number + "th";

            return number + GetSuffix(number % 10);
        }

        /// <summary>
        /// Converts a string to title case (first letter of each word capitalized).
        /// </summary>
        /// <param name="input">The string to convert.</param>
        /// <returns>The string in title case.</returns>
        public static string ToTitleCase(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLowerInvariant());
        }

        /// <summary>
        /// Truncates a string to the specified maximum length, optionally adding an ellipsis.
        /// </summary>
        /// <param name="input">The string to truncate.</param>
        /// <param name="maxLength">The maximum length of the result.</param>
        /// <param name="useEllipsis">Whether to add "..." when truncating.</param>
        /// <returns>The truncated string.</returns>
        public static string Truncate(string input, int maxLength, bool useEllipsis = true)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            if (input.Length <= maxLength)
                return input;

            if (useEllipsis && maxLength > 3)
            {
                return input.Substring(0, maxLength - 3) + "...";
            }

            return input.Substring(0, maxLength);
        }

        /// <summary>
        /// Removes invalid filename characters from a string, replacing them with underscores.
        /// </summary>
        /// <param name="filename">The filename to sanitize.</param>
        /// <returns>A safe filename with invalid characters replaced.</returns>
        public static string SanitizeFileName(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                return filename;

            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(filename);

            foreach (var invalidChar in invalidChars)
            {
                sb.Replace(invalidChar, '_');
            }

            // Also replace some additional problematic characters
            sb.Replace(' ', '_');  // Spaces can be problematic
            sb.Replace('"', '\''); // Double quotes to single quotes

            return sb.ToString();
        }

        /// <summary>
        /// Checks if a string is null, empty, or contains only whitespace.
        /// </summary>
        /// <param name="input">The string to check.</param>
        /// <returns>True if the string is null, empty, or whitespace only.</returns>
        public static bool IsNullOrWhiteSpace(string input)
        {
            return string.IsNullOrWhiteSpace(input);
        }

        /// <summary>
        /// Safely gets a substring without throwing exceptions for invalid indices.
        /// </summary>
        /// <param name="input">The source string.</param>
        /// <param name="startIndex">The starting index.</param>
        /// <param name="length">The length of the substring.</param>
        /// <returns>The substring, or empty string if indices are invalid.</returns>
        public static string SafeSubstring(string input, int startIndex, int length)
        {
            if (string.IsNullOrEmpty(input) || startIndex < 0 || startIndex >= input.Length)
                return string.Empty;

            var actualLength = Math.Min(length, input.Length - startIndex);
            return actualLength <= 0 ? string.Empty : input.Substring(startIndex, actualLength);
        }

        /// <summary>
        /// Safely gets a substring from the start index to the end of the string.
        /// </summary>
        /// <param name="input">The source string.</param>
        /// <param name="startIndex">The starting index.</param>
        /// <returns>The substring from start index to end, or empty if invalid.</returns>
        public static string SafeSubstring(string input, int startIndex)
        {
            if (string.IsNullOrEmpty(input) || startIndex < 0 || startIndex >= input.Length)
                return string.Empty;

            return input.Substring(startIndex);
        }

        /// <summary>
        /// Pads a string to a specific length, truncating if too long.
        /// </summary>
        /// <param name="input">The string to pad or truncate.</param>
        /// <param name="totalLength">The desired total length.</param>
        /// <param name="paddingChar">The character to use for padding.</param>
        /// <param name="padLeft">True to pad on the left, false to pad on the right.</param>
        /// <returns>A string of exactly the specified length.</returns>
        public static string PadOrTruncate(string input, int totalLength, char paddingChar = ' ', bool padLeft = false)
        {
            if (string.IsNullOrEmpty(input))
                input = string.Empty;

            if (input.Length == totalLength)
                return input;

            if (input.Length > totalLength)
                return input.Substring(0, totalLength);

            return padLeft
                ? input.PadLeft(totalLength, paddingChar)
                : input.PadRight(totalLength, paddingChar);
        }

        /// <summary>
        /// Converts a string to a safe identifier (letters, numbers, underscores only).
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns>A safe identifier string.</returns>
        public static string ToSafeIdentifier(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "Identifier";

            var sb = new StringBuilder();

            foreach (char c in input)
            {
                if (char.IsLetterOrDigit(c))
                {
                    sb.Append(c);
                }
                else if (c == ' ' || c == '-' || c == '.')
                {
                    sb.Append('_');
                }
            }

            var result = sb.ToString();

            // Ensure it starts with a letter or underscore
            if (result.Length > 0 && char.IsDigit(result[0]))
            {
                result = "_" + result;
            }

            return string.IsNullOrEmpty(result) ? "Identifier" : result;
        }

        #region Private Helper Methods

        private static string GetSuffix(int lastDigit)
        {
            switch (lastDigit)
            {
                case 1: return "st";
                case 2: return "nd";
                case 3: return "rd";
                default: return "th";
            }
        }

        #endregion
    }
}
