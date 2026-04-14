using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExportDXF.Utilities
{
    public static class ContentHasher
    {
        /// <summary>
        /// Computes a SHA256 hash of DXF file content, skipping the HEADER section
        /// which contains timestamps that change on every save.
        /// </summary>
        public static string ComputeDxfContentHash(string filePath)
        {
            var text = File.ReadAllText(filePath);
            var contentStart = FindEndOfHeader(text);
            var content = contentStart >= 0 ? text.Substring(contentStart) : text;

            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(content));
                return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Computes a SHA256 hash of the entire file contents (for PDFs and other binary files).
        /// </summary>
        public static string ComputeFileHash(string filePath)
        {
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                var bytes = sha.ComputeHash(stream);
                return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Finds the position immediately after the HEADER section's ENDSEC marker.
        /// DXF HEADER format:
        ///   0\nSECTION\n2\nHEADER\n...variables...\n0\nENDSEC\n
        /// Returns -1 if no HEADER section is found.
        /// </summary>
        private static int FindEndOfHeader(string text)
        {
            // Find the HEADER section start
            var headerIndex = FindGroupCode(text, 0, "2", "HEADER");
            if (headerIndex < 0)
                return -1;

            // Advance past the HEADER value line so pair scanning stays aligned
            var headerLineEnd = text.IndexOf('\n', headerIndex);
            if (headerLineEnd < 0)
                return -1;

            // Find the ENDSEC that closes the HEADER section
            var pos = headerLineEnd + 1;
            while (pos < text.Length)
            {
                var endsecIndex = FindGroupCode(text, pos, "0", "ENDSEC");
                if (endsecIndex < 0)
                    return -1;

                // Move past the ENDSEC line
                var lineEnd = text.IndexOf('\n', endsecIndex);
                return lineEnd >= 0 ? lineEnd + 1 : text.Length;
            }

            return -1;
        }

        /// <summary>
        /// Finds a DXF group code pair (code line followed by value line) starting from the given position.
        /// Returns the position of the value line, or -1 if not found.
        /// </summary>
        private static int FindGroupCode(string text, int startIndex, string groupCode, string value)
        {
            var pos = startIndex;
            while (pos < text.Length)
            {
                // Skip whitespace/newlines to find the group code
                while (pos < text.Length && (text[pos] == '\r' || text[pos] == '\n' || text[pos] == ' '))
                    pos++;

                if (pos >= text.Length)
                    break;

                // Read the group code line
                var codeLineEnd = text.IndexOf('\n', pos);
                if (codeLineEnd < 0)
                    break;

                var codeLine = text.Substring(pos, codeLineEnd - pos).Trim();

                // Move to the value line
                var valueStart = codeLineEnd + 1;
                if (valueStart >= text.Length)
                    break;

                var valueLineEnd = text.IndexOf('\n', valueStart);
                if (valueLineEnd < 0)
                    valueLineEnd = text.Length;

                var valueLine = text.Substring(valueStart, valueLineEnd - valueStart).Trim();

                if (codeLine == groupCode && string.Equals(valueLine, value, StringComparison.OrdinalIgnoreCase))
                    return valueStart;

                // Move to the next pair
                pos = valueLineEnd + 1;
            }

            return -1;
        }
    }
}
