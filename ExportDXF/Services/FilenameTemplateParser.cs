using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExportDXF.Services
{
    public static class FilenameTemplateParser
    {
        private static readonly Regex PlaceholderPattern = new Regex(
            @"\{(?<name>\w+)(?::(?<pad>\d+))?\}",
            RegexOptions.Compiled);

        /// <summary>
        /// Evaluates a template string for a given item.
        /// e.g. "4321 A01 PT{item_no:2}" with item_no=3 → "4321 A01 PT03"
        /// </summary>
        public static string Evaluate(string template, Item item)
        {
            return PlaceholderPattern.Replace(template, match =>
            {
                var name = match.Groups["name"].Value.ToLowerInvariant();
                var padStr = match.Groups["pad"].Value;
                int pad = string.IsNullOrEmpty(padStr) ? 0 : int.Parse(padStr);

                string value;
                switch (name)
                {
                    case "item_no":
                        value = item.ItemNo ?? "0";
                        if (pad > 0)
                            value = value.PadLeft(pad, '0');
                        break;
                    case "part_name":
                        value = item.PartName ?? "";
                        break;
                    case "config":
                        value = item.Configuration ?? "";
                        break;
                    case "material":
                        value = item.Material ?? "";
                        break;
                    default:
                        value = match.Value;
                        break;
                }

                return value;
            });
        }

        /// <summary>
        /// Extracts the literal prefix before the first placeholder.
        /// Used for naming the xlsx and log files.
        /// Falls back to documentName if prefix is empty.
        /// </summary>
        public static string GetPrefix(string template, string documentName)
        {
            var match = PlaceholderPattern.Match(template);
            if (!match.Success)
                return template.Trim();

            var prefix = template.Substring(0, match.Index).Trim();
            if (string.IsNullOrEmpty(prefix))
                return Path.GetFileNameWithoutExtension(documentName);

            return prefix;
        }

        /// <summary>
        /// Validates that the template contains {item_no} to prevent filename collisions.
        /// </summary>
        public static bool Validate(string template, out string error)
        {
            error = null;

            if (string.IsNullOrWhiteSpace(template))
            {
                error = "Template cannot be empty.";
                return false;
            }

            var hasItemNo = PlaceholderPattern.Matches(template)
                .Cast<Match>()
                .Any(m => m.Groups["name"].Value.Equals("item_no", StringComparison.OrdinalIgnoreCase));

            if (!hasItemNo)
            {
                error = "Template must contain {item_no} or {item_no:N} to avoid filename collisions.";
                return false;
            }

            return true;
        }
    }
}
