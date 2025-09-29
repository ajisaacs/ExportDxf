using System;

namespace ExportDXF.Extensions
{
    /// <summary>
    /// Extension methods for TimeSpan formatting.
    /// </summary>
    public static class TimeSpanExtensions
    {
        /// <summary>
        /// Formats a TimeSpan into a human-readable string.
        /// </summary>
        /// <param name="ts">The TimeSpan to format.</param>
        /// <returns>A human-readable duration string.</returns>
        public static string ToReadableFormat(this TimeSpan ts)
        {
            if (ts.TotalHours >= 1)
            {
                var totalHours = (int)(ts.TotalDays * 24 + ts.Hours);
                return $"{totalHours}hrs {ts.Minutes}min {ts.Seconds}sec";
            }

            if (ts.TotalMinutes >= 1)
            {
                return $"{ts.Minutes}min {ts.Seconds}sec";
            }

            return $"{ts.Seconds} seconds";
        }
    }
}
