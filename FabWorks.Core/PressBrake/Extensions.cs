using System;
using System.Xml.Linq;

namespace FabWorks.Core.PressBrake
{
    internal static class Extensions
    {
        private static bool? ToBool(this string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return null;

            int intValue;

            if (!int.TryParse(s, out intValue))
                return null;

            return Convert.ToBoolean(intValue);
        }

        public static bool ToBool(this XAttribute a, bool defaultValue = false)
        {
            if (a == null)
                return defaultValue;

            var b = a.Value.ToBool();

            return b != null ? b.Value : defaultValue;
        }

        public static bool? ToBoolOrNull(this XAttribute a)
        {
            if (a == null)
                return null;

            return a.Value.ToBool();
        }

        private static int? ToInt(this string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return null;

            int intValue;

            if (!int.TryParse(s, out intValue))
                return null;

            return intValue;
        }

        public static int ToInt(this XAttribute a, int defaultValue = 0)
        {
            if (a == null)
                return defaultValue;

            var b = a.Value.ToInt();

            return b != null ? b.Value : defaultValue;
        }

        public static int? ToIntOrNull(this XAttribute a)
        {
            if (a == null)
                return null;

            return a.Value.ToInt();
        }

        public static int ToInt(this XElement a, int defaultValue = 0)
        {
            if (a == null)
                return defaultValue;

            var b = a.Value.ToInt();

            return b != null ? b.Value : defaultValue;
        }

        public static int? ToIntOrNull(this XElement a)
        {
            if (a == null)
                return null;

            return a.Value.ToInt();
        }

        private static double? ToDouble(this string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return null;

            double d;

            if (!double.TryParse(s, out d))
                return null;

            return d;
        }

        public static double ToDouble(this XAttribute a, double defaultValue = 0)
        {
            if (a == null)
                return defaultValue;

            var b = a.Value.ToDouble();

            return b != null ? b.Value : defaultValue;
        }

        public static double? ToDoubleOrNull(this XAttribute a)
        {
            if (a == null)
                return null;

            return a.Value.ToDouble();
        }

        public static double ToDouble(this XElement a, double defaultValue = 0)
        {
            if (a == null)
                return defaultValue;

            var b = a.Value.ToDouble();

            return b != null ? b.Value : defaultValue;
        }

        public static double? ToDoubleOrNull(this XElement a)
        {
            if (a == null)
                return null;

            return a.Value.ToDouble();
        }

        public static DateTime? ToDateTime(this XAttribute a)
        {
            if (a == null || string.IsNullOrWhiteSpace(a.Value))
                return null;

            DateTime d;

            if (!DateTime.TryParse(a.Value, out d))
                return null;

            return d;
        }

        public static TimeSpan? ToTimeSpan(this XElement e)
        {
            if (e == null || string.IsNullOrWhiteSpace(e.Value))
                return null;

            TimeSpan d;

            if (!TimeSpan.TryParse(e.Value, out d))
                return null;

            return d;
        }

        public static DateTime RoundDown(this DateTime dt, TimeSpan d)
        {
            var modTicks = dt.Ticks % d.Ticks;
            var delta = -modTicks;

            return new DateTime(dt.Ticks + delta, dt.Kind);
        }
    }
}
