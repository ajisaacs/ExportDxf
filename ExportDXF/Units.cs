using System;

namespace ExportDXF
{
    public static class Units
    {
        /// <summary>
        /// Multiply factor needed to convert the desired units to meters.
        /// </summary>
        public static double ScaleFactor = 0.0254; // inches to meters

        public static double ToSldWorks(this double d)
        {
            return Math.Round(d * ScaleFactor, 8);
        }

        public static double FromSldWorks(this double d)
        {
            return Math.Round(d / ScaleFactor, 8);
        }
    }
}