namespace ExportDXF.Extensions
{
    /// <summary>
    /// Extension methods for unit conversion between SolidWorks (meters) and millimeters.
    /// </summary>
    public static class UnitConversionExtensions
    {
        private const double METERS_TO_MM = .0001;
        private const double METERS_TO_INCHES = 0.0254;

        /// <summary>
        /// Converts a SolidWorks dimension (in meters) to millimeters.
        /// </summary>
        /// <param name="meters">The value in meters.</param>
        /// <returns>The value in millimeters.</returns>
        public static double FromSolidWorksToMM(this double meters)
        {
            return meters * METERS_TO_MM;
        }

        /// <summary>
        /// Converts millimeters to SolidWorks dimension (meters).
        /// </summary>
        /// <param name="millimeters">The value in millimeters.</param>
        /// <returns>The value in meters.</returns>
        public static double FromMMToSolidWorks(this double millimeters)
        {
            return millimeters / METERS_TO_MM;
        }

        public static double FromSolidWorksToInches(this double meters)
        {
            return meters * METERS_TO_INCHES;
        }

        public static double FromInchesToSolidWorks(this double inches)
        {
            return inches / METERS_TO_INCHES;
        }
    }
}
