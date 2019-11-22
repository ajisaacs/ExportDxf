namespace ExportDXF
{
    internal class Bend
    {
        public double BendLineAngle { get; set; }

        public BendDirection Direction { get; set; }

        public double Angle { get; set; }

        /// <summary>
        /// X coordinate of the bend line in meters
        /// </summary>
        public double X { get; set; }

        /// <summary>
        /// Y coordinate of the bend line in meters
        /// </summary>
        public double Y { get; set; }
    }
}