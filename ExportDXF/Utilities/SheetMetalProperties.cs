namespace ExportDXF.Utilities
{
    /// <summary>
    /// Contains sheet metal properties extracted from a SolidWorks part.
    /// </summary>
    public class SheetMetalProperties
    {
        /// <summary>
        /// Material thickness in millimeters.
        /// </summary>
        public double Thickness { get; set; }

        /// <summary>
        /// K-factor for bend calculations.
        /// </summary>
        public double KFactor { get; set; }

        /// <summary>
        /// Bend radius in millimeters.
        /// </summary>
        public double BendRadius { get; set; }

        /// <summary>
        /// Bend allowance in millimeters.
        /// </summary>
        public double BendAllowance { get; set; }

        /// <summary>
        /// Whether auto relief is enabled.
        /// </summary>
        public bool AutoRelief { get; set; }

        /// <summary>
        /// Relief ratio for auto relief.
        /// </summary>
        public double ReliefRatio { get; set; }

        public override string ToString()
        {
            return $"Thickness: {Thickness:F2}mm, K-Factor: {KFactor:F3}, Bend Radius: {BendRadius:F2}mm";
        }
    }
}
