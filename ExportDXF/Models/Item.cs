using SolidWorks.Interop.sldworks;

namespace ExportDXF.Services
{
    /// <summary>
    /// Represents an item extracted from a BOM or assembly.
    /// </summary>
    public class Item
    {
        /// <summary>
        /// Item number from the BOM.
        /// </summary>
        public string ItemNo { get; set; }

        /// <summary>
        /// Part name or file name.
        /// </summary>
        public string PartName { get; set; }

        /// <summary>
        /// Configuration name.
        /// </summary>
        public string Configuration { get; set; }

        /// <summary>
        /// Item description.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Quantity of this item.
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// Material specification.
        /// </summary>
        public string Material { get; set; }

        /// <summary>
        /// Sheet metal thickness in millimeters.
        /// </summary>
        public double Thickness { get; set; }

        /// <summary>
        /// Sheet metal K-factor.
        /// </summary>
        public double KFactor { get; set; }

        /// <summary>
        /// Bend radius in millimeters.
        /// </summary>
        public double BendRadius { get; set; }

        /// <summary>
        /// The exported DXF filename (without path or extension).
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// The SolidWorks component reference.
        /// </summary>
        public Component2 Component { get; set; }
    }
}