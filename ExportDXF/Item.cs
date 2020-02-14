using SolidWorks.Interop.sldworks;

namespace ExportDXF
{
    public class Item
    {
        public string ItemNo { get; set; }

        public string FileName { get; set; }

        public string PartName { get; set; }

        public string Configuration { get; set; }

        public int Quantity { get; set; }

        public string Description { get; set; }

        public double Thickness { get; set; }

        public double KFactor { get; set; }

        public double BendRadius { get; set; }

        public string Material { get; set; }

        public Component2 Component { get; set; }
    }
}