using SolidWorks.Interop.sldworks;

namespace ExportDXF
{
	public class Item
    {
        public string Name { get; set; }

        public int Quantity { get; set; }

        public Component2 Component { get; set; }
    }
}
