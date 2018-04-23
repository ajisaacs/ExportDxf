using SolidWorks.Interop.sldworks;

namespace ExportDXF
{
	public class Item
    {
		public string ItemNo { get; set; }

		public string PartNo { get; set; }

        public int Quantity { get; set; }

		public string Description { get; set; }

		public Component2 Component { get; set; }
    }
}
