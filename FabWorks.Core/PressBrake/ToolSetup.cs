using System.Collections.Generic;

namespace FabWorks.Core.PressBrake
{
    public class ToolSetup
    {
        public ToolSetup()
        {
            Segments = new List<SegEntry>();
        }

        public string Name { get; set; }

        public int Id { get; set; }

        public double Length { get; set; }

		public int StackedHolderType { get; set; }

		public double HolderHeight { get; set; }

		public List<SegEntry> Segments { get; set; }
    }
}
