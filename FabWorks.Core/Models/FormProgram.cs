namespace FabWorks.Core.Models
{
    public class FormProgram
    {
        public int Id { get; set; }
        public string ProgramFilePath { get; set; } = "";
        public string ContentHash { get; set; }
        public string ProgramName { get; set; } = "";
        public double? Thickness { get; set; }
        public string MaterialType { get; set; } = "";
        public double? KFactor { get; set; }
        public int BendCount { get; set; }
        public string UpperToolNames { get; set; } = "";
        public string LowerToolNames { get; set; } = "";
        public string SetupNotes { get; set; } = "";

        public int BomItemId { get; set; }
        public virtual BomItem BomItem { get; set; }
    }
}
