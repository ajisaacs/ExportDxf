namespace FabWorks.Core.Models
{
    public class BomItem
    {
        public int ID { get; set; }
        public string ItemNo { get; set; } = "";
        public string PartNo { get; set; } = "";
        public int SortOrder { get; set; }
        public int? Qty { get; set; }
        public int? TotalQty { get; set; }
        public string Description { get; set; } = "";
        public string PartName { get; set; } = "";
        public string ConfigurationName { get; set; } = "";
        public string Material { get; set; } = "";

        public int ExportRecordId { get; set; }
        public virtual ExportRecord ExportRecord { get; set; }

        public virtual CutTemplate CutTemplate { get; set; }
        public virtual FormProgram FormProgram { get; set; }
    }
}
