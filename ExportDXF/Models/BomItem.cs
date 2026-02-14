namespace ExportDXF.Models
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

        // EF Core relationship to ExportRecord
        public int ExportRecordId { get; set; }
        public virtual ExportRecord ExportRecord { get; set; }

        // Optional 1:1 relationship to CutTemplate (only for sheet metal parts)
        public virtual CutTemplate CutTemplate { get; set; }
    }

    public struct Size
    {
        public double Width { get; set; }
        public double Height { get; set; }
    }
}
