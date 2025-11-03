using System;

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
        public int DrawingID { get; set; }
        public int? CutTemplateID { get; set; }
        public string CutTemplateName { get; set; } = "";

        // Sheet metal properties from CutTemplate
        public double? Thickness { get; set; }
        public double? KFactor { get; set; }
        public double? DefaultBendRadius { get; set; }
    }

    public struct Size
    {
        public double Width { get; set; }
        public double Height { get; set; }
    }
}
