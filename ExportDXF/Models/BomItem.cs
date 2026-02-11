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
        public string CutTemplateName { get; set; } = "";
        public string DxfFilePath { get; set; } = "";

        // Sheet metal properties
        private double? _thickness;
        public double? Thickness
        {
            get => _thickness;
            set => _thickness = value.HasValue ? Math.Round(value.Value, 8) : null;
        }

        public double? KFactor { get; set; }

        private double? _defaultBendRadius;
        public double? DefaultBendRadius
        {
            get => _defaultBendRadius;
            set => _defaultBendRadius = value.HasValue ? Math.Round(value.Value, 8) : null;
        }

        // EF Core relationship to ExportRecord
        public int ExportRecordId { get; set; }
        public virtual ExportRecord ExportRecord { get; set; }
    }

    public struct Size
    {
        public double Width { get; set; }
        public double Height { get; set; }
    }
}
