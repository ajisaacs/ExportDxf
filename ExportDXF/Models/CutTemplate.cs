using System;

namespace ExportDXF.Models
{
    public class CutTemplate
    {
        public int Id { get; set; }
        public string DxfFilePath { get; set; } = "";
        public string ContentHash { get; set; }
        public string CutTemplateName { get; set; } = "";

        // Sheet metal properties (moved from BomItem)
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

        // FK back to BomItem
        public int BomItemId { get; set; }
        public virtual BomItem BomItem { get; set; }
    }
}
