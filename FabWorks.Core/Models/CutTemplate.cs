using System;

namespace FabWorks.Core.Models
{
    public class CutTemplate
    {
        public int Id { get; set; }
        public string DxfFilePath { get; set; } = "";
        public string ContentHash { get; set; }
        public string CutTemplateName { get; set; } = "";

        private double? _thickness;
        public double? Thickness
        {
            get => _thickness;
            set => _thickness = value.HasValue ? Math.Round(value.Value, 8) : null;
        }

        public int Revision { get; set; } = 1;

        public double? KFactor { get; set; }

        private double? _defaultBendRadius;
        public double? DefaultBendRadius
        {
            get => _defaultBendRadius;
            set => _defaultBendRadius = value.HasValue ? Math.Round(value.Value, 8) : null;
        }

        public int BomItemId { get; set; }
        public virtual BomItem BomItem { get; set; }
    }
}
