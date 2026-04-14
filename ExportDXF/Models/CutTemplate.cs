using System;

namespace ExportDXF.Models
{
    public class CutTemplate
    {
        public string DxfFilePath { get; set; } = "";
        public string ContentHash { get; set; }
        public string CutTemplateName { get; set; } = "";
        public int Revision { get; set; } = 1;

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
    }
}
