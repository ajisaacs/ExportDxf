using System.Text.RegularExpressions;

namespace ExportDXF
{
    public class DrawingInfo
    {
        private static Regex drawingFormatRegex = new Regex(@"(?<equipmentNo>[345]\d{3}(-\d+\w{1,2})?)\s?(?<dwgNo>[ABEP]\d+(-?(\d+[A-Z]?))?)", RegexOptions.IgnoreCase);
        private static Regex equipmentOnlyRegex = new Regex(@"^(?<equipmentNo>[345]\d{3}(-\d+\w{1,2})?)\b", RegexOptions.IgnoreCase);

        public string EquipmentNo { get; set; }

        public string DrawingNo { get; set; }

        public string Source { get; set; }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(DrawingNo))
                return EquipmentNo ?? string.Empty;
            return $"{EquipmentNo} {DrawingNo}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            return obj.ToString() == ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public static DrawingInfo Parse(string input)
        {
            var match = drawingFormatRegex.Match(input);

            if (match.Success == false)
            {
                // Try matching just the equipment number (e.g. "5028 Prox switch bracket")
                var eqMatch = equipmentOnlyRegex.Match(input);
                if (eqMatch.Success)
                {
                    return new DrawingInfo
                    {
                        EquipmentNo = eqMatch.Groups["equipmentNo"].Value,
                        DrawingNo = null,
                        Source = input
                    };
                }

                return null;
            }

            var dwg = new DrawingInfo();

            dwg.EquipmentNo = match.Groups["equipmentNo"].Value;
            dwg.DrawingNo = match.Groups["dwgNo"].Value;
            dwg.Source = input;

            return dwg;
        }
    }
}
