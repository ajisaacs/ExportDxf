using System.Text.RegularExpressions;

namespace ExportDXF
{
    public class DrawingInfo
    {
        public string JobNo { get; set; }

        public string DrawingNo { get; set; }

        public string Source { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1}", JobNo, DrawingNo);
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
            const string pattern = @"(?<jobNo>[34]\d{3})\s?(?<dwgNo>[ABEP]\d+(-\d+|[A-Z])?)";

            var match = Regex.Match(input, pattern);

            if (match.Success == false)
                return null;

            var dwg = new DrawingInfo();

            dwg.JobNo = match.Groups["jobNo"].Value;
            dwg.DrawingNo = match.Groups["dwgNo"].Value;
            dwg.Source = input;

            return dwg;
        }
    }
}
