using System.IO;

namespace ExportDXF.Services
{
    /// <summary>
    /// Extracts equipment/drawing numbers from document names matching the
    /// workplace format (e.g., "4321 A01.SLDDRW" → equipment "4321", drawing "A01").
    /// Uses the existing DrawingInfo.Parse() logic.
    /// </summary>
    public class EquipmentDrawingInfoExtractor : IDrawingInfoExtractor
    {
        public bool TryExtract(string documentName, out DrawingInfoResult info)
        {
            info = null;

            var name = Path.GetFileNameWithoutExtension(documentName);
            var parsed = DrawingInfo.Parse(name);

            if (parsed == null || string.IsNullOrEmpty(parsed.EquipmentNo))
                return false;

            var template = !string.IsNullOrEmpty(parsed.DrawingNo)
                ? $"{parsed.EquipmentNo} {parsed.DrawingNo} PT{{item_no:2}}"
                : $"{parsed.EquipmentNo} PT{{item_no:2}}";

            info = new DrawingInfoResult
            {
                EquipmentNumber = parsed.EquipmentNo,
                DrawingNumber = parsed.DrawingNo,
                DefaultTemplate = template
            };

            return true;
        }
    }
}
