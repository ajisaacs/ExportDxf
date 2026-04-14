using System.Configuration;
using System.IO;

namespace ExportDXF.Services
{
    /// <summary>
    /// Fallback extractor that uses the document name as prefix
    /// and appends the configured default suffix.
    /// Always returns true — this is the catch-all.
    /// </summary>
    public class DefaultDrawingInfoExtractor : IDrawingInfoExtractor
    {
        public bool TryExtract(string documentName, out DrawingInfoResult info)
        {
            var name = Path.GetFileNameWithoutExtension(documentName);
            var suffix = ConfigurationManager.AppSettings["DefaultSuffix"] ?? "PT{item_no:2}";

            info = new DrawingInfoResult
            {
                EquipmentNumber = null,
                DrawingNumber = null,
                DefaultTemplate = $"{name} {suffix}"
            };

            return true;
        }
    }
}
