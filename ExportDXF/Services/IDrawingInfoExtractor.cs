namespace ExportDXF.Services
{
    public interface IDrawingInfoExtractor
    {
        bool TryExtract(string documentName, out DrawingInfoResult info);
    }

    public class DrawingInfoResult
    {
        public string EquipmentNumber { get; set; }
        public string DrawingNumber { get; set; }
        public string DefaultTemplate { get; set; }
    }
}
