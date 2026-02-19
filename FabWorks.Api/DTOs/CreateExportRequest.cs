namespace FabWorks.Api.DTOs
{
    public class CreateExportRequest
    {
        public string DrawingNumber { get; set; }
        public string Title { get; set; }
        public string EquipmentNo { get; set; }
        public string DrawingNo { get; set; }
        public string SourceFilePath { get; set; }
        public string OutputFolder { get; set; }
    }

    public class UpdatePdfHashRequest
    {
        public string PdfContentHash { get; set; }
    }
}
