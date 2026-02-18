namespace FabWorks.Api.DTOs
{
    public class CreateExportRequest
    {
        public string DrawingNumber { get; set; }
        public string SourceFilePath { get; set; }
        public string OutputFolder { get; set; }
    }
}
