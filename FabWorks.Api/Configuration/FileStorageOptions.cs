namespace FabWorks.Api.Configuration
{
    public class FileStorageOptions
    {
        public const string SectionName = "FileStorage";

        public string OutputFolder { get; set; } = @"C:\ExportDXF\Output";
    }
}
