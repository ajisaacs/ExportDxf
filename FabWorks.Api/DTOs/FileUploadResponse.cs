namespace FabWorks.Api.DTOs
{
    public class FileUploadResponse
    {
        public string StoredFilePath { get; set; }  // kept for client compat, contains logical filename
        public string ContentHash { get; set; }
        public string FileName { get; set; }
        public bool WasUnchanged { get; set; }
        public bool IsNewFile { get; set; }
    }
}
