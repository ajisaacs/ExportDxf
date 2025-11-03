using System;

namespace ExportDXF.Models
{
    public enum LogLevel { Info, Warning, Error }

    public enum LogAction { Start, FindBom, CreateFlat, FlipView, SavePdf, UploadPdf, UploadDxf, CreateBomItem }

    public sealed class LogEvent
    {
        public DateTime Time { get; set; } = DateTime.Now;
        public LogLevel Level { get; set; }
        public string Equipment { get; set; } = "";
        public string Drawing { get; set; } = "";
        public string Part { get; set; } = "";
        public LogAction Action { get; set; }
        public string Target { get; set; } = "";
        public string Result { get; set; } = "OK";
        public int DurationMs { get; set; }
        public string Message { get; set; } = "";
    }
}