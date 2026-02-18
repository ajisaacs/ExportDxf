using System;
using System.Collections.Generic;

namespace FabWorks.Core.Models
{
    public class ExportRecord
    {
        public int Id { get; set; }
        public string DrawingNumber { get; set; }
        public string SourceFilePath { get; set; }
        public string OutputFolder { get; set; }
        public DateTime ExportedAt { get; set; }
        public string ExportedBy { get; set; }
        public string PdfContentHash { get; set; }

        public virtual ICollection<BomItem> BomItems { get; set; } = new List<BomItem>();
    }
}
