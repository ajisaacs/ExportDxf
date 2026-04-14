using System;
using System.Collections.Generic;

namespace ExportDXF.ApiClient
{
    public class ApiExportDetail
    {
        public int Id { get; set; }
        public string DrawingNumber { get; set; }
        public string Title { get; set; }
        public string EquipmentNo { get; set; }
        public string DrawingNo { get; set; }
        public string SourceFilePath { get; set; }
        public string OutputFolder { get; set; }
        public DateTime ExportedAt { get; set; }
        public string ExportedBy { get; set; }
        public string PdfContentHash { get; set; }
        public List<ApiBomItem> BomItems { get; set; } = new();
    }

    public class ApiBomItem
    {
        public int ID { get; set; }
        public string ItemNo { get; set; }
        public string PartNo { get; set; }
        public int SortOrder { get; set; }
        public int? Qty { get; set; }
        public int? TotalQty { get; set; }
        public string Description { get; set; }
        public string PartName { get; set; }
        public string ConfigurationName { get; set; }
        public string Material { get; set; }
        public ApiCutTemplate CutTemplate { get; set; }
        public ApiFormProgram FormProgram { get; set; }
    }

    public class ApiCutTemplate
    {
        public int Id { get; set; }
        public string DxfFilePath { get; set; }
        public string ContentHash { get; set; }
        public int Revision { get; set; }
        public double? Thickness { get; set; }
        public double? KFactor { get; set; }
        public double? DefaultBendRadius { get; set; }
    }

    public class ApiFormProgram
    {
        public int Id { get; set; }
        public string ProgramFilePath { get; set; }
        public string ContentHash { get; set; }
        public string ProgramName { get; set; }
        public double? Thickness { get; set; }
        public string MaterialType { get; set; }
        public double? KFactor { get; set; }
        public int BendCount { get; set; }
        public string UpperToolNames { get; set; }
        public string LowerToolNames { get; set; }
        public string SetupNotes { get; set; }
    }

    public class ApiCreateExportRequest
    {
        public string DrawingNumber { get; set; }
        public string Title { get; set; }
        public string EquipmentNo { get; set; }
        public string DrawingNo { get; set; }
        public string SourceFilePath { get; set; }
        public string OutputFolder { get; set; }
    }

    public class ApiUpdatePdfHashRequest
    {
        public string PdfContentHash { get; set; }
    }

    public class ApiFileUploadResponse
    {
        public string StoredFilePath { get; set; }
        public string ContentHash { get; set; }
        public string FileName { get; set; }
        public bool WasUnchanged { get; set; }
        public bool IsNewFile { get; set; }
    }
}
