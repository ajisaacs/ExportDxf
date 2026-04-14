using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExportDXF.ApiClient
{
    public interface IFabWorksApiClient
    {
        Task<ApiExportDetail> CreateExportAsync(string drawingNumber, string equipmentNo, string drawingNo, string sourceFilePath, string outputFolder, string title = null);
        Task<ApiExportDetail> GetExportBySourceFileAsync(string filePath);
        Task<List<string>> GetDrawingNumbersAsync();
        Task<List<string>> GetEquipmentNumbersAsync();
        Task<List<string>> GetDrawingNumbersByEquipmentAsync(string equipmentNo = null);
        Task<string> GetNextItemNumberAsync(string drawingNumber);
        Task UpdatePdfHashAsync(int exportId, string pdfContentHash);
        Task<string> GetPreviousPdfHashAsync(string drawingNumber, int? excludeId = null);
        Task<ApiBomItem> FindExistingBomItemAsync(int exportId, string partName, string configurationName);
        Task<ApiBomItem> CreateBomItemAsync(int exportId, ApiBomItem bomItem);
        Task<ApiCutTemplate> GetPreviousCutTemplateAsync(string drawingNumber, string itemNo);
        Task<ApiFileUploadResponse> UploadDxfAsync(string localFilePath, string equipment, string drawingNo, string itemNo, string contentHash);
        Task<ApiFileUploadResponse> UploadPdfAsync(string localFilePath, string equipment, string drawingNo, string contentHash, int? exportRecordId = null);
    }
}
