using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;

namespace ExportDXF.ApiClient
{
    public class FabWorksApiClient : IFabWorksApiClient
    {
        private readonly HttpClient _http;

        public FabWorksApiClient(HttpClient httpClient)
        {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        }

        public async Task<ApiExportDetail> CreateExportAsync(string drawingNumber, string equipmentNo, string drawingNo, string sourceFilePath, string outputFolder, string title = null)
        {
            var request = new ApiCreateExportRequest
            {
                DrawingNumber = drawingNumber,
                Title = title,
                EquipmentNo = equipmentNo,
                DrawingNo = drawingNo,
                SourceFilePath = sourceFilePath,
                OutputFolder = outputFolder
            };

            var response = await _http.PostAsJsonAsync("api/exports", request);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<ApiExportDetail>();
        }

        public async Task<ApiExportDetail> GetExportBySourceFileAsync(string filePath)
        {
            var response = await _http.GetAsync($"api/exports/by-source?path={Uri.EscapeDataString(filePath)}");
            if (response.StatusCode == HttpStatusCode.NotFound)
                return null;
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<ApiExportDetail>();
        }

        public async Task<List<string>> GetDrawingNumbersAsync()
        {
            var response = await _http.GetAsync("api/exports/drawing-numbers");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<List<string>>();
        }

        public async Task<List<string>> GetEquipmentNumbersAsync()
        {
            var response = await _http.GetAsync("api/exports/equipment-numbers");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<List<string>>();
        }

        public async Task<List<string>> GetDrawingNumbersByEquipmentAsync(string equipmentNo = null)
        {
            var url = "api/exports/drawing-numbers-by-equipment";
            if (!string.IsNullOrEmpty(equipmentNo))
                url += $"?equipmentNo={Uri.EscapeDataString(equipmentNo)}";

            var response = await _http.GetAsync(url);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<List<string>>();
        }

        public async Task<string> GetNextItemNumberAsync(string drawingNumber)
        {
            var response = await _http.GetAsync($"api/exports/next-item-number?drawingNumber={Uri.EscapeDataString(drawingNumber)}");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        public async Task UpdatePdfHashAsync(int exportId, string pdfContentHash)
        {
            var request = new ApiUpdatePdfHashRequest { PdfContentHash = pdfContentHash };
            var response = await _http.PatchAsJsonAsync($"api/exports/{exportId}/pdf-hash", request);
            response.EnsureSuccessStatusCode();
        }

        public async Task<string> GetPreviousPdfHashAsync(string drawingNumber, int? excludeId = null)
        {
            var url = $"api/exports/previous-pdf-hash?drawingNumber={Uri.EscapeDataString(drawingNumber)}";
            if (excludeId.HasValue)
                url += $"&excludeId={excludeId.Value}";

            var response = await _http.GetAsync(url);
            if (response.StatusCode == HttpStatusCode.NotFound)
                return null;
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        public async Task<ApiBomItem> FindExistingBomItemAsync(int exportId, string partName, string configurationName)
        {
            var url = $"api/exports/{exportId}/bom-items/find?partName={Uri.EscapeDataString(partName ?? "")}&configurationName={Uri.EscapeDataString(configurationName ?? "")}";
            var response = await _http.GetAsync(url);
            if (response.StatusCode == HttpStatusCode.NotFound)
                return null;
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<ApiBomItem>();
        }

        public async Task<ApiBomItem> CreateBomItemAsync(int exportId, ApiBomItem bomItem)
        {
            var response = await _http.PostAsJsonAsync($"api/exports/{exportId}/bom-items", bomItem);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<ApiBomItem>();
        }

        public async Task<ApiCutTemplate> GetPreviousCutTemplateAsync(string drawingNumber, string itemNo)
        {
            var response = await _http.GetAsync($"api/exports/previous-cut-template?drawingNumber={Uri.EscapeDataString(drawingNumber)}&itemNo={Uri.EscapeDataString(itemNo)}");
            if (response.StatusCode == HttpStatusCode.NotFound)
                return null;
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<ApiCutTemplate>();
        }

        public async Task<ApiFileUploadResponse> UploadDxfAsync(string localFilePath, string equipment, string drawingNo, string itemNo, string contentHash)
        {
            using var content = new MultipartFormDataContent();
            using var fileStream = new FileStream(localFilePath, FileMode.Open, FileAccess.Read);
            var fileContent = new StreamContent(fileStream);
            content.Add(fileContent, "file", Path.GetFileName(localFilePath));
            content.Add(new StringContent(equipment ?? ""), "equipment");
            content.Add(new StringContent(drawingNo ?? ""), "drawingNo");
            content.Add(new StringContent(itemNo ?? ""), "itemNo");
            content.Add(new StringContent(contentHash ?? ""), "contentHash");

            var response = await _http.PostAsync("api/files/dxf", content);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<ApiFileUploadResponse>();
        }

        public async Task<ApiFileUploadResponse> UploadPdfAsync(string localFilePath, string equipment, string drawingNo, string contentHash, int? exportRecordId = null)
        {
            using var content = new MultipartFormDataContent();
            using var fileStream = new FileStream(localFilePath, FileMode.Open, FileAccess.Read);
            var fileContent = new StreamContent(fileStream);
            content.Add(fileContent, "file", Path.GetFileName(localFilePath));
            content.Add(new StringContent(equipment ?? ""), "equipment");
            content.Add(new StringContent(drawingNo ?? ""), "drawingNo");
            content.Add(new StringContent(contentHash ?? ""), "contentHash");
            if (exportRecordId.HasValue)
                content.Add(new StringContent(exportRecordId.Value.ToString()), "exportRecordId");

            var response = await _http.PostAsync("api/files/pdf", content);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<ApiFileUploadResponse>();
        }
    }
}
