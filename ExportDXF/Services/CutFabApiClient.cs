using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using static ExportDXF.Services.CutFabApiClient;

namespace ExportDXF.Services
{
    public interface ICutFabApiClient
    {
        string BaseUrl { get; }
        Task<int?> ResolveDrawingIdAsync(string drawingNumber);
        Task<int?> CreateDrawingAsync(int equipmentId, string drawingNumber);
        Task<CutFabApiClient.ApiResponse<int?>> CreateDrawingWithInfoAsync(int equipmentId, string drawingNumber);
        Task<bool> UploadDrawingPdfAsync(string drawingNumber, string pdfPath, string uploadedBy = null, string notes = null);
        Task<bool> UploadDxfZipAsync(
            int drawingId,
            string zipPath,
            double? thickness = null,
            double? kfactor = null,
            double? defaultBendRadius = null,
            string material = null);
        Task<int?> CreateBomItemAsync(object upsertBomItemDto);
        Task<bool> AutoLinkTemplatesAsync(int drawingId);
        Task<List<ApiEquipment>> GetEquipmentAsync();
        Task<List<ApiDrawingSummary>> GetDrawingsForEquipmentAsync(int equipmentId);
        Task<List<ApiBomItem>> GetBomItemsForDrawingAsync(int drawingId);
        Task<List<ApiCutTemplate>> GetCutTemplatesAsync();
    }

    public class CutFabApiClient : ICutFabApiClient, IDisposable
    {
        private readonly HttpClient _http;
        private readonly string _baseUrl;
        public string BaseUrl => _baseUrl;
        public class ApiResponse<T>
        {
            public bool Success { get; set; }
            public int StatusCode { get; set; }
            public T Data { get; set; }
            public string RawBody { get; set; }
            public string Error { get; set; }
        }

        public CutFabApiClient(string baseUrl)
        {
            _baseUrl = (baseUrl ?? string.Empty).TrimEnd('/');

            if (string.IsNullOrWhiteSpace(_baseUrl))
            {
                // Default to deployed API port from deployment script
                _baseUrl = "http://localhost:7027";
            }

            _http = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(100)
            };
        }

        public async Task<int?> ResolveDrawingIdAsync(string drawingNumber)
        {
            try
            {
                var url = $"{_baseUrl}/api/Drawings/resolve?drawingNumber={Uri.EscapeDataString(drawingNumber)}";
                var resp = await _http.GetAsync(url).ConfigureAwait(false);
                if (!resp.IsSuccessStatusCode) return null;
                var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                var dict = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(json);
                if (dict != null && dict.ContainsKey("ID"))
                {
                    var idObj = dict["ID"]; // serializer returns int or double depending
                    if (idObj is int i) return i;
                    if (idObj is double d) return (int)d;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public async Task<int?> CreateDrawingAsync(int equipmentId, string drawingNumber)
        {
            var payload = new { DrawingNumber = drawingNumber, Description = (string)null, Qty = 1, EquipmentID = equipmentId };
            var json = new JavaScriptSerializer().Serialize(payload);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var resp = await _http.PostAsync($"{_baseUrl}/api/Drawings", content).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode) return null;
            var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            try
            {
                var dict = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(body);
                if (dict != null && dict.ContainsKey("ID"))
                {
                    var idObj = dict["ID"]; if (idObj is int i) return i; if (idObj is double d) return (int)d;
                }
            }
            catch { }
            return null;
        }

        public async Task<ApiResponse<int?>> CreateDrawingWithInfoAsync(int equipmentId, string drawingNumber)
        {
            var result = new ApiResponse<int?> { Success = false, StatusCode = 0, Data = null, RawBody = null, Error = null };
            try
            {
                var payload = new { DrawingNumber = drawingNumber, Description = (string)null, Qty = 1, EquipmentID = equipmentId };
                var json = new JavaScriptSerializer().Serialize(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var resp = await _http.PostAsync($"{_baseUrl}/api/Drawings", content).ConfigureAwait(false);
                result.StatusCode = (int)resp.StatusCode;
                result.Success = resp.IsSuccessStatusCode;
                var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                result.RawBody = body;
                if (!resp.IsSuccessStatusCode)
                {
                    result.Error = "HTTP " + ((int)resp.StatusCode) + " " + resp.ReasonPhrase;
                    return result;
                }
                try
                {
                    var dict = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(body);
                    if (dict != null)
                    {
                        object v;
                        if (TryGetCI(dict, new[] { "ID", "id" }, out v))
                        {
                            if (v is int i) { result.Data = i; return result; }
                            if (v is double d) { result.Data = (int)d; return result; }
                            int parsed; if (int.TryParse(Convert.ToString(v), out parsed)) { result.Data = parsed; return result; }
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Error = ex.Message;
                }
                return result;
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
                return result;
            }
        }

        public async Task<bool> UploadDrawingPdfAsync(string drawingNumber, string pdfPath, string uploadedBy = null, string notes = null)
        {
            if (!File.Exists(pdfPath)) return false;
            using (var form = new MultipartFormDataContent())
            {
                form.Add(new StringContent(drawingNumber ?? string.Empty), "drawingNumber");
                if (!string.IsNullOrWhiteSpace(uploadedBy)) form.Add(new StringContent(uploadedBy), "uploadedBy");
                if (!string.IsNullOrWhiteSpace(notes)) form.Add(new StringContent(notes), "notes");

                var fileContent = new StreamContent(File.OpenRead(pdfPath));
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/pdf");
                var fileName = Path.GetFileName(pdfPath);
                form.Add(fileContent, "file", fileName);

                var resp = await _http.PostAsync($"{_baseUrl}/api/DrawingRevisions/upload", form).ConfigureAwait(false);
                return resp.IsSuccessStatusCode;
            }
        }

        public async Task<bool> UploadDxfZipAsync(int drawingId, string zipPath, double? thickness = null, double? kfactor = null, double? defaultBendRadius = null, string material = null)
        {
            if (!File.Exists(zipPath)) return false;
            using (var form = new MultipartFormDataContent())
            {
                var fileContent = new StreamContent(File.OpenRead(zipPath));
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/zip");
                var fileName = Path.GetFileName(zipPath);
                form.Add(fileContent, "file", fileName);

                // Add thickness and kfactor if provided
                if (thickness.HasValue)
                    form.Add(new StringContent(thickness.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)), "thickness");
                if (kfactor.HasValue)
                    form.Add(new StringContent(kfactor.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)), "kfactor");
                // Add default bend radius and material if provided
                if (defaultBendRadius.HasValue)
                    form.Add(new StringContent(defaultBendRadius.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)), "defaultBendRadius");
                if (!string.IsNullOrWhiteSpace(material))
                    form.Add(new StringContent(material), "material");

                var resp = await _http.PostAsync($"{_baseUrl}/api/Drawings/{drawingId}/upload-dxf-templates", form).ConfigureAwait(false);
                return resp.IsSuccessStatusCode;
            }
        }

        public async Task<int?> CreateBomItemAsync(object upsertBomItemDto)
        {
            var json = new JavaScriptSerializer().Serialize(upsertBomItemDto);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var resp = await _http.PostAsync($"{_baseUrl}/api/BomItems", content).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode) return null;
            var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            try
            {
                var dict = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(body);
                if (dict != null)
                {
                    object v;
                    if (TryGetCI(dict, new[] { "ID", "id" }, out v))
                    {
                        if (v is int i) return i;
                        if (v is double d) return (int)d;
                        int parsed; if (int.TryParse(Convert.ToString(v), out parsed)) return parsed;
                    }
                }
            }
            catch { }
            // Successful HTTP with empty/minimal body: treat as success
            return 0;
        }

        public async Task<bool> AutoLinkTemplatesAsync(int drawingId)
        {
            var url = $"{_baseUrl}/api/Drawings/{drawingId}/auto-link-templates";
            var resp = await _http.PostAsync(url, new ByteArrayContent(new byte[0])).ConfigureAwait(false);
            return resp.IsSuccessStatusCode;
        }

        public async Task<List<ApiBomItem>> GetBomItemsForDrawingAsync(int drawingId)
        {
            var url = $"{_baseUrl}/api/BomItems/drawing/{drawingId}";
            var resp = await _http.GetAsync(url).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode) return new List<ApiBomItem>();
            var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            var serializer = new JavaScriptSerializer();
            var raw = serializer.DeserializeObject(json);
            var result = new List<ApiBomItem>();

            if (raw is System.Collections.IEnumerable enumerable && !(raw is string))
            {
                foreach (var item in enumerable)
                {
                    var dict = item as Dictionary<string, object>;
                    if (dict == null) continue;
                    var bomItem = new ApiBomItem();
                    object v;
                    if (TryGetCI(dict, new[] { "ID", "id" }, out v)) bomItem.ID = ToInt(v);
                    if (TryGetCI(dict, new[] { "ItemNo", "itemNo" }, out v)) bomItem.ItemNo = v?.ToString();
                    if (TryGetCI(dict, new[] { "PartNo", "partNo" }, out v)) bomItem.PartNo = v?.ToString();
                    if (TryGetCI(dict, new[] { "SortOrder", "sortOrder" }, out v)) bomItem.SortOrder = ToInt(v);
                    if (TryGetCI(dict, new[] { "Qty", "qty" }, out v)) bomItem.Qty = v != null ? (int?)ToInt(v) : null;
                    if (TryGetCI(dict, new[] { "TotalQty", "totalQty" }, out v)) bomItem.TotalQty = v != null ? (int?)ToInt(v) : null;
                    if (TryGetCI(dict, new[] { "Description", "description" }, out v)) bomItem.Description = v?.ToString();
                    if (TryGetCI(dict, new[] { "PartName", "partName" }, out v)) bomItem.PartName = v?.ToString();
                    if (TryGetCI(dict, new[] { "ConfigurationName", "configurationName" }, out v)) bomItem.ConfigurationName = v?.ToString();
                    if (TryGetCI(dict, new[] { "Material", "material" }, out v)) bomItem.Material = v?.ToString();
                    if (TryGetCI(dict, new[] { "DrawingID", "drawingID", "drawingId" }, out v)) bomItem.DrawingID = ToInt(v);
                    if (TryGetCI(dict, new[] { "CutTemplateID", "cutTemplateID", "cutTemplateId" }, out v)) bomItem.CutTemplateID = v != null ? (int?)ToInt(v) : null;

                    // Try to get CutTemplate info if available
                    if (TryGetCI(dict, new[] { "CutTemplate", "cutTemplate" }, out v))
                    {
                        var templateDict = v as Dictionary<string, object>;
                        if (templateDict != null)
                        {
                            object tv;
                            if (TryGetCI(templateDict, new[] { "Name", "name" }, out tv)) bomItem.CutTemplateName = tv?.ToString();
                            if (TryGetCI(templateDict, new[] { "Thickness", "thickness" }, out tv)) bomItem.Thickness = ToDouble(tv);
                            if (TryGetCI(templateDict, new[] { "KFactor", "kFactor", "kfactor" }, out tv)) bomItem.KFactor = ToDouble(tv);
                            if (TryGetCI(templateDict, new[] { "DefaultBendRadius", "defaultBendRadius" }, out tv)) bomItem.DefaultBendRadius = ToDouble(tv);
                        }
                    }

                    result.Add(bomItem);
                }
            }

            return result;
        }

        public async Task<List<ApiCutTemplate>> GetCutTemplatesAsync()
        {
            var url = $"{_baseUrl}/api/CutTemplates";
            var resp = await _http.GetAsync(url).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode) return new List<ApiCutTemplate>();
            var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            var serializer = new JavaScriptSerializer();
            var raw = serializer.DeserializeObject(json);
            var result = new List<ApiCutTemplate>();

            if (raw is System.Collections.IEnumerable enumerable && !(raw is string))
            {
                foreach (var item in enumerable)
                {
                    var dict = item as Dictionary<string, object>;
                    if (dict == null) continue;
                    var template = new ApiCutTemplate();
                    object v;
                    if (TryGetCI(dict, new[] { "ID", "id" }, out v)) template.ID = ToInt(v);
                    if (TryGetCI(dict, new[] { "Name", "name" }, out v)) template.Name = v?.ToString();
                    if (TryGetCI(dict, new[] { "FilePath", "filePath" }, out v)) template.FilePath = v?.ToString();
                    if (TryGetCI(dict, new[] { "Description", "description" }, out v)) template.Description = v?.ToString();
                    if (TryGetCI(dict, new[] { "Material", "material" }, out v)) template.Material = v?.ToString();
                    if (TryGetCI(dict, new[] { "Thickness", "thickness" }, out v)) template.Thickness = ToDouble(v);
                    if (TryGetCI(dict, new[] { "KFactor", "kFactor", "kfactor" }, out v)) template.KFactor = ToDouble(v);
                    if (TryGetCI(dict, new[] { "DefaultBendRadius", "defaultBendRadius" }, out v)) template.DefaultBendRadius = ToDouble(v);
                    if (TryGetCI(dict, new[] { "Version", "version" }, out v)) template.Version = ToInt(v);
                    result.Add(template);
                }
            }

            return result;
        }

        public void Dispose()
        {
            _http?.Dispose();
        }

        // Lightweight DTOs for UI binding
        public class ApiEquipment
        {
            public int ID { get; set; }
            public string EquipmentNumber { get; set; }
            public string Description { get; set; }
            public override string ToString() => EquipmentNumber ?? base.ToString();
        }

        public class ApiDrawingSummary
        {
            public int ID { get; set; }
            public string DrawingNumber { get; set; }
            public string Description { get; set; }
            public override string ToString() => DrawingNumber ?? base.ToString();
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
            public int DrawingID { get; set; }
            public int? CutTemplateID { get; set; }
            public string CutTemplateName { get; set; }
            public double? Thickness { get; set; }
            public double? KFactor { get; set; }
            public double? DefaultBendRadius { get; set; }
        }

        public class ApiCutTemplate
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public string FilePath { get; set; }
            public string Description { get; set; }
            public string Material { get; set; }
            public double? Thickness { get; set; }
            public double? KFactor { get; set; }
            public double? DefaultBendRadius { get; set; }
            public int Version { get; set; }
        }

        public async Task<List<ApiEquipment>> GetEquipmentAsync()
        {
            var url = $"{_baseUrl}/api/Equipment";
            var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var resp = await _http.SendAsync(req).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();
            var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);

            var serializer = new JavaScriptSerializer();
            var raw = serializer.DeserializeObject(json);
            var result = new List<ApiEquipment>();

            if (raw is System.Collections.IEnumerable enumerable && !(raw is string))
            {
                foreach (var item in enumerable)
                {
                    var dict = item as Dictionary<string, object>;
                    if (dict == null) continue;
                    var eq = new ApiEquipment();
                    object v;
                    if (TryGetCI(dict, new[] { "ID", "id" }, out v)) eq.ID = ToInt(v);
                    if (TryGetCI(dict, new[] { "EquipmentNumber", "equipmentNumber", "equipmentNo" }, out v)) eq.EquipmentNumber = v?.ToString();
                    if (TryGetCI(dict, new[] { "Description", "description" }, out v)) eq.Description = v?.ToString();
                    result.Add(eq);
                }
            }

            return result;
        }

        public async Task<List<ApiDrawingSummary>> GetDrawingsForEquipmentAsync(int equipmentId)
        {
            var url = $"{_baseUrl}/api/Equipment/{equipmentId}";
            var resp = await _http.GetAsync(url).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode) return new List<ApiDrawingSummary>();
            var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            var serializer = new JavaScriptSerializer();
            var root = serializer.DeserializeObject(json) as Dictionary<string, object>;
            var results = new List<ApiDrawingSummary>();
            if (root != null)
            {
                object dval;
                if (!TryGetCI(root, new[] { "Drawings", "drawings" }, out dval)) return results;
                if (dval is System.Collections.IEnumerable arr && !(dval is string))
                {
                    foreach (var item in arr)
                    {
                        var d = item as Dictionary<string, object>;
                        if (d == null) continue;
                        var summary = new ApiDrawingSummary();
                        object v;
                        if (TryGetCI(d, new[] { "ID", "id" }, out v)) summary.ID = ToInt(v);
                        if (TryGetCI(d, new[] { "DrawingNumber", "drawingNumber" }, out v)) summary.DrawingNumber = v?.ToString();
                        if (TryGetCI(d, new[] { "Description", "description" }, out v)) summary.Description = v?.ToString();
                        results.Add(summary);
                    }
                }
            }
            return results;
        }

        private static bool TryGetCI(Dictionary<string, object> dict, IEnumerable<string> keys, out object value)
        {
            foreach (var k in keys)
            {
                if (dict.ContainsKey(k)) { value = dict[k]; return true; }
                foreach (var dk in dict.Keys)
                {
                    if (string.Equals(dk, k, StringComparison.OrdinalIgnoreCase)) { value = dict[dk]; return true; }
                }
            }
            value = null; return false;
        }

        private static int ToInt(object v)
        {
            if (v == null) return 0;
            if (v is int i) return i;
            if (v is long l) return (int)l;
            if (v is double d) return (int)d;
            int parsed; if (int.TryParse(v.ToString(), out parsed)) return parsed; return 0;
        }

        private static double? ToDouble(object v)
        {
            if (v == null) return null;
            if (v is double d) return d;
            if (v is int i) return (double)i;
            if (v is long l) return (double)l;
            if (v is float f) return (double)f;
            double parsed; if (double.TryParse(v.ToString(), out parsed)) return parsed;
            return null;
        }
    }
}
