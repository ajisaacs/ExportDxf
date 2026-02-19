using ExportDXF.ApiClient;
using ExportDXF.Forms;
using ExportDXF.Services;
using System;
using System.Configuration;
using System.Net.Http;
using System.Windows.Forms;

namespace ExportDXF
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var container = new ServiceContainer();
            var mainForm = container.ResolveMainForm();
            Application.Run(mainForm);
        }
    }

    /// <summary>
    /// Simple dependency injection container.
    /// </summary>
    public class ServiceContainer
    {
        private readonly string _apiBaseUrl;

        public ServiceContainer()
        {
            _apiBaseUrl = ConfigurationManager.AppSettings["FabWorksApiUrl"] ?? "http://localhost:5206";
        }

        public MainForm ResolveMainForm()
        {
            var solidWorksService = new SolidWorksService();
            var bomExtractor = new BomExtractor();
            var partExporter = new PartExporter();
            var drawingExporter = new DrawingExporter();

            var httpClient = new HttpClient
            {
                BaseAddress = new Uri(_apiBaseUrl),
                Timeout = TimeSpan.FromSeconds(30)
            };
            var apiClient = new FabWorksApiClient(httpClient);

            var exportService = new DxfExportService(
                solidWorksService,
                bomExtractor,
                partExporter,
                drawingExporter,
                apiClient);

            return new MainForm(solidWorksService, exportService, apiClient);
        }
    }
}
