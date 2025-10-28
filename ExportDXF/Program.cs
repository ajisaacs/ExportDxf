using ExportDXF.Forms;
using ExportDXF.Services;
using System;
using System.Configuration;
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

            // Simple DI setup - could use a DI container like Microsoft.Extensions.DependencyInjection
            var container = new ServiceContainer();

            var mainForm = container.Resolve<MainForm>();

            Application.Run(mainForm);
        }
    }

    /// <summary>
    /// Simple dependency injection container.
    /// For production, consider using Microsoft.Extensions.DependencyInjection or similar.
    /// </summary>
    public class ServiceContainer
    {
        public MainForm Resolve<T>() where T : MainForm
        {
            // Create the dependency tree
            var solidWorksService = new SolidWorksService();

            var bomExtractor = new BomExtractor();
            var partExporter = new PartExporter();
            var drawingExporter = new DrawingExporter();
            var baseUrl = ConfigurationManager.AppSettings["CutFab.ApiBaseUrl"] ?? "http://localhost:7027";
            var apiClient = new CutFabApiClient(baseUrl);

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
