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

            // Show drawing selection dialog first
            var drawingSelectionForm = container.ResolveDrawingSelection();
            var result = drawingSelectionForm.ShowDialog();

            if (result == DialogResult.OK && drawingSelectionForm.SelectedDrawingId.HasValue)
            {
                // User selected a drawing, proceed to main form
                var mainForm = container.Resolve<MainForm>(
                    drawingSelectionForm.SelectedDrawingId.Value,
                    drawingSelectionForm.SelectedDrawingNumber);

                Application.Run(mainForm);
            }
            // If user cancelled, just exit
        }
    }

    /// <summary>
    /// Simple dependency injection container.
    /// For production, consider using Microsoft.Extensions.DependencyInjection or similar.
    /// </summary>
    public class ServiceContainer
    {
        private readonly string _baseUrl;
        private readonly CutFabApiClient _apiClient;

        public ServiceContainer()
        {
            _baseUrl = ConfigurationManager.AppSettings["CutFab.ApiBaseUrl"] ?? "http://localhost:7027";
            _apiClient = new CutFabApiClient(_baseUrl);
        }

        public DrawingSelectionForm ResolveDrawingSelection()
        {
            return new DrawingSelectionForm(_apiClient);
        }

        public MainForm Resolve<T>(int selectedDrawingId, string selectedDrawingNumber) where T : MainForm
        {
            // Create the dependency tree
            var solidWorksService = new SolidWorksService();

            var bomExtractor = new BomExtractor();
            var partExporter = new PartExporter();
            var drawingExporter = new DrawingExporter();

            var exportService = new DxfExportService(
                solidWorksService,
                bomExtractor,
                partExporter,
                drawingExporter,
                _apiClient);

            return new MainForm(solidWorksService, exportService, _apiClient, selectedDrawingId, selectedDrawingNumber);
        }
    }
}
