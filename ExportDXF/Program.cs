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
        private readonly string _outputFolder;

        public ServiceContainer()
        {
            _outputFolder = ConfigurationManager.AppSettings["ExportOutputFolder"] ?? @"C:\ExportDXF\Output";
        }

        public MainForm ResolveMainForm()
        {
            var solidWorksService = new SolidWorksService();
            var bomExtractor = new BomExtractor();
            var partExporter = new PartExporter();
            var drawingExporter = new DrawingExporter();
            var fileExportService = new FileExportService(_outputFolder);

            var exportService = new DxfExportService(
                solidWorksService,
                bomExtractor,
                partExporter,
                drawingExporter,
                fileExportService);

            return new MainForm(solidWorksService, exportService, fileExportService);
        }
    }
}
