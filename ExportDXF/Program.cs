using ExportDXF.Forms;
using ExportDXF.Services;
using System;
using System.Windows.Forms;

namespace ExportDXF
{
    static class Program
    {
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

    public class ServiceContainer
    {
        public MainForm ResolveMainForm()
        {
            var solidWorksService = new SolidWorksService();
            var bomExtractor = new BomExtractor();
            var partExporter = new PartExporter();
            var drawingExporter = new DrawingExporter();
            var excelExportService = new ExcelExportService();
            var logFileService = new LogFileService();

            var exportService = new DxfExportService(
                solidWorksService,
                bomExtractor,
                partExporter,
                drawingExporter,
                excelExportService,
                logFileService);

            var extractors = new IDrawingInfoExtractor[]
            {
                new EquipmentDrawingInfoExtractor(),
                new DefaultDrawingInfoExtractor()
            };

            return new MainForm(solidWorksService, exportService, extractors);
        }
    }
}
