using ExportDXF.ItemExtractors;
using ExportDXF.ViewFlipDeciders;
using OfficeOpenXml;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportDXF.Forms
{
    public partial class MainForm : Form
    {
        private SldWorks sldWorks;
        private BackgroundWorker worker;
        private DrawingDoc templateDrawing;
        private DateTime timeStarted;
        private IViewFlipDecider viewFlipDecider;

        public MainForm()
        {
            InitializeComponent();

            worker = new BackgroundWorker();
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;

            var type = typeof(IViewFlipDecider);
            var types = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && p.IsClass)
                .ToList();

            comboBox1.DataSource = GetItems();
            comboBox1.DisplayMember = "Name";

            //viewFlipDecider = new AskViewFlipDecider();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            button1.Enabled = false;

            var task = new Task(ConnectToSolidWorks);

            task.ContinueWith((t) =>
            {
                Invoke(new MethodInvoker(() =>
                {
                    SetActiveDocName();
                    button1.Enabled = true;
                }));
            });

            task.Start();
        }

        private List<ViewFlipDeciderComboboxItem> GetItems()
        {
            var types = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => typeof(IViewFlipDecider).IsAssignableFrom(p) && p.IsClass)
                .ToList();

            var items = new List<ViewFlipDeciderComboboxItem>();

            foreach (var type in types)
            {
                var obj = (IViewFlipDecider)Activator.CreateInstance(type);

                items.Add(new ViewFlipDeciderComboboxItem
                {
                    Name = obj.Name,
                    ViewFlipDecider = obj
                });
            }

            return items;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (worker.IsBusy)
            {
                button1.Enabled = false;
                worker.CancelAsync();
                return;
            }
            else
            {
                worker.RunWorkerAsync();
            }
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            timeStarted = DateTime.Now;

            sldWorks.UserControl = false;
            sldWorks.ActiveModelDocChangeNotify -= SldWorks_ActiveModelDocChangeNotify;

            Invoke(new MethodInvoker(() =>
            {
                var item = comboBox1.SelectedItem as ViewFlipDeciderComboboxItem;
                viewFlipDecider = item.ViewFlipDecider;

                activeDocTitleBox.Enabled = false;
                prefixTextBox.Enabled = false;
                comboBox1.Enabled = false;

                button1.Image = Properties.Resources.stop_alt;

                if (richTextBox1.TextLength != 0)
                    richTextBox1.AppendText("\n\n");
            }));

            var model = sldWorks.ActiveDoc as ModelDoc2;

            Print("Started at " + DateTime.Now.ToShortTimeString());

            DetermineModelTypeAndExportToDXF(model);

            sldWorks.UserControl = true;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Invoke(new MethodInvoker(() =>
            {
                if (templateDrawing != null)
                {
                    sldWorks.CloseDoc(((ModelDoc2)templateDrawing).GetTitle());
                    templateDrawing = null;
                }

                activeDocTitleBox.Enabled = true;
                prefixTextBox.Enabled = true;
                comboBox1.Enabled = true;

                button1.Image = Properties.Resources.play;
                button1.Enabled = true;
            }));

            var duration = DateTime.Now - timeStarted;

            Print("Run time: " + duration.ToReadableFormat());
            Print("Done.", Color.Green);

            sldWorks.ActiveModelDocChangeNotify += SldWorks_ActiveModelDocChangeNotify;
        }

        private int SldWorks_ActiveModelDocChangeNotify()
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(() =>
                {
                    SetActiveDocName();
                }));
            }
            else
            {
                SetActiveDocName();
            }

            return 1;
        }

        private void Print(string s)
        {
            Invoke(new MethodInvoker(() =>
            {
                richTextBox1.AppendText(s + System.Environment.NewLine);
                richTextBox1.ScrollToCaret();
            }));
        }

        private void Print(string s, Color color)
        {
            Invoke(new MethodInvoker(() =>
            {
                richTextBox1.AppendText(s + System.Environment.NewLine, color);
                richTextBox1.ScrollToCaret();
            }));
        }

        private void ConnectToSolidWorks()
        {
            Print("Connecting to SolidWorks, this may take a minute...");
            sldWorks = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;

            if (sldWorks == null)
            {
                MessageBox.Show("Failed to connect to SolidWorks.");
                Application.Exit();
                return;
            }

            sldWorks.Visible = true;
            sldWorks.ActiveModelDocChangeNotify += SldWorks_ActiveModelDocChangeNotify;

            Print("Ready", Color.Green);
        }

        private void SetActiveDocName()
        {
            var model = sldWorks.ActiveDoc as ModelDoc2;
            activeDocTitleBox.Text = model == null ? "<No Document Open>" : model.GetTitle();
        }

        private void DetermineModelTypeAndExportToDXF(ModelDoc2 model)
        {
            if (model is PartDoc)
            {
                Print("Active document is a Part");
                ExportToDXF(model as PartDoc);
            }
            else if (model is DrawingDoc)
            {
                Print("Active document is a Drawing");
                ExportToDXF(model as DrawingDoc);
            }
            else if (model is AssemblyDoc)
            {
                Print("Active document is a Assembly");
                ExportToDXF(model as AssemblyDoc);
            }
        }

        private string GetPdfSavePath(DrawingDoc drawingDoc)
        {
            var model = drawingDoc as ModelDoc2;
            var pdfPath = model.GetPathName();
            var ext = Path.GetExtension(pdfPath);
            pdfPath = pdfPath.Remove(pdfPath.Length - ext.Length) + ".pdf";

            return pdfPath;
        }

        private void ExportDrawingToPDF(DrawingDoc drawingDoc, string savePath)
        {
            var model = drawingDoc as ModelDoc2;

            var exportData = sldWorks.GetExportFileData((int)swExportDataFileType_e.swExportPdfData) as ExportPdfData;
            exportData.ViewPdfAfterSaving = false;
            exportData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets, drawingDoc);

            int errors = 0;
            int warnings = 0;

            var modelExtension = model.Extension;
            modelExtension.SaveAs(savePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, exportData, ref errors, ref warnings);

            Print($"Saved drawing to PDF file \"{savePath}\"", Color.Green);
        }

        private void ExportToDXF(DrawingDoc drawing)
        {
            Print("Finding BOM tables...");
            var bomTables = drawing.GetBomTables();

            if (bomTables.Count == 0)
            {
                Print("Error: Bill of materials not found.", Color.Red);
                return;
            }

            Print($"Found {bomTables.Count} BOM table(s)\n");

            var items = new List<Item>();

            foreach (var bom in bomTables)
            {
                if (worker.CancellationPending)
                    return;

                var itemExtractor = new BomItemExtractor(bom);
                itemExtractor.SkipHiddenRows = true;

                Print($"Fetching components from {bom.BomFeature.Name}");

                var bomItems = itemExtractor.GetItems();
                items.AddRange(bomItems);
            }

            Print($"Found {items.Count} component(s)");

            ExportDrawingToPDF(drawing);
            ExportToDXF(items);
        }

        private void ExportToDXF(PartDoc part)
        {
            var prefix = prefixTextBox.Text;
            var model = part as ModelDoc2;
            var activeConfig = ((Configuration)model.GetActiveConfiguration()).Name;

            var dir = UserSelectFolder();

            if (dir == null)
            {
                Print("Canceled\n", Color.Red);
                return;
            }

            if (dir == null)
                return;

            var title = model.GetTitle().Replace(".SLDPRT", "");
            var config = model.ConfigurationManager.ActiveConfiguration.Name;
            var name = config.ToLower() == "default" ? title : string.Format("{0} [{1}]", title, config);
            var savePath = Path.Combine(dir, prefix + name + ".dxf");

            SavePartToDXF(part, savePath);

            model.ShowConfiguration(activeConfig);
        }

        private void ExportToDXF(AssemblyDoc assembly)
        {
            Print("Fetching components...");

            var itemExtractor = new AssemblyItemExtractor(assembly);
            itemExtractor.TopLevelOnly = false;

            var items = itemExtractor.GetItems();

            Print($"Found {items.Count} item(s).\n");
            ExportToDXF(items);
        }

        private void ExportToDXF(IEnumerable<Item> items)
        {
            var savePath = UserSelectFolder();
            var prefix = prefixTextBox.Text;

            if (savePath == null)
            {
                Print("Canceled\n", Color.Red);
                return;
            }

            templateDrawing = CreateDrawing();

            Print("");

            foreach (var item in items)
            {
                if (worker.CancellationPending)
                    return;

                if (item.Component == null)
                    continue;

                var fileName = GetFileName(item);
                var savepath = Path.Combine(savePath, fileName + ".dxf");

                item.Component.SetLightweightToResolved();

                var model = item.Component.GetModelDoc2() as ModelDoc2;
                var part = model as PartDoc;

                if (part == null)
                {
                    Print(item.ItemNo + " - skipped, not a part document");
                    continue;
                }

                var config = item.Component.ReferencedConfiguration;

                var sheetMetal = model.GetFeatureByTypeName("SheetMetal");
                var sheetMetalData = sheetMetal?.GetDefinition() as SheetMetalFeatureData;

                if (sheetMetalData != null)
                {
                    item.Thickness = sheetMetalData.Thickness.FromSldWorks();
                    item.KFactor = sheetMetalData.KFactor;
                    item.BendRadius = sheetMetalData.BendRadius.FromSldWorks();
                }

                if (item.Description == null)
                    item.Description = model.Extension.CustomPropertyManager[config].Get("Description");

                if (item.Description == null)
                    item.Description = model.Extension.CustomPropertyManager[""].Get("Description");

                item.Description = RemoveFontXml(item.Description);

                var db = string.Empty;

                item.Material = part.GetMaterialPropertyName2(config, out db);

                if (part == null)
                    continue;

                if (SavePartToDXF(part, config, savepath))
                {
                    item.FileName = Path.GetFileNameWithoutExtension(savepath);
                }
                else
                {
                    var desc = item.Description.ToLower();

                    if (desc.Contains("laser"))
                    {
                        Print($"Failed to export item #{item.ItemNo} but description says it is laser cut.", Color.Red);
                    }
                    else if (desc.Contains("plasma"))
                    {
                        Print($"Failed to export item #{item.ItemNo} but description says it is plasma cut.", Color.Red);
                    }
                }

                Print("");

                Application.DoEvents();
            }

            try
            {
                var drawingInfo = DrawingInfo.Parse(prefix);
                var bomName = drawingInfo != null ? $"{drawingInfo.JobNo} {drawingInfo.DrawingNo} BOM" : "BOM";
                var bomFile = Path.Combine(savePath, bomName + ".xlsx");

                var excelReport = new BomToExcel();
                excelReport.CreateBOMExcelFile(bomFile, items.ToList());
            }
            catch (Exception ex)
            {
                Print(ex.Message, Color.Red);
            }
        }

        private string RemoveFontXml(string s)
        {
            if (s == null)
                return null;

            var fontXmlRegex = new Regex("<FONT.*?\\>");
            var matches = fontXmlRegex.Matches(s)
                .Cast<Match>()
                .OrderByDescending(m => m.Index);

            foreach (var match in matches)
            {
                s = s.Remove(match.Index, match.Length);
            }

            return s;
        }

        private string GetFileName(Item item)
        {
            var prefix = prefixTextBox.Text.Replace("\"", "''");

            if (string.IsNullOrWhiteSpace(item.ItemNo))
            {
                return prefix + item.PartName;
            }
            else
            {
                return prefix + item.ItemNo.PadLeft(2, '0');
            }
        }

        private bool SavePartToDXF(PartDoc part, string savePath)
        {
            var partModel = part as ModelDoc2;
            var config = partModel.ConfigurationManager.ActiveConfiguration.Name;

            return SavePartToDXF(part, config, savePath);
        }

        private bool SavePartToDXF(PartDoc part, string partConfiguration, string savePath)
        {
            try
            {
                var partModel = part as ModelDoc2;

                if (partModel.IsSheetMetal() == false)
                {
                    Print(partModel.GetTitle() + " - skipped, not sheet metal");
                    return false;
                }

                if (templateDrawing == null)
                    templateDrawing = CreateDrawing();

                var sheet = templateDrawing.IGetCurrentSheet();
                var modelName = Path.GetFileNameWithoutExtension(partModel.GetPathName());
                sheet.SetName(modelName);

                Print(partModel.GetTitle() + " - Creating flat pattern.");
                SolidWorks.Interop.sldworks.View view;
                view = templateDrawing.CreateFlatPatternViewFromModelView3(partModel.GetPathName(), partConfiguration, 0, 0, 0, false, false);

                if (view == null)
                    return false;

                view.ShowSheetMetalBendNotes = true;

                var drawingModel = templateDrawing as ModelDoc2;
                drawingModel.ViewZoomtofit2();

                if (ViewHelper.HasSupressedBends(view))
                {
                    Print("A bend is suppressed, please check flat pattern!", Color.Red);
                }

                if (ViewHelper.HideModelSketches(view))
                {
                    // delete the current view that has sketches shown
                    drawingModel.SelectByName(0, view.Name);
                    drawingModel.DeleteSelection(false);

                    // recreate the flat pattern view
                    view = templateDrawing.CreateFlatPatternViewFromModelView3(partModel.GetPathName(), partConfiguration, 0, 0, 0, false, false);
                    view.ShowSheetMetalBendNotes = true;
                }

                if (viewFlipDecider.ShouldFlip(view))
                {
                    Print(partModel.GetTitle() + " - Flipped view", Color.Blue);
                    view.FlipView = true;
                }

                drawingModel.SaveAs(savePath);

                var etcher = new EtchBendLines.Etcher();
                etcher.AddEtchLines(savePath);

                Print(partModel.GetTitle() + " - Saved to \"" + savePath + "\"", Color.Green);

                drawingModel.SelectByName(0, view.Name);
                drawingModel.DeleteSelection(false);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private string UserSelectFolder()
        {
            string path = null;

            Invoke(new MethodInvoker(() =>
            {
                var dlg = new FolderBrowserDialog();
                dlg.Description = "Where do you want to save the DXF files?";

                if (dlg.ShowDialog() != DialogResult.OK)
                    throw new Exception("Export canceled by user.");

                path = dlg.SelectedPath;
            }));

            return path;
        }

        private DrawingDoc CreateDrawing()
        {
            return sldWorks.NewDocument(
                DrawingTemplatePath,
                (int)swDwgPaperSizes_e.swDwgPaperDsize,
                1,
                1) as DrawingDoc;
        }

        private static string DrawingTemplatePath
        {
            get { return Path.Combine(Application.StartupPath, "Templates", "Blank.drwdot"); }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var model = sldWorks.ActiveDoc as ModelDoc2;
            var isDrawing = model is DrawingDoc;

            if (!isDrawing)
                return;

            var drawingInfo = DrawingInfo.Parse(activeDocTitleBox.Text);

            if (drawingInfo == null)
                return;

            prefixTextBox.Text = string.Format("{0} {1} PT", drawingInfo.JobNo, drawingInfo.DrawingNo);
            prefixTextBox.SelectionStart = prefixTextBox.Text.Length;
        }
    }
}