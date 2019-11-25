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

        private void button2_Click(object sender, EventArgs e)
        {
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

                Print(bom.BomFeature.Name);
                Print("Fetching components...");
                items.AddRange(GetItems(bom));
                Print($"Found {items.Count} component(s)");
            }

            ExportToDXF(items);
        }

        private void ExportToDXF(PartDoc part)
        {
            var prefix = prefixTextBox.Text;
            var model = part as ModelDoc2;
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
        }

        private void ExportToDXF(AssemblyDoc assembly)
        {
            Print("Fetching components...");

            var items = GetItems(assembly, false);

            Print("Found " + items.Count);
            Print("");

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

                var fileName = GetFileName(item);
                var savepath = Path.Combine(savePath, fileName + ".dxf");

                SetLightweightToResolved(item.Component);

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
                var bomName = drawingInfo != null ? string.Format("{0} {1} BOM", drawingInfo.JobNo, drawingInfo.DrawingNo) : "BOM";
                var bomFile = Path.Combine(savePath, bomName + ".xlsx");
                CreateBOMExcelFile(bomFile, items.ToList());
            }
            catch (Exception ex)
            {
                Print(ex.Message, Color.Red);
            }
        }

        private void SetLightweightToResolved(Component2 component)
        {
            var isSuppressed = component.IsSuppressed();

            if (isSuppressed)
                return;
            var suppressionState = (swComponentSuppressionState_e)component.GetSuppression();

            switch (suppressionState)
            {
                case swComponentSuppressionState_e.swComponentFullyResolved:
                case swComponentSuppressionState_e.swComponentResolved:
                    return;

                case swComponentSuppressionState_e.swComponentFullyLightweight:
                case swComponentSuppressionState_e.swComponentLightweight:
                    var error = (swSuppressionError_e)component.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);

                    if (error == swSuppressionError_e.swSuppressionChangeOk)
                    {
                        var model = component.GetModelDoc2() as ModelDoc2;

                        if (model != null)
                        {
                            model.ForceRebuild3(false);
                        }
                    }
                    break;
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

        private string ChangePathExtension(string fullpath, string newExtension)
        {
            var dir = Path.GetDirectoryName(fullpath);
            var name = Path.GetFileNameWithoutExtension(fullpath);

            return Path.Combine(dir, name + newExtension);
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
                view.ShowSheetMetalBendNotes = true;

                var drawingModel = templateDrawing as ModelDoc2;
                drawingModel.ViewZoomtofit2();

                if (HasSupressedBends(view))
                {
                    Print("A bend is suppressed, please check flat pattern!", Color.Red);
                }

                if (HideModelSketches(view))
                {
                    // delete the current view that has sketches shown
                    drawingModel.SelectByName(0, view.Name);
                    drawingModel.DeleteSelection(false);

                    // recreate the flat pattern view
                    view = templateDrawing.CreateFlatPatternViewFromModelView3(partModel.GetPathName(), partConfiguration, 0, 0, 0, false, false);
                    view.ShowSheetMetalBendNotes = true;
                }

                if (ShouldFlipView(view))
                {
                    Print(partModel.GetTitle() + " - Flipped view", Color.Blue);
                    view.FlipView = true;
                }

                drawingModel.SaveAs(savePath);

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

        private bool HasSupressedBends(IView view)
        {
            var model = view.ReferencedDocument;
            var refConfig = view.ReferencedConfiguration;
            model.ShowConfiguration(refConfig);

            var flatPattern = model.GetFeatureByTypeName("FlatPattern");
            var bends = flatPattern.GetAllSubFeaturesByTypeName("UiBend");

            foreach (var bend in bends)
            {
                if (bend.IsSuppressed())
                    return true;
            }

            return false;

        }

        private bool HideModelSketches(IView view)
        {
            var model = view.ReferencedDocument;
            var activeConfig = ((Configuration)model.GetActiveConfiguration()).Name;

            var modelChanged = false;
            var refConfig = view.ReferencedConfiguration;
            model.ShowConfiguration(refConfig);

            var sketches = model.GetAllFeaturesByTypeName("ProfileFeature");

            foreach (var sketch in sketches)
            {
                var visible = (swVisibilityState_e)sketch.Visible;

                if (visible == swVisibilityState_e.swVisibilityStateShown)
                {
                    sketch.Select2(true, -1);
                    model.BlankSketch();
                    modelChanged = true;
                }
            }

            model.ShowConfiguration(activeConfig);

            return modelChanged;
        }

        private void CreateBOMExcelFile(string filepath, IList<Item> items)
        {
            var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "BomTemplate.xlsx");

            File.Copy(templatePath, filepath, true);

            var newFile = new FileInfo(filepath);

            using (var pkg = new ExcelPackage(newFile))
            {
                var workbook = pkg.Workbook;
                var partsSheet = workbook.Worksheets["Parts"];

                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    var row = i + 2;
                    var col = 1;

                    partsSheet.Cells[row, col++].Value = item.ItemNo;
                    partsSheet.Cells[row, col++].Value = item.FileName;
                    partsSheet.Cells[row, col++].Value = item.Quantity;
                    partsSheet.Cells[row, col++].Value = item.Description;
                    partsSheet.Cells[row, col++].Value = item.PartName;

                    if (item.Thickness > 0)
                        partsSheet.Cells[row, col].Value = item.Thickness;
                    col++;

                    partsSheet.Cells[row, col++].Value = item.Material;

                    if (item.KFactor > 0)
                        partsSheet.Cells[row, col].Value = item.KFactor;
                    col++;

                    if (item.BendRadius > 0)
                        partsSheet.Cells[row, col].Value = item.BendRadius;
                    col++;
                }

                for (int i = 1; i <= 9; i++)
                {
                    var column = partsSheet.Column(i);

                    if (column.Style.WrapText)
                        continue;

                    column.AutoFit();
                    column.Width += 1;
                }

                workbook.Calculate();
                pkg.Save();
            }
        }

        private string UserSelectFolder()
        {
            string path = null;

            Invoke(new MethodInvoker(() =>
            {
                var dlg = new FolderBrowserDialog();
                dlg.Description = "Where do you want to save the DXF files?";

                path = dlg.ShowDialog() == DialogResult.OK ? dlg.SelectedPath : null;
            }));

            return path;
        }

        private bool ShouldFlipView(SolidWorks.Interop.sldworks.View view)
        {
            return viewFlipDecider.ShouldFlip(view);
        }

        private DrawingDoc CreateDrawing()
        {
            return sldWorks.NewDocument(
                DrawingTemplatePath,
                (int)swDwgPaperSizes_e.swDwgPaperDsize,
                1,
                1) as DrawingDoc;
        }

        private List<Item> GetItems(BomTableAnnotation bom)
        {
            var items = new List<Item>();
            var table = bom as TableAnnotation;
            var itemNoColumnIndex = table.IndexOfColumnType(swTableColumnTypes_e.swBomTableColumnType_ItemNumber);

            if (itemNoColumnIndex == -1)
            {
                Print("Error: Item number column not found.");
                return null;
            }
            else
            {
                Print("Item numbers are in the " + Helper.GetNumWithSuffix(itemNoColumnIndex + 1) + " column.");
            }

            var qtyColumnIndex = table.IndexOfColumnType(swTableColumnTypes_e.swBomTableColumnType_Quantity);
            if (qtyColumnIndex == -1)
            {
                Print("Error: Quantity column not found.");
                return null;
            }

            var descriptionColumnIndex = table.IndexOfColumnTitle("Description");
            if (descriptionColumnIndex == -1)
            {
                Print("Error: Description column not found.");
                return null;
            }

            var partNoColumnIndex = table.IndexOfColumnType(swTableColumnTypes_e.swBomTableColumnType_PartNumber);
            if (partNoColumnIndex == -1)
            {
                Print("Error: Part number column not found.");
                return null;
            }

            var isBOMPartsOnly = bom.BomFeature.TableType == (int)swBomType_e.swBomType_PartsOnly;

            for (int rowIndex = 0; rowIndex < table.RowCount; rowIndex++)
            {
                //if (table.RowHidden[rowIndex] == true)
                //    continue;

                var bomComponents = isBOMPartsOnly ?
                    ((Array)bom.GetComponents2(rowIndex, bom.BomFeature.Configuration))?.Cast<Component2>() :
                    ((Array)bom.GetComponents(rowIndex))?.Cast<Component2>();

                if (bomComponents == null)
                    continue;

                var distinctComponents = bomComponents
                    .GroupBy(c => c.ReferencedConfiguration)
                    .Select(group => group.First())
                    .ToList();

                foreach (var component in bomComponents)
                {
                    if (component.IsSuppressed())
                    {
                        continue;
                    }

                    var qtyString = table.DisplayedText[rowIndex, qtyColumnIndex];
                    int qty = 0;

                    int.TryParse(qtyString, out qty);

                    items.Add(new Item
                    {
                        PartName = table.DisplayedText[rowIndex, partNoColumnIndex],
                        Quantity = qty,
                        Description = table.DisplayedText[rowIndex, descriptionColumnIndex],
                        ItemNo = table.DisplayedText[rowIndex, itemNoColumnIndex],
                        Component = component
                    });

                    break;
                }
            }

            return items;
        }

        private List<Item> GetItems(AssemblyDoc assembly, bool topLevel)
        {
            var list = new List<Item>();

            assembly.ResolveAllLightWeightComponents(false);

            var assemblyComponents = ((Array)assembly.GetComponents(topLevel))
                .Cast<Component2>()
                .Where(c => !c.IsHidden(true));

            var componentGroups = assemblyComponents
                .GroupBy(c => c.GetTitle() + c.ReferencedConfiguration);

            foreach (var group in componentGroups)
            {
                var component = group.First();

                var model = component.GetModelDoc2() as ModelDoc2;

                if (model == null)
                    continue;

                var n1 = Path.GetFileNameWithoutExtension(component.GetTitle());
                var name = component.ReferencedConfiguration.ToLower() == "default" ? n1 : string.Format("{0} [{1}]", n1, component.ReferencedConfiguration);

                list.Add(new Item
                {
                    PartName = name,
                    Quantity = group.Count(),
                    Component = component
                });
            }

            return list;
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