using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

        public MainForm()
        {
            InitializeComponent();

            worker = new BackgroundWorker();
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (worker.IsBusy)
            {
                worker.CancelAsync();
                return;
            }

            worker.RunWorkerAsync();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            timeStarted = DateTime.Now;

            sldWorks.ActiveModelDocChangeNotify -= SldWorks_ActiveModelDocChangeNotify;

            Invoke(new MethodInvoker(() =>
            {
                button1.Image = Properties.Resources.stop_alt;

                if (richTextBox1.TextLength != 0)
                    richTextBox1.AppendText("\n\n");
            }));
            
            var model = sldWorks.ActiveDoc as ModelDoc2;

            Print("Started at " + DateTime.Now.ToShortTimeString());

            DetermineModelTypeAndExportToDXF(model);

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

                button1.Image = Properties.Resources.play;
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

            textBox1.Text = model != null ? model.GetTitle() : "<No Document Open>";
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
            var bomTables = GetBomTables(drawing);

            if (bomTables.Count == 0)
            {
                Print("Error: Bill of materials not found.", Color.Red);
                return;
            }

            Print("Found " + bomTables.Count);
            Print("");

            foreach (var bom in bomTables)
            {
                if (worker.CancellationPending)
                    return;

                Print(bom.BomFeature.Name);

                Print("Fetching components...");

                var items = GetItems(bom);

                Print("Found " + items.Count);
                Print("");

                ExportToDXF(items);
            }
        }

        private void ExportToDXF(PartDoc part)
        {
            var prefix = textBox2.Text;
            var model = part as ModelDoc2;
            var dir = UserSelectFolder();

            if (dir == null)
            {
                Print("Cancelled\n", Color.Red);
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
            var prefix = textBox2.Text;

            if (savePath == null)
            {
                Print("Cancelled\n", Color.Red);
                return;
            }

            templateDrawing = CreateDrawing();

            Print("");

            foreach (var item in items)
            {
                if (worker.CancellationPending)
                    break;

                var fileName = prefix + item.Name + ".dxf";
                var savepath = Path.Combine(savePath, fileName);
                var part = item.Component.GetModelDoc2() as PartDoc;

                if (part == null)
                    continue;

                SavePartToDXF(part, item.Component.ReferencedConfiguration, savepath);
                Application.DoEvents();
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
                view.ShowSheetMetalBendNotes = true;

                var drawingModel = templateDrawing as ModelDoc2;
                drawingModel.ViewZoomtofit2();

                if (ShouldFlipView(view))
                {
                    Print(partModel.GetTitle() + " - Flipped view", Color.Blue);
                    view.FlipView = true;
                }

                drawingModel.SaveAs(savePath);

                Print(partModel.GetTitle() + " - Saved to \"" + savePath + "\"", Color.Green);
                Print("");

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

                path = dlg.ShowDialog() == DialogResult.OK ? dlg.SelectedPath : null;
            }));

            return path;
        }

        private bool ShouldFlipView(SolidWorks.Interop.sldworks.View view)
        {
            var notes = (view.GetNotes() as Array)?.Cast<Note>();

            var upCount = 0;
            var dnCount = 0;

            Note leftMost = null;
            double leftMostValue = double.MaxValue;

            foreach (var note in notes)
            {
                var pt = (note.GetTextPoint() as double[]);

                if (pt[0] < leftMostValue)
                {
                    leftMostValue = pt[0];
                    leftMost = note;
                }

                var txt = note.GetText();

                if (txt.ToUpper().Contains("UP"))
                    upCount++;
                else
                    dnCount++;
            }

            Print(string.Format("Found {0} bends,  {1} UP,  {2} DOWN", notes.Count(), upCount, dnCount), Color.Blue);

            if (dnCount == upCount && leftMost != null)
                return !leftMost.GetText().Contains("UP");

            return dnCount > upCount;
        }

        private DrawingDoc CreateDrawing()
        {
            return sldWorks.NewDocument(
                DrawingTemplatePath,
                (int)swDwgPaperSizes_e.swDwgPaperDsize,
                1,
                1) as DrawingDoc;
        }

        private List<BomTableAnnotation> GetBomTables(DrawingDoc drawing)
        {
            var model = drawing as ModelDoc2;

            return model.GetAllFeaturesByTypeName("BomFeat")
                .Select(f => f.GetSpecificFeature2() as BomFeature)
                .Select(f => (f.GetTableAnnotations() as Array)?.Cast<BomTableAnnotation>().FirstOrDefault())
                .ToList();
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

            var isBOMPartsOnly = bom.BomFeature.TableType == (int)swBomType_e.swBomType_PartsOnly;

            for (int rowIndex = 0; rowIndex < table.RowCount; rowIndex++)
            {
                if (table.RowHidden[rowIndex] == true)
                    continue;

                var bomComponents = isBOMPartsOnly ?
                    ((Array)bom.GetComponents2(rowIndex, bom.BomFeature.Configuration))?.Cast<Component2>() :
                    ((Array)bom.GetComponents(rowIndex))?.Cast<Component2>();

                if (bomComponents == null)
                    continue;

                var distinctComponents = bomComponents
                    .GroupBy(c => c.ReferencedConfiguration)
                    .Select(group => group.First());

                var itemNumber = table.Text[rowIndex, itemNoColumnIndex].PadLeft(2, '0');
                var rev = 'A';

                if (distinctComponents.Count() > 1)
                {
                    foreach (var comp in distinctComponents)
                    {
                        items.Add(new Item
                        {
                            Name = itemNumber + rev++,
                            Component = comp
                        });
                    }
                }
                else
                {
                    items.Add(new Item
                    {
                        Name = itemNumber,
                        Component = distinctComponents.First()
                    });
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

                var name = component.ReferencedConfiguration.ToLower() == "default" ?
                    component.GetTitle() :
                    string.Format("{0} [{1}]", component.GetTitle(), component.ReferencedConfiguration);

                list.Add(new Item
                {
                    Name = name,
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
    }

    public static class Helper
    {
        public static Feature GetFeatureByTypeName(this ModelDoc2 model, string featureName)
        {
            var feature = model.FirstFeature() as Feature;

            while (feature != null)
            {
                if (feature.GetTypeName() == featureName)
                    return feature;

                feature = feature.GetNextFeature() as Feature;
            }

            return feature;
        }

        public static List<Feature> GetAllFeaturesByTypeName(this ModelDoc2 model, string featureName)
        {
            var feature = model.FirstFeature() as Feature;
            var list = new List<Feature>();

            while (feature != null)
            {
                if (feature.GetTypeName() == featureName)
                    list.Add(feature);

                feature = feature.GetNextFeature() as Feature;
            }

            return list;
        }

        public static bool HasFlatPattern(this ModelDoc2 model)
        {
            return model.GetBendState() != (int)swSMBendState_e.swSMBendStateNone;
        }

        public static bool IsSheetMetal(this ModelDoc2 model)
        {
            if (model is PartDoc == false)
                return false;

            if (model.HasFlatPattern() == false)
                return false;

            return true;
        }

        public static bool IsPart(this ModelDoc2 model)
        {
            return model is PartDoc;
        }

        public static bool IsDrawing(this ModelDoc2 model)
        {
            return model is DrawingDoc;
        }

        public static bool IsAssembly(this ModelDoc2 model)
        {
            return model is AssemblyDoc;
        }

        public static string GetTitle(this Component2 component)
        {
            var model = component.GetModelDoc2() as ModelDoc2;
            return model.GetTitle();
        }

        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;

            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }

        public static string ToReadableFormat(this TimeSpan ts)
        {
            var s = new StringBuilder();

            if (ts.TotalHours >= 1)
            {
                var hrs = ts.Hours + ts.Days * 24.0;

                s.Append(string.Format("{0}hrs ", hrs));
                s.Append(string.Format("{0}min ", ts.Minutes));
                s.Append(string.Format("{0}sec", ts.Seconds));
            }
            else if (ts.TotalMinutes >= 1)
            {
                s.Append(string.Format("{0}min ", ts.Minutes));
                s.Append(string.Format("{0}sec", ts.Seconds));
            }
            else
            {
                s.Append(string.Format("{0} seconds", ts.Seconds));
            }

            return s.ToString();
        }

        public static string GetNumWithSuffix(int i)
        {
            if (i >= 11 && i <= 13)
                return i.ToString() + "th";

            var j = i % 10;

            switch (j)
            {
                case 1: return i.ToString() + "st";
                case 2: return i.ToString() + "nd";
                case 3: return i.ToString() + "rd";
                default: return i.ToString() + "th";
            }
        }

        public static int IndexOfColumnType(this TableAnnotation table, swTableColumnTypes_e columnType)
        {
            for (int columnIndex = 0; columnIndex < table.ColumnCount; ++columnIndex)
            {
                var currentColumnType = (swTableColumnTypes_e)table.GetColumnType(columnIndex);

                if (currentColumnType == columnType)
                    return columnIndex;
            }

            return -1;
        }

        public static int IndexOfColumnTitle(this TableAnnotation table, string columnTitle)
        {
            var lowercaseColumnTitle = columnTitle.ToLower();

            for (int columnIndex = 0; columnIndex < table.ColumnCount; ++columnIndex)
            {
                var currentColumnType = table.GetColumnTitle(columnIndex);

                if (currentColumnType.ToLower() == lowercaseColumnTitle)
                    return columnIndex;
            }

            return -1;
        }
    }

    public class Item
    {
        public string Name { get; set; }

        public int Quantity { get; set; }

        public Component2 Component { get; set; }
    }
}
