using OfficeOpenXml;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
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

            viewFlipDecider = new ViewFlipDecider();
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

            sldWorks.UserControl = false;
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
			var prefix = textBox2.Text;

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
                    break;

				item.ItemNo = prefix + item.ItemNo;

				var fileName = item.ItemNo + ".dxf";
                var savepath = Path.Combine(savePath, fileName);
				var model = item.Component.GetModelDoc2() as ModelDoc2;
				var part = model as PartDoc;

				if (part == null)
				{
					Print(model.GetTitle() + " - skipped, not a part document");
					continue;
				}

				var config = item.Component.ReferencedConfiguration;

				var sheetMetal = model.GetFeatureByTypeName("SheetMetal");
				var sheetMetalData = sheetMetal?.GetDefinition() as SheetMetalFeatureData;

				if (sheetMetalData != null)
				{
					item.Thickness = sheetMetalData.Thickness.FromSldWorks();
					item.KFactor = sheetMetalData.KFactor;
				}

				var db = string.Empty;

				item.Material = part.GetMaterialPropertyName2(config, out db);

				if (part == null)
                    continue;

                SavePartToDXF(part, config, savepath);
                Application.DoEvents();
            }

			try
			{
				var bomFile = Path.Combine(savePath, "BOM.xlsx");
				CreateBOMExcelFile(bomFile, items.ToList());
			}
			catch (Exception ex)
			{
				Print(ex.Message, Color.Red);
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

					partsSheet.Cells[row, 1].Value = item.ItemNo;
					partsSheet.Cells[row, 2].Value = item.Quantity;
					partsSheet.Cells[row, 3].Value = item.Description;
					partsSheet.Cells[row, 4].Value = item.PartNo;

					if (item.Thickness > 0)
						partsSheet.Cells[row, 5].Value = item.Thickness;

					partsSheet.Cells[row, 6].Value = item.Material;

					if (item.KFactor > 0)
						partsSheet.Cells[row, 7].Value = item.KFactor;
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
                    .Select(group => group.First());

                //var itemNumber = table.Text[rowIndex, itemNoColumnIndex].PadLeft(2, '0');
                //var rev = 'A';

                if (distinctComponents.Count() > 1)
                {
					throw new NotImplementedException();

                    //foreach (var comp in distinctComponents)
                    //{
                    //    items.Add(new Item
                    //    {
                    //        Name = itemNumber + rev++,
                    //        Component = comp
                    //    });
                    //}
                }
                else
                {
					items.Add(new Item
					{
						PartNo = table.DisplayedText[rowIndex, partNoColumnIndex],
						Quantity = int.Parse(table.DisplayedText[rowIndex, qtyColumnIndex]),
						Description = table.DisplayedText[rowIndex, descriptionColumnIndex],
						ItemNo = table.DisplayedText[rowIndex, itemNoColumnIndex].PadLeft(2, '0'),
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
                    PartNo = name,
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

	
}
