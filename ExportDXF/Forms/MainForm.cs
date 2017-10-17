using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

    public interface IViewFlipDecider
    {
        bool ShouldFlip(SolidWorks.Interop.sldworks.View view);
    }

    public class ViewFlipDecider : IViewFlipDecider
    {
        public bool ShouldFlip(SolidWorks.Interop.sldworks.View view)
        {
            var orientation = GetOrientation(view);
			var bounds = GetBounds(view);
			var bends = GetBends(view);

			var up = bends.Where(b => b.Direction == BendDirection.Up).ToList();
			var down = bends.Where(b => b.Direction == BendDirection.Down).ToList();

			if (down.Count == 0)
				return false;

			if (up.Count == 0)
				return true;

			var bend = ClosestToBounds(bounds, bends);

			return bend.Direction == BendDirection.Down;
		}

		private static Bounds GetBounds(SolidWorks.Interop.sldworks.View view)
		{
			var outline = view.GetOutline() as double[];

			var minX = outline[0] / 0.0254;
			var minY = outline[1] / 0.0254;
			var maxX = outline[2] / 0.0254;
			var maxY = outline[3] / 0.0254;

			var width = Math.Abs(minX) + Math.Abs(maxX);
			var height = Math.Abs(minY) + Math.Abs(maxY);

			return new Bounds
			{
				X = minX,
				Y = minY,
				Width = width,
				Height = height
			};
		}

		private static Bend ClosestToBounds(Bounds bounds, IList<Bend> bends)
		{
			var hBends = bends.Where(b => GetAngleOrientation(b.ParallelBendAngle) == BendOrientation.Horizontal).ToList();
			var vBends = bends.Where(b => GetAngleOrientation(b.ParallelBendAngle) == BendOrientation.Vertical).ToList();

			Bend minVBend = null;
			double minVBendDist = double.MaxValue;

			foreach (var bend in vBends)
			{
				double distFromLft = Math.Abs(bend.X - bounds.Left);
				double distFromRgt = Math.Abs(bounds.Right - bend.X);

				double minDist = Math.Min(distFromLft, distFromRgt);

				if (minDist < minVBendDist)
				{
					minVBendDist = minDist;
					minVBend = bend;
				}
			}

			Bend minHBend = null;
			double minHBendDist = double.MaxValue;

			foreach (var bend in hBends)
			{
				double distFromBtm = Math.Abs(bend.Y - bounds.Bottom);
				double distFromTop = Math.Abs(bounds.Top - bend.Y);

				double minDist = Math.Min(distFromBtm, distFromTop);

				if (minDist < minHBendDist)
				{
					minHBendDist = minDist;
					minHBend = bend;
				}
			}

			return minHBendDist < minVBendDist ? minHBend : minVBend;
		}

		private static Bend SmallestYCoordinate(IList<Bend> bends)
		{
			double dist = double.MaxValue;
			int index = -1;

			for (int i = 0; i < bends.Count; i++)
			{
				var bend = bends[i];

				if (bend.Y < dist)
				{
					dist = bend.Y;
					index = i;
				}
			}

			return index == -1 ? null : bends[index];
		}

		private static Bend SmallestXCoordinate(IList<Bend> bends)
		{
			double dist = double.MaxValue;
			int index = -1;

			for (int i = 0; i < bends.Count; i++)
			{
				var bend = bends[i];

				if (bend.X < dist)
				{
					dist = bend.X;
					index = i;
				}
			}

			return index == -1 ? null : bends[index];
		}

		private static BendDirection GetBendDirection(Note note)
        {
            var txt = note.GetText();

            return txt.ToUpper().Contains("UP") ? BendDirection.Up : BendDirection.Down;
        }

        private static IEnumerable<Note> GetBendNotes(SolidWorks.Interop.sldworks.View view)
        {
            return (view.GetNotes() as Array)?.Cast<Note>();
        }

        private static Note GetLeftMostNote(SolidWorks.Interop.sldworks.View view)
        {
            var notes = GetBendNotes(view);

            Note leftMostNote = null;
            var leftMostValue = double.MaxValue;

            foreach (var note in notes)
            {
                var pt = (note.GetTextPoint() as double[]);
                var x = pt[0];

                if (x < leftMostValue)
                {
                    leftMostValue = x;
                    leftMostNote = note;
                }
            }

            return leftMostNote;
        }

        private static Note GetBottomMostNote(SolidWorks.Interop.sldworks.View view)
        {
            var notes = GetBendNotes(view);

            Note btmMostNote = null;
            var btmMostValue = double.MaxValue;

            foreach (var note in notes)
            {
                var pt = (note.GetTextPoint() as double[]);
                var y = pt[1];

                if (y < btmMostValue)
                {
                    btmMostValue = y;
                    btmMostNote = note;
                }
            }

            return btmMostNote;
        }

        private static IEnumerable<double> GetBendAngles(SolidWorks.Interop.sldworks.View view)
        {
            var angles = new List<double>();
            var notes = GetBendNotes(view);

			foreach (var note in notes)
            {
                var angle = RadiansToDegrees(note.Angle);
                angles.Add(angle);
            }

            return angles;
        }

		private static List<Bend> GetBends(SolidWorks.Interop.sldworks.View view)
		{
			var bends = new List<Bend>();
			var notes = GetBendNotes(view);

			const string pattern = @"(?<DIRECTION>(UP|DOWN))\s*(?<ANGLE>(\d+))°";

			foreach (var note in notes)
			{
				var pos = note.GetTextPoint2() as double[];

				var x = pos[0] / 0.0254;
				var y = pos[1] / 0.0254;

				var text = note.GetText();
				var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);

				if (!match.Success)
					continue;

				var angle = double.Parse(match.Groups["ANGLE"].Value);
				var direection = match.Groups["DIRECTION"].Value;

				var bend = new Bend
				{
					ParallelBendAngle = RadiansToDegrees(note.Angle),
					Angle = angle,
					Direction = direection == "UP" ? BendDirection.Up : BendDirection.Down,
					X = x,
					Y = y
				};

				bends.Add(bend);
			}

			return bends;
		}

		private static BendOrientation GetOrientation(SolidWorks.Interop.sldworks.View view)
        {
            var angles = GetBendAngles(view);

			var bends = GetBends(view);

            var vertical = 0;
            var horizontal = 0;

            foreach (var angle in angles)
            {
                var o = GetAngleOrientation(angle);

                switch (o)
                {
                    case BendOrientation.Horizontal:
                        horizontal++;
                        break;

                    case BendOrientation.Vertical:
                        vertical++;
                        break;
                }
            }

            if (vertical == 0 && horizontal == 0)
                return BendOrientation.Unknown;

            return vertical > horizontal ? BendOrientation.Vertical : BendOrientation.Horizontal;
        }

        private static BendOrientation GetAngleOrientation(double angleInDegrees)
        {
            if (angleInDegrees < 10 || angleInDegrees > 350)
                return BendOrientation.Horizontal;

            if (angleInDegrees > 170 && angleInDegrees < 190)
                return BendOrientation.Horizontal;

            if (angleInDegrees > 80 && angleInDegrees < 100)
                return BendOrientation.Vertical;

            if (angleInDegrees > 260 && angleInDegrees < 280)
                return BendOrientation.Vertical;

            return BendOrientation.Unknown;
        }

        private static double RadiansToDegrees(double angleInRadians)
        {
            return Math.Round(angleInRadians * 180.0 / Math.PI, 8);
        }
    }

	class Bend
	{
		public BendDirection Direction { get; set; }

		public double ParallelBendAngle { get; set; }

		public double Angle { get; set; }

		public double X { get; set; }
		public double Y { get; set; }
	}

    enum BendDirection
    {
        Up,
        Down
    }

    enum BendOrientation
    {
        Vertical,
        Horizontal,
        Unknown
    }

	class Size
	{
		public double Width { get; set; }
		public double Height { get; set; }
	}

	class Bounds
	{
		public double X { get; set; }
		public double Y { get; set; }
		public double Width { get; set; }
		public double Height { get; set; }

		public double Left
		{
			get { return X; }
		}

		public double Right
		{
			get { return X + Width; }
		}

		public double Bottom
		{
			get { return Y; }
		}

		public double Top
		{
			get { return Y + Height; }
		}
	}
}
