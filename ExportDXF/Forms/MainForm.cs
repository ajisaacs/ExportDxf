using ExportDXF.Extensions;
using ExportDXF.Models;
using ExportDXF.Services;
using ExportDXF.ViewFlipDeciders;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace ExportDXF.Forms
{
    public partial class MainForm : Form
    {
        private readonly ISolidWorksService _solidWorksService;
        private readonly IDxfExportService _exportService;
        private readonly ICutFabApiClient _apiClient;
        private CancellationTokenSource _cancellationTokenSource;
        public MainForm(ISolidWorksService solidWorksService, IDxfExportService exportService, ICutFabApiClient apiClient)
        {
            InitializeComponent();
            _solidWorksService = solidWorksService ??
                throw new ArgumentNullException(nameof(solidWorksService));
            _exportService = exportService ??
                throw new ArgumentNullException(nameof(exportService));
            _apiClient = apiClient ??
                throw new ArgumentNullException(nameof(apiClient));
            InitializeViewFlipDeciders();
        }
        ~MainForm()
        {
            _cancellationTokenSource?.Dispose();
            _solidWorksService?.Dispose();
            components?.Dispose();
            Dispose(false);
        }
        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            button1.Enabled = false;
            await InitializeAsync();
        }
        private async Task InitializeAsync()
        {
            try
            {
                LogMessage("Connecting to SolidWorks, this may take a minute...");
                await _solidWorksService.ConnectAsync();
                _solidWorksService.ActiveDocumentChanged += OnActiveDocumentChanged;
                LogMessage("Ready", Color.Green);
                UpdateActiveDocumentDisplay();
                button1.Enabled = true;
                // Populate equipment and (initial) drawings
                await PopulateEquipmentAsync();
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to connect to SolidWorks: {ex.Message}", Color.Red);
                MessageBox.Show("Failed to connect to SolidWorks.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }
        private async Task PopulateEquipmentAsync()
        {
            try
            {
                LogMessage($"Loading equipment from {_apiClient.BaseUrl} ...");
                var list = await _apiClient.GetEquipmentAsync();
                equipmentNoBox.DisplayMember = nameof(CutFabApiClient.ApiEquipment.EquipmentNumber);
                equipmentNoBox.ValueMember = nameof(CutFabApiClient.ApiEquipment.ID);
                equipmentNoBox.DataSource = list;
                equipmentNoBox.SelectedIndexChanged -= equipmentNoBox_SelectedIndexChanged;
                equipmentNoBox.SelectedIndexChanged += equipmentNoBox_SelectedIndexChanged;
                if (list != null && list.Count > 0)
                {
                    equipmentNoBox.SelectedIndex = 0;
                    await PopulateDrawingsForSelectedEquipmentAsync();
                    LogMessage($"Loaded {list.Count} equipment record(s)", Color.Green);
                }
                else
                {
                    LogMessage("No equipment returned by API", Color.DarkBlue);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to load equipment: {ex.Message}", Color.Red);
            }
        }
        private async void equipmentNoBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            await PopulateDrawingsForSelectedEquipmentAsync();
        }
        private async Task PopulateDrawingsForSelectedEquipmentAsync()
        {
            try
            {
                var selected = equipmentNoBox.SelectedItem as CutFabApiClient.ApiEquipment;
                if (selected == null)
                {
                    comboBox2.DataSource = null;
                    return;
                }
                var drawings = await _apiClient.GetDrawingsForEquipmentAsync(selected.ID);
                comboBox2.DisplayMember = nameof(CutFabApiClient.ApiDrawingSummary.DrawingNumber);
                comboBox2.ValueMember = nameof(CutFabApiClient.ApiDrawingSummary.ID);
                comboBox2.DataSource = drawings;
                LogMessage($"Loaded {drawings.Count} drawing(s) for equipment {selected.EquipmentNumber}");
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to load drawings: {ex.Message}", Color.Red);
            }
        }
        private void InitializeViewFlipDeciders()
        {
            var items = ViewFlipDeciderFactory.GetAvailableDeciders()
                .Select(d => new ViewFlipDeciderComboboxItem
                {
                    Name = d.Name,
                    ViewFlipDecider = d
                })
                .ToList();
            // Move "Automatic" to the top if it exists
            var automatic = items.FirstOrDefault(i => i.Name == "Automatic");
            if (automatic != null)
            {
                items.Remove(automatic);
                items.Insert(0, automatic);
            }
            comboBox1.DataSource = items;
            comboBox1.DisplayMember = "Name";
        }
        private async void button1_Click(object sender, EventArgs e)
        {
            if (_cancellationTokenSource != null)
            {
                CancelExport();
            }
            else
            {
                await StartExportAsync();
            }
        }
        private async Task StartExportAsync()
        {
            try
            {
                _cancellationTokenSource = new CancellationTokenSource();
                var token = _cancellationTokenSource.Token;
                UpdateUIForExportStart();
                var activeDoc = _solidWorksService.GetActiveDocument();
                if (activeDoc == null)
                {
                    LogMessage("No active document.", Color.Red);
                    return;
                }
                var viewFlipDecider = GetSelectedViewFlipDecider();
                var drawingNumberText = comboBox2.Text?.Trim();
                var selectedEquipment = equipmentNoBox.SelectedItem as CutFabApiClient.ApiEquipment;
                var exportContext = new ExportContext
                {
                    ActiveDocument = activeDoc,
                    ViewFlipDecider = viewFlipDecider,
                    FilePrefix = drawingNumberText,
                    EquipmentId = selectedEquipment?.ID,
                    CancellationToken = token,
                    ProgressCallback = LogMessage
                };
                LogMessage($"Started at {DateTime.Now:t}");
                await Task.Run(() => _exportService.Export(exportContext), token);
                LogMessage("Done.", Color.Green);
            }
            catch (OperationCanceledException)
            {
                LogMessage("Export canceled.", Color.Red);
            }
            catch (Exception ex)
            {
                LogMessage($"Export failed: {ex.Message}", Color.Red);
                MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UpdateUIForExportComplete();
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }
        private void CancelExport()
        {
            button1.Enabled = false;
            _cancellationTokenSource?.Cancel();
        }
        private IViewFlipDecider GetSelectedViewFlipDecider()
        {
            var item = comboBox1.SelectedItem as ViewFlipDeciderComboboxItem;
            return item?.ViewFlipDecider;
        }
        private void UpdateUIForExportStart()
        {
            activeDocTitleBox.Enabled = false;
            comboBox1.Enabled = false;
            button1.Image = Properties.Resources.stop_alt;
            if (richTextBox1.TextLength != 0)
            {
                richTextBox1.AppendText("\n\n");
            }
        }
        private void UpdateUIForExportComplete()
        {
            activeDocTitleBox.Enabled = true;
            comboBox1.Enabled = true;
            button1.Image = Properties.Resources.play;
            button1.Enabled = true;
        }
        private void OnActiveDocumentChanged(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => OnActiveDocumentChanged(sender, e)));
                return;
            }
            UpdateActiveDocumentDisplay();
            UpdatePrefixFromActiveDocument();
        }
        private void UpdateActiveDocumentDisplay()
        {
            var activeDoc = _solidWorksService.GetActiveDocument();
            activeDocTitleBox.Text = activeDoc?.Title ?? "<No Document Open>";
        }
        private void UpdatePrefixFromActiveDocument()
        {
            var activeDoc = _solidWorksService.GetActiveDocument();
            if (activeDoc == null)
            {
                return;
            }
            if (activeDoc.DocumentType == DocumentType.Drawing)
            {
                var drawingInfo = DrawingInfo.Parse(activeDoc.Title);
                if (drawingInfo != null)
                {
                }
            }
            else
            {
            }
        }
        private void activeDocTitleBox_TextChanged(object sender, EventArgs e)
        {
            UpdatePrefixFromActiveDocument();
        }
        private void LogMessage(string message, Color? color = null)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => LogMessage(message, color)));
                return;
            }
            if (color.HasValue)
            {
                richTextBox1.AppendText(message + System.Environment.NewLine, color.Value);
            }
            else
            {
                richTextBox1.AppendText(message + System.Environment.NewLine);
            }
            richTextBox1.ScrollToCaret();
        }
    }
}