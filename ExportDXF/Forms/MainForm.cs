using ExportDXF.ApiClient;
using ExportDXF.Extensions;
using ExportDXF.Models;
using ExportDXF.Services;
using ExportDXF.ViewFlipDeciders;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        private readonly IFabWorksApiClient _apiClient;
        private CancellationTokenSource _cancellationTokenSource;
        private readonly BindingList<LogEvent> _logEvents;
        private readonly BindingList<BomItem> _bomItems;
        private readonly BindingList<CutTemplate> _cutTemplates;
        private List<DrawingInfo> _allDrawings;

        public MainForm(ISolidWorksService solidWorksService, IDxfExportService exportService, IFabWorksApiClient apiClient)
        {
            InitializeComponent();
            _solidWorksService = solidWorksService ??
                throw new ArgumentNullException(nameof(solidWorksService));
            _solidWorksService.ActiveDocumentChanged += OnActiveDocumentChanged;
            _exportService = exportService ??
                throw new ArgumentNullException(nameof(exportService));
            _apiClient = apiClient ??
                throw new ArgumentNullException(nameof(apiClient));
            _logEvents = new BindingList<LogEvent>();
            _bomItems = new BindingList<BomItem>();
            _cutTemplates = new BindingList<CutTemplate>();
            _allDrawings = new List<DrawingInfo>();
            InitializeViewFlipDeciders();
            InitializeLogEventsGrid();
            InitializeBomGrid();
            InitializeCutTemplatesGrid();
            InitializeDrawingDropdowns();
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
            runButton.Enabled = false;
            await InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            try
            {
                LogMessage("Connecting to SolidWorks, this may take a minute...");
                await _solidWorksService.ConnectAsync();
                _solidWorksService.ActiveDocumentChanged += OnActiveDocumentChanged;
                LogMessage("Files will be uploaded to FabWorks API");
                await LoadDrawingDropdownsAsync();
                LogMessage("Ready");
                await UpdateActiveDocumentDisplayAsync();
                runButton.Enabled = true;
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to connect to SolidWorks: {ex.Message}", LogLevel.Error);
                MessageBox.Show("Failed to connect to SolidWorks.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
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
            viewFlipDeciderBox.DataSource = items;
            viewFlipDeciderBox.DisplayMember = "Name";
        }

        private void InitializeLogEventsGrid()
        {
            // Clear any existing columns first
            logEventsDataGrid.Columns.Clear();

            // Configure grid settings
            logEventsDataGrid.AutoGenerateColumns = false;
            logEventsDataGrid.AllowUserToAddRows = false;
            logEventsDataGrid.AllowUserToDeleteRows = false;
            logEventsDataGrid.ReadOnly = true;
            logEventsDataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            // Add columns
            logEventsDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(LogEvent.Time),
                HeaderText = "Time",
                Width = 80,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "HH:mm:ss" }
            });

            logEventsDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(LogEvent.Level),
                HeaderText = "Level",
                Width = 70
            });

            logEventsDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(LogEvent.Part),
                HeaderText = "File",
                Width = 180
            });

            logEventsDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(LogEvent.Message),
                HeaderText = "Message",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });

            // Add row coloring based on log level
            logEventsDataGrid.CellFormatting += LogEventsDataGrid_CellFormatting;

            // Set the data source AFTER adding columns
            logEventsDataGrid.DataSource = _logEvents;
        }

        private void InitializeBomGrid()
        {
            // Clear any existing columns first
            bomDataGrid.Columns.Clear();

            // Configure grid settings
            bomDataGrid.AutoGenerateColumns = false;
            bomDataGrid.AllowUserToAddRows = false;
            bomDataGrid.AllowUserToDeleteRows = false;
            bomDataGrid.ReadOnly = true;
            bomDataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            // Add columns
            bomDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(BomItem.ItemNo),
                HeaderText = "Item #",
                Width = 60
            });

            bomDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(BomItem.PartNo),
                HeaderText = "Part #",
                Width = 100
            });

            bomDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(BomItem.Qty),
                HeaderText = "Qty",
                Width = 50
            });

            bomDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(BomItem.Description),
                HeaderText = "Description",
                Width = 200
            });

            bomDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(BomItem.PartName),
                HeaderText = "Part Name",
                Width = 150
            });

            bomDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(BomItem.ConfigurationName),
                HeaderText = "Configuration",
                Width = 120
            });

            bomDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(BomItem.Material),
                HeaderText = "Material",
                Width = 120
            });

            // Set the data source AFTER adding columns
            bomDataGrid.DataSource = _bomItems;
        }

        private void InitializeCutTemplatesGrid()
        {
            cutTemplatesDataGrid.Columns.Clear();

            cutTemplatesDataGrid.AutoGenerateColumns = false;
            cutTemplatesDataGrid.AllowUserToAddRows = false;
            cutTemplatesDataGrid.AllowUserToDeleteRows = false;
            cutTemplatesDataGrid.ReadOnly = true;
            cutTemplatesDataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutTemplate.CutTemplateName),
                HeaderText = "Template Name",
                Width = 150
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutTemplate.DxfFilePath),
                HeaderText = "DXF File",
                Width = 250
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutTemplate.Thickness),
                HeaderText = "Thickness",
                Width = 80
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutTemplate.KFactor),
                HeaderText = "K-Factor",
                Width = 80
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutTemplate.DefaultBendRadius),
                HeaderText = "Bend Radius",
                Width = 90
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutTemplate.ContentHash),
                HeaderText = "Content Hash",
                Width = 150
            });

            cutTemplatesDataGrid.DataSource = _cutTemplates;
        }

        private void InitializeDrawingDropdowns()
        {
            // Wire up event handler; actual data loading happens in LoadDrawingDropdownsAsync
            equipmentBox.SelectedIndexChanged += EquipmentBox_SelectedIndexChanged;
        }

        private async Task LoadDrawingDropdownsAsync()
        {
            try
            {
                var equipmentNumbers = await _apiClient.GetEquipmentNumbersAsync();

                equipmentBox.Items.Clear();
                equipmentBox.Items.Add("");
                foreach (var eq in equipmentNumbers)
                {
                    equipmentBox.Items.Add(eq);
                }

                // Clear _allDrawings — drawing list is now loaded on equipment selection
                _allDrawings = new List<DrawingInfo>();
                await UpdateDrawingDropdownAsync();
            }
            catch (Exception ex)
            {
                // API might not be available yet - that's OK
                System.Diagnostics.Debug.WriteLine($"Failed to load equipment numbers from API: {ex.Message}");
            }
        }

        private async void EquipmentBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            await UpdateDrawingDropdownAsync();
        }

        private async Task UpdateDrawingDropdownAsync()
        {
            var selectedEquipment = equipmentBox.SelectedItem?.ToString();

            drawingNoBox.Items.Clear();
            drawingNoBox.Items.Add("");

            try
            {
                var drawingNumbers = await _apiClient.GetDrawingNumbersByEquipmentAsync(
                    string.IsNullOrEmpty(selectedEquipment) ? null : selectedEquipment);

                foreach (var dn in drawingNumbers)
                {
                    drawingNoBox.Items.Add(dn);
                }
            }
            catch
            {
                // API might not be available
            }

            if (drawingNoBox.Items.Count > 0)
            {
                drawingNoBox.SelectedIndex = 0;
            }
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
                    LogMessage("No active document.", LogLevel.Error);
                    return;
                }

                // Use equipment/drawing values from the UI dropdowns
                var equipment = equipmentBox.Text?.Trim();
                var drawingNo = drawingNoBox.Text?.Trim();
                var filePrefix = !string.IsNullOrEmpty(equipment)
                    ? (!string.IsNullOrEmpty(drawingNo) ? $"{equipment} {drawingNo}" : equipment)
                    : activeDoc.Title;

                var viewFlipDecider = GetSelectedViewFlipDecider();
                var title = titleBox.Text?.Trim();

                var exportContext = new ExportContext
                {
                    ActiveDocument = activeDoc,
                    ViewFlipDecider = viewFlipDecider,
                    FilePrefix = filePrefix,
                    Equipment = equipment,
                    DrawingNo = drawingNo,
                    Title = string.IsNullOrEmpty(title) ? null : title,
                    EquipmentId = null,
                    CancellationToken = token,
                    ProgressCallback = (msg, level, file) => LogMessage(msg, level, file),
                    BomItemCallback = AddBomItem
                };

                // Clear previous BOM items and cut templates
                _bomItems.Clear();
                _cutTemplates.Clear();

                LogMessage($"Started at {DateTime.Now:t}");
                LogMessage("Exporting (files will be uploaded to API)...");

                _solidWorksService.SetCommandInProgress(true);

                await Task.Run(async () => await _exportService.ExportAsync(exportContext), token);

                LogMessage("Done.");
            }
            catch (OperationCanceledException)
            {
                LogMessage("Export canceled.", LogLevel.Error);
            }
            catch (Exception ex)
            {
                LogMessage($"Export failed: {ex.Message}", LogLevel.Error);
                MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _solidWorksService.SetCommandInProgress(false);
                UpdateUIForExportComplete();
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }

        private void CancelExport()
        {
            runButton.Enabled = false;
            _cancellationTokenSource?.Cancel();
        }

        private IViewFlipDecider GetSelectedViewFlipDecider()
        {
            var item = viewFlipDeciderBox.SelectedItem as ViewFlipDeciderComboboxItem;
            return item?.ViewFlipDecider;
        }

        private void UpdateUIForExportStart()
        {
            viewFlipDeciderBox.Enabled = false;
            runButton.Text = "Stop";
        }

        private void UpdateUIForExportComplete()
        {
            viewFlipDeciderBox.Enabled = true;
            runButton.Text = "Start";
            runButton.Enabled = true;
        }

        private async void OnActiveDocumentChanged(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(async () => await UpdateActiveDocumentDisplayAsync()));
                return;
            }
            await UpdateActiveDocumentDisplayAsync();
        }

        private async Task UpdateActiveDocumentDisplayAsync()
        {
            var activeDoc = _solidWorksService.GetActiveDocument();
            var docTitle = activeDoc?.Title ?? "No Document Open";
            this.Text = $"ExportDXF - {docTitle}";

            if (activeDoc == null)
                return;

            // Try API first: look up the most recent export for this file path
            DrawingInfo drawingInfo = null;

            if (!string.IsNullOrEmpty(activeDoc.FilePath))
            {
                drawingInfo = await LookupDrawingInfoFromHistoryAsync(activeDoc.FilePath);
            }

            // Fall back to parsing the document title
            if (drawingInfo == null)
            {
                drawingInfo = DrawingInfo.Parse(activeDoc.Title);
            }

            if (drawingInfo != null)
            {
                // Detach event to prevent async race when setting equipment
                equipmentBox.SelectedIndexChanged -= EquipmentBox_SelectedIndexChanged;

                if (!string.IsNullOrEmpty(drawingInfo.EquipmentNo))
                {
                    if (!equipmentBox.Items.Contains(drawingInfo.EquipmentNo))
                        equipmentBox.Items.Add(drawingInfo.EquipmentNo);
                    equipmentBox.Text = drawingInfo.EquipmentNo;
                }

                // Load drawings for the selected equipment, then set drawing number
                await UpdateDrawingDropdownAsync();

                equipmentBox.SelectedIndexChanged += EquipmentBox_SelectedIndexChanged;

                if (!string.IsNullOrEmpty(drawingInfo.DrawingNo))
                {
                    if (!drawingNoBox.Items.Contains(drawingInfo.DrawingNo))
                        drawingNoBox.Items.Add(drawingInfo.DrawingNo);
                    drawingNoBox.Text = drawingInfo.DrawingNo;
                }
            }
        }

        private async Task<DrawingInfo> LookupDrawingInfoFromHistoryAsync(string filePath)
        {
            try
            {
                var dto = await _apiClient.GetExportBySourceFileAsync(filePath);
                if (dto != null && !string.IsNullOrEmpty(dto.DrawingNumber))
                {
                    if (!string.IsNullOrEmpty(dto.Title))
                        titleBox.Text = dto.Title;
                    return DrawingInfo.Parse(dto.DrawingNumber);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to look up drawing info from API: {ex.Message}");
            }

            return null;
        }

        private void LogMessage(string message, LogLevel level = LogLevel.Info, string file = null)
        {
            AddLogEvent(level, LogAction.Start, message, part: file);
        }

        private void AddLogEvent(LogLevel level, LogAction action, string message, string equipment = "", string drawing = "", string part = "", string target = "", string result = "OK", int durationMs = 0)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => AddLogEvent(level, action, message, equipment, drawing, part, target, result, durationMs)));
                return;
            }

            var logEvent = new LogEvent
            {
                Time = DateTime.Now,
                Level = level,
                Action = action,
                Message = message,
                Equipment = equipment,
                Drawing = drawing,
                Part = part,
                Target = target,
                Result = result,
                DurationMs = durationMs
            };

            _logEvents.Add(logEvent);

            // Auto-scroll to the last row
            if (logEventsDataGrid.Rows.Count > 0)
            {
                logEventsDataGrid.FirstDisplayedScrollingRowIndex = logEventsDataGrid.Rows.Count - 1;
            }
        }

        public void AddBomItem(BomItem item)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => AddBomItem(item)));
                return;
            }
            _bomItems.Add(item);

            if (item.CutTemplate != null)
            {
                _cutTemplates.Add(item.CutTemplate);
            }
        }

        private void LogEventsDataGrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _logEvents.Count)
                return;

            var logEvent = _logEvents[e.RowIndex];
            var row = logEventsDataGrid.Rows[e.RowIndex];

            switch (logEvent.Level)
            {
                case LogLevel.Warning:
                    row.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow;
                    row.DefaultCellStyle.ForeColor = Color.DarkOrange;
                    break;
                case LogLevel.Error:
                    row.DefaultCellStyle.BackColor = Color.MistyRose;
                    row.DefaultCellStyle.ForeColor = Color.DarkRed;
                    break;
                default:
                    row.DefaultCellStyle.BackColor = Color.White;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                    break;
            }
        }
    }
}
