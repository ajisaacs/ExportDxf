using ExportDXF.Data;
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
        private readonly IFileExportService _fileExportService;
        private readonly Func<ExportDxfDbContext> _dbContextFactory;
        private CancellationTokenSource _cancellationTokenSource;
        private readonly BindingList<LogEvent> _logEvents;
        private readonly BindingList<BomItem> _bomItems;
        private readonly BindingList<CutTemplate> _cutTemplates;
        private List<DrawingInfo> _allDrawings;

        public MainForm(ISolidWorksService solidWorksService, IDxfExportService exportService, IFileExportService fileExportService, Func<ExportDxfDbContext> dbContextFactory = null)
        {
            InitializeComponent();
            _solidWorksService = solidWorksService ??
                throw new ArgumentNullException(nameof(solidWorksService));
            _solidWorksService.ActiveDocumentChanged += OnActiveDocumentChanged;
            _exportService = exportService ??
                throw new ArgumentNullException(nameof(exportService));
            _fileExportService = fileExportService ??
                throw new ArgumentNullException(nameof(fileExportService));
            _dbContextFactory = dbContextFactory ?? (() => new ExportDxfDbContext());
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
                LogMessage($"Output folder: {_fileExportService.OutputFolder}");
                LogMessage("Ready");
                UpdateActiveDocumentDisplay();
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
            try
            {
                using (var db = _dbContextFactory())
                {
                    // Get all drawing numbers from the database
                    var drawingNumbers = db.ExportRecords
                        .Select(r => r.DrawingNumber)
                        .Where(d => !string.IsNullOrEmpty(d))
                        .Distinct()
                        .ToList();

                    // Parse into DrawingInfo objects
                    _allDrawings = drawingNumbers
                        .Select(DrawingInfo.Parse)
                        .Where(d => d != null)
                        .Distinct()
                        .OrderBy(d => d.EquipmentNo)
                        .ThenBy(d => d.DrawingNo)
                        .ToList();

                    // Get distinct equipment numbers
                    var equipmentNumbers = _allDrawings
                        .Select(d => d.EquipmentNo)
                        .Distinct()
                        .OrderBy(e => e)
                        .ToList();

                    // Populate equipment dropdown
                    equipmentBox.Items.Clear();
                    equipmentBox.Items.Add(""); // Empty option for "all"
                    foreach (var eq in equipmentNumbers)
                    {
                        equipmentBox.Items.Add(eq);
                    }

                    // Populate drawing dropdown with all drawings initially
                    UpdateDrawingDropdown();

                    // Wire up event handler for equipment selection change
                    equipmentBox.SelectedIndexChanged += EquipmentBox_SelectedIndexChanged;
                }
            }
            catch (Exception ex)
            {
                // Database might not exist yet - that's OK
                System.Diagnostics.Debug.WriteLine($"Failed to load drawings from database: {ex.Message}");
            }
        }

        private void EquipmentBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDrawingDropdown();
        }

        private void UpdateDrawingDropdown()
        {
            var selectedEquipment = equipmentBox.SelectedItem?.ToString();

            var filteredDrawings = string.IsNullOrEmpty(selectedEquipment)
                ? _allDrawings
                : _allDrawings.Where(d => d.EquipmentNo == selectedEquipment).ToList();

            drawingNoBox.Items.Clear();
            drawingNoBox.Items.Add(""); // Empty option
            foreach (var drawing in filteredDrawings)
            {
                drawingNoBox.Items.Add(drawing.DrawingNo);
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
                var filePrefix = !string.IsNullOrEmpty(equipment) && !string.IsNullOrEmpty(drawingNo)
                    ? $"{equipment} {drawingNo}"
                    : activeDoc.Title;

                var viewFlipDecider = GetSelectedViewFlipDecider();
                var exportContext = new ExportContext
                {
                    ActiveDocument = activeDoc,
                    ViewFlipDecider = viewFlipDecider,
                    FilePrefix = filePrefix,
                    EquipmentId = null,
                    CancellationToken = token,
                    ProgressCallback = (msg, level, file) => LogMessage(msg, level, file),
                    BomItemCallback = AddBomItem
                };

                // Clear previous BOM items and cut templates
                _bomItems.Clear();
                _cutTemplates.Clear();

                LogMessage($"Started at {DateTime.Now:t}");
                LogMessage($"Exporting to: {_fileExportService.OutputFolder}");

                _solidWorksService.SetCommandInProgress(true);

                await Task.Run(() => _exportService.Export(exportContext), token);

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

        private void OnActiveDocumentChanged(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => OnActiveDocumentChanged(sender, e)));
                return;
            }
            UpdateActiveDocumentDisplay();
        }

        private void UpdateActiveDocumentDisplay()
        {
            var activeDoc = _solidWorksService.GetActiveDocument();
            var docTitle = activeDoc?.Title ?? "No Document Open";
            this.Text = $"ExportDXF - {docTitle}";

            // Parse the file name and fill Equipment/Drawing dropdowns
            if (activeDoc != null)
            {
                var drawingInfo = DrawingInfo.Parse(activeDoc.Title);
                if (drawingInfo != null)
                {
                    if (!equipmentBox.Items.Contains(drawingInfo.EquipmentNo))
                        equipmentBox.Items.Add(drawingInfo.EquipmentNo);
                    equipmentBox.Text = drawingInfo.EquipmentNo;

                    if (!drawingNoBox.Items.Contains(drawingInfo.DrawingNo))
                        drawingNoBox.Items.Add(drawingInfo.DrawingNo);
                    drawingNoBox.Text = drawingInfo.DrawingNo;
                }
            }
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
