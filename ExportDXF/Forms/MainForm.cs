using ExportDXF.Extensions;
using ExportDXF.Models;
using ExportDXF.Services;
using ExportDXF.ViewFlipDeciders;
using System;
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
        private readonly ICutFabApiClient _apiClient;
        private CancellationTokenSource _cancellationTokenSource;
        private readonly BindingList<LogEvent> _logEvents;
        private readonly BindingList<BomItem> _bomItems;
        private readonly BindingList<CutFabApiClient.ApiCutTemplate> _cutTemplates;
        private readonly int _selectedDrawingId;
        private readonly string _selectedDrawingNumber;

        public MainForm(ISolidWorksService solidWorksService, IDxfExportService exportService, ICutFabApiClient apiClient, int selectedDrawingId, string selectedDrawingNumber)
        {
            InitializeComponent();
            _solidWorksService = solidWorksService ??
                throw new ArgumentNullException(nameof(solidWorksService));
            _exportService = exportService ??
                throw new ArgumentNullException(nameof(exportService));
            _apiClient = apiClient ??
                throw new ArgumentNullException(nameof(apiClient));
            _selectedDrawingId = selectedDrawingId;
            _selectedDrawingNumber = selectedDrawingNumber ?? throw new ArgumentNullException(nameof(selectedDrawingNumber));
            _logEvents = new BindingList<LogEvent>();
            _bomItems = new BindingList<BomItem>();
            _cutTemplates = new BindingList<CutFabApiClient.ApiCutTemplate>();
            InitializeViewFlipDeciders();
            InitializeLogEventsGrid();
            InitializeBomGrid();
            InitializeCutTemplatesGrid();
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
                LogMessage("Ready", Color.Green);
                UpdateActiveDocumentDisplay();
                runButton.Enabled = true;

                // Load the selected drawing's BOM items
                LogMessage($"Loading BOM items for drawing {_selectedDrawingNumber}...");
                await LoadBomItemsForDrawingAsync(_selectedDrawingId);

                // Load cut templates
                await LoadCutTemplatesAsync();
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to connect to SolidWorks: {ex.Message}", Color.Red);
                MessageBox.Show("Failed to connect to SolidWorks.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }
        private async Task LoadBomItemsForDrawingAsync(int drawingId)
        {
            try
            {
                _bomItems.Clear();

                var apiBomItems = await _apiClient.GetBomItemsForDrawingAsync(drawingId);

                foreach (var apiItem in apiBomItems)
                {
                    var bomItem = new BomItem
                    {
                        ID = apiItem.ID,
                        ItemNo = apiItem.ItemNo ?? "",
                        PartNo = apiItem.PartNo ?? "",
                        SortOrder = apiItem.SortOrder,
                        Qty = apiItem.Qty,
                        TotalQty = apiItem.TotalQty,
                        Description = apiItem.Description ?? "",
                        PartName = apiItem.PartName ?? "",
                        ConfigurationName = apiItem.ConfigurationName ?? "",
                        Material = apiItem.Material ?? "",
                        DrawingID = apiItem.DrawingID,
                        CutTemplateID = apiItem.CutTemplateID,
                        CutTemplateName = apiItem.CutTemplateName ?? "",
                        Thickness = apiItem.Thickness,
                        KFactor = apiItem.KFactor,
                        DefaultBendRadius = apiItem.DefaultBendRadius
                    };

                    _bomItems.Add(bomItem);
                }

                LogMessage($"Loaded {apiBomItems.Count} BOM item(s)", Color.Green);
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to load BOM items: {ex.Message}", Color.Red);
            }
        }

        private async Task LoadCutTemplatesAsync()
        {
            try
            {
                _cutTemplates.Clear();

                LogMessage("Loading cut templates...");

                var apiCutTemplates = await _apiClient.GetCutTemplatesAsync();

                foreach (var template in apiCutTemplates)
                {
                    _cutTemplates.Add(template);
                }

                LogMessage($"Loaded {apiCutTemplates.Count} cut template(s)", Color.Green);
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to load cut templates: {ex.Message}", Color.Red);
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
                DataPropertyName = nameof(LogEvent.Message),
                HeaderText = "Message",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });

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

            bomDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(BomItem.CutTemplateName),
                HeaderText = "Cut Template",
                Width = 120
            });

            // Set the data source AFTER adding columns
            bomDataGrid.DataSource = _bomItems;
        }

        private void InitializeCutTemplatesGrid()
        {
            // Clear any existing columns first
            cutTemplatesDataGrid.Columns.Clear();

            // Configure grid settings
            cutTemplatesDataGrid.AutoGenerateColumns = false;
            cutTemplatesDataGrid.AllowUserToAddRows = false;
            cutTemplatesDataGrid.AllowUserToDeleteRows = false;
            cutTemplatesDataGrid.ReadOnly = true;
            cutTemplatesDataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            // Add columns
            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutFabApiClient.ApiCutTemplate.Name),
                HeaderText = "Name",
                Width = 200
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutFabApiClient.ApiCutTemplate.Material),
                HeaderText = "Material",
                Width = 120
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutFabApiClient.ApiCutTemplate.Thickness),
                HeaderText = "Thickness",
                Width = 80,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "0.##" }
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutFabApiClient.ApiCutTemplate.KFactor),
                HeaderText = "K-Factor",
                Width = 80,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "0.####" }
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutFabApiClient.ApiCutTemplate.DefaultBendRadius),
                HeaderText = "Bend Radius",
                Width = 90,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "0.##" }
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutFabApiClient.ApiCutTemplate.Version),
                HeaderText = "Version",
                Width = 70
            });

            cutTemplatesDataGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(CutFabApiClient.ApiCutTemplate.Description),
                HeaderText = "Description",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });

            // Set the data source AFTER adding columns
            cutTemplatesDataGrid.DataSource = _cutTemplates;
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
                var exportContext = new ExportContext
                {
                    ActiveDocument = activeDoc,
                    ViewFlipDecider = viewFlipDecider,
                    FilePrefix = _selectedDrawingNumber,
                    EquipmentId = null, // Not needed for export
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
            runButton.Image = Properties.Resources.stop_alt;
        }
        private void UpdateUIForExportComplete()
        {
            viewFlipDeciderBox.Enabled = true;
            runButton.Image = Properties.Resources.play;
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
        }
        private void LogMessage(string message, Color? color = null)
        {
            var level = color.HasValue && color.Value == Color.Red ? LogLevel.Error :
                        color.HasValue && color.Value == Color.Green ? LogLevel.Info :
                        color.HasValue && color.Value == Color.DarkBlue ? LogLevel.Warning :
                        LogLevel.Info;

            AddLogEvent(level, LogAction.Start, message);
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
    }
}