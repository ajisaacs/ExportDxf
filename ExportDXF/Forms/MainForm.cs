using ExportDXF.Extensions;
using ExportDXF.Models;
using ExportDXF.Services;
using ExportDXF.ViewFlipDeciders;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
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
        private readonly IDrawingInfoExtractor[] _extractors;
        private CancellationTokenSource _cancellationTokenSource;
        private readonly BindingList<LogEvent> _logEvents;
        private readonly BindingList<BomItem> _bomItems;
        private readonly BindingList<CutTemplate> _cutTemplates;

        public MainForm(ISolidWorksService solidWorksService, IDxfExportService exportService, IDrawingInfoExtractor[] extractors)
        {
            InitializeComponent();
            _solidWorksService = solidWorksService ??
                throw new ArgumentNullException(nameof(solidWorksService));
            _solidWorksService.ActiveDocumentChanged += OnActiveDocumentChanged;
            _exportService = exportService ??
                throw new ArgumentNullException(nameof(exportService));
            _extractors = extractors ??
                throw new ArgumentNullException(nameof(extractors));
            _logEvents = new BindingList<LogEvent>();
            _bomItems = new BindingList<BomItem>();
            _cutTemplates = new BindingList<CutTemplate>();
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
            logEventsDataGrid.Columns.Clear();
            logEventsDataGrid.AutoGenerateColumns = false;
            logEventsDataGrid.AllowUserToAddRows = false;
            logEventsDataGrid.AllowUserToDeleteRows = false;
            logEventsDataGrid.ReadOnly = true;
            logEventsDataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

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

            logEventsDataGrid.CellFormatting += LogEventsDataGrid_CellFormatting;
            logEventsDataGrid.DataSource = _logEvents;
        }

        private void InitializeBomGrid()
        {
            bomDataGrid.Columns.Clear();
            bomDataGrid.AutoGenerateColumns = false;
            bomDataGrid.AllowUserToAddRows = false;
            bomDataGrid.AllowUserToDeleteRows = false;
            bomDataGrid.ReadOnly = true;
            bomDataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

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
                DataPropertyName = nameof(CutTemplate.Revision),
                HeaderText = "Rev",
                Width = 50
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
                var template = txtFilenameTemplate.Text.Trim();
                if (!FilenameTemplateParser.Validate(template, out var error))
                {
                    MessageBox.Show(error, "Invalid Template", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _cancellationTokenSource = new CancellationTokenSource();
                var token = _cancellationTokenSource.Token;
                UpdateUIForExportStart();

                var activeDoc = _solidWorksService.GetActiveDocument();
                if (activeDoc == null)
                {
                    LogMessage("No active document.", LogLevel.Error);
                    return;
                }

                var sourceDir = Path.GetDirectoryName(activeDoc.FilePath);
                var outputFolder = Path.Combine(sourceDir, "Templates");

                var viewFlipDecider = GetSelectedViewFlipDecider();

                var exportContext = new ExportContext
                {
                    ActiveDocument = activeDoc,
                    ViewFlipDecider = viewFlipDecider,
                    FilenameTemplate = template,
                    OutputFolder = outputFolder,
                    CancellationToken = token,
                    ProgressCallback = (msg, level, file) => LogMessage(msg, level, file),
                    BomItemCallback = AddBomItem
                };

                _bomItems.Clear();
                _cutTemplates.Clear();

                LogMessage($"Started at {DateTime.Now:t}");
                LogMessage($"Output: {outputFolder}");

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
            txtFilenameTemplate.Enabled = false;
            runButton.Text = "Stop";
        }

        private void UpdateUIForExportComplete()
        {
            viewFlipDeciderBox.Enabled = true;
            txtFilenameTemplate.Enabled = true;
            runButton.Text = "Start";
            runButton.Enabled = true;
        }

        private void OnActiveDocumentChanged(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateActiveDocumentDisplay()));
                return;
            }
            UpdateActiveDocumentDisplay();
        }

        private void UpdateActiveDocumentDisplay()
        {
            var activeDoc = _solidWorksService.GetActiveDocument();
            var docTitle = activeDoc?.Title ?? "No Document Open";
            this.Text = $"ExportDXF - {docTitle}";

            if (activeDoc == null)
                return;

            // Try each extractor to auto-fill the template
            foreach (var extractor in _extractors)
            {
                if (extractor.TryExtract(activeDoc.Title, out var info))
                {
                    txtFilenameTemplate.Text = info.DefaultTemplate;
                    return;
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
