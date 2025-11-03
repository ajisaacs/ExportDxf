using ExportDXF.Services;
using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportDXF.Forms
{
    public partial class DrawingSelectionForm : Form
    {
        private readonly ICutFabApiClient _apiClient;

        public int? SelectedDrawingId { get; private set; }
        public string SelectedDrawingNumber { get; private set; }

        public DrawingSelectionForm(ICutFabApiClient apiClient)
        {
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
            InitializeComponent();
        }

        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            await LoadEquipmentAsync();
        }

        private async Task LoadEquipmentAsync()
        {
            try
            {
                statusLabel.Text = "Loading equipment...";
                statusLabel.ForeColor = Color.Black;

                var equipment = await _apiClient.GetEquipmentAsync();

                equipmentComboBox.DisplayMember = nameof(CutFabApiClient.ApiEquipment.EquipmentNumber);
                equipmentComboBox.ValueMember = nameof(CutFabApiClient.ApiEquipment.ID);
                equipmentComboBox.DataSource = equipment;

                if (equipment.Count > 0)
                {
                    equipmentComboBox.SelectedIndex = 0;
                    statusLabel.Text = $"Loaded {equipment.Count} equipment record(s)";
                    statusLabel.ForeColor = Color.Green;
                }
                else
                {
                    statusLabel.Text = "No equipment found";
                    statusLabel.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"Error loading equipment: {ex.Message}";
                statusLabel.ForeColor = Color.Red;
                MessageBox.Show($"Failed to load equipment: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void equipmentComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            await LoadDrawingsAsync();
        }

        private async Task LoadDrawingsAsync()
        {
            try
            {
                var selectedEquipment = equipmentComboBox.SelectedItem as CutFabApiClient.ApiEquipment;
                if (selectedEquipment == null)
                {
                    drawingComboBox.DataSource = null;
                    return;
                }

                statusLabel.Text = "Loading drawings...";
                statusLabel.ForeColor = Color.Black;

                var drawings = await _apiClient.GetDrawingsForEquipmentAsync(selectedEquipment.ID);

                drawingComboBox.DisplayMember = nameof(CutFabApiClient.ApiDrawingSummary.DrawingNumber);
                drawingComboBox.ValueMember = nameof(CutFabApiClient.ApiDrawingSummary.ID);
                drawingComboBox.DataSource = drawings;

                if (drawings.Count > 0)
                {
                    statusLabel.Text = $"Loaded {drawings.Count} drawing(s)";
                    statusLabel.ForeColor = Color.Green;
                    drawingComboBox.Enabled = true;
                }
                else
                {
                    statusLabel.Text = "No drawings found for this equipment";
                    statusLabel.ForeColor = Color.DarkBlue;
                    drawingComboBox.Enabled = false;
                }

                UpdateOkButtonState();
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"Error loading drawings: {ex.Message}";
                statusLabel.ForeColor = Color.Red;
            }
        }

        private void newDrawingButton_Click(object sender, EventArgs e)
        {
            ToggleNewDrawingMode(!newDrawingPanel.Visible);
        }

        private void ToggleNewDrawingMode(bool enabled)
        {
            newDrawingPanel.Visible = enabled;
            drawingComboBox.Enabled = !enabled;
            newDrawingButton.Text = enabled ? "Cancel New" : "New Drawing";

            if (enabled)
            {
                drawingNumberTextBox.Focus();
            }

            UpdateOkButtonState();
        }

        private void drawingNumberTextBox_TextChanged(object sender, EventArgs e)
        {
            UpdateOkButtonState();
        }

        private void UpdateOkButtonState()
        {
            if (newDrawingPanel.Visible)
            {
                // Creating new drawing - require drawing number
                okButton.Enabled = !string.IsNullOrWhiteSpace(drawingNumberTextBox.Text);
            }
            else
            {
                // Selecting existing drawing
                okButton.Enabled = drawingComboBox.SelectedItem != null;
            }
        }

        private async void okButton_Click(object sender, EventArgs e)
        {
            try
            {
                okButton.Enabled = false;
                statusLabel.Text = "Processing...";
                statusLabel.ForeColor = Color.Black;

                if (newDrawingPanel.Visible)
                {
                    // Create new drawing
                    await CreateNewDrawingAsync();
                }
                else
                {
                    // Use existing drawing
                    var selectedDrawing = drawingComboBox.SelectedItem as CutFabApiClient.ApiDrawingSummary;
                    if (selectedDrawing != null)
                    {
                        SelectedDrawingId = selectedDrawing.ID;
                        SelectedDrawingNumber = selectedDrawing.DrawingNumber;
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                }
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"Error: {ex.Message}";
                statusLabel.ForeColor = Color.Red;
                MessageBox.Show($"Operation failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                okButton.Enabled = true;
            }
        }

        private async Task CreateNewDrawingAsync()
        {
            var selectedEquipment = equipmentComboBox.SelectedItem as CutFabApiClient.ApiEquipment;
            if (selectedEquipment == null)
            {
                MessageBox.Show("Please select equipment first.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var drawingNumber = drawingNumberTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(drawingNumber))
            {
                MessageBox.Show("Please enter a drawing number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            statusLabel.Text = "Creating new drawing...";
            statusLabel.ForeColor = Color.Black;

            var response = await _apiClient.CreateDrawingWithInfoAsync(selectedEquipment.ID, drawingNumber);

            if (response.Success && response.Data.HasValue)
            {
                SelectedDrawingId = response.Data.Value;
                SelectedDrawingNumber = drawingNumber;
                statusLabel.Text = "Drawing created successfully";
                statusLabel.ForeColor = Color.Green;
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                var errorMsg = response.Error ?? "Unknown error occurred";
                statusLabel.Text = $"Failed to create drawing: {errorMsg}";
                statusLabel.ForeColor = Color.Red;
                MessageBox.Show($"Failed to create drawing: {errorMsg}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                okButton.Enabled = true;
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
