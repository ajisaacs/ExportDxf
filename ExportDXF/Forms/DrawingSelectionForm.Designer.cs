namespace ExportDXF.Forms
{
    partial class DrawingSelectionForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.equipmentLabel = new System.Windows.Forms.Label();
            this.equipmentComboBox = new System.Windows.Forms.ComboBox();
            this.drawingLabel = new System.Windows.Forms.Label();
            this.drawingComboBox = new System.Windows.Forms.ComboBox();
            this.newDrawingButton = new System.Windows.Forms.Button();
            this.newDrawingPanel = new System.Windows.Forms.Panel();
            this.qtyNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.qtyLabel = new System.Windows.Forms.Label();
            this.descriptionTextBox = new System.Windows.Forms.TextBox();
            this.descriptionLabel = new System.Windows.Forms.Label();
            this.drawingNumberTextBox = new System.Windows.Forms.TextBox();
            this.drawingNumberLabel = new System.Windows.Forms.Label();
            this.currentDrawingLabel = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.statusLabel = new System.Windows.Forms.Label();
            this.newDrawingPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.qtyNumericUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // equipmentLabel
            // 
            this.equipmentLabel.AutoSize = true;
            this.equipmentLabel.Location = new System.Drawing.Point(12, 15);
            this.equipmentLabel.Name = "equipmentLabel";
            this.equipmentLabel.Size = new System.Drawing.Size(82, 17);
            this.equipmentLabel.TabIndex = 0;
            this.equipmentLabel.Text = "Equipment #";
            // 
            // equipmentComboBox
            // 
            this.equipmentComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.equipmentComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.equipmentComboBox.FormattingEnabled = true;
            this.equipmentComboBox.Location = new System.Drawing.Point(100, 12);
            this.equipmentComboBox.Name = "equipmentComboBox";
            this.equipmentComboBox.Size = new System.Drawing.Size(470, 25);
            this.equipmentComboBox.TabIndex = 1;
            this.equipmentComboBox.SelectedIndexChanged += new System.EventHandler(this.equipmentComboBox_SelectedIndexChanged);
            // 
            // drawingLabel
            // 
            this.drawingLabel.AutoSize = true;
            this.drawingLabel.Location = new System.Drawing.Point(12, 48);
            this.drawingLabel.Name = "drawingLabel";
            this.drawingLabel.Size = new System.Drawing.Size(68, 17);
            this.drawingLabel.TabIndex = 2;
            this.drawingLabel.Text = "Drawing #";
            // 
            // drawingComboBox
            // 
            this.drawingComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.drawingComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.drawingComboBox.FormattingEnabled = true;
            this.drawingComboBox.Location = new System.Drawing.Point(100, 45);
            this.drawingComboBox.Name = "drawingComboBox";
            this.drawingComboBox.Size = new System.Drawing.Size(370, 25);
            this.drawingComboBox.TabIndex = 3;
            // 
            // newDrawingButton
            // 
            this.newDrawingButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.newDrawingButton.Location = new System.Drawing.Point(476, 43);
            this.newDrawingButton.Name = "newDrawingButton";
            this.newDrawingButton.Size = new System.Drawing.Size(94, 27);
            this.newDrawingButton.TabIndex = 4;
            this.newDrawingButton.Text = "New Drawing";
            this.newDrawingButton.UseVisualStyleBackColor = true;
            this.newDrawingButton.Click += new System.EventHandler(this.newDrawingButton_Click);
            // 
            // newDrawingPanel
            // 
            this.newDrawingPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.newDrawingPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.newDrawingPanel.Controls.Add(this.qtyNumericUpDown);
            this.newDrawingPanel.Controls.Add(this.qtyLabel);
            this.newDrawingPanel.Controls.Add(this.descriptionTextBox);
            this.newDrawingPanel.Controls.Add(this.descriptionLabel);
            this.newDrawingPanel.Controls.Add(this.drawingNumberTextBox);
            this.newDrawingPanel.Controls.Add(this.drawingNumberLabel);
            this.newDrawingPanel.Location = new System.Drawing.Point(15, 85);
            this.newDrawingPanel.Name = "newDrawingPanel";
            this.newDrawingPanel.Size = new System.Drawing.Size(555, 120);
            this.newDrawingPanel.TabIndex = 5;
            this.newDrawingPanel.Visible = false;
            // 
            // qtyNumericUpDown
            // 
            this.qtyNumericUpDown.Location = new System.Drawing.Point(124, 72);
            this.qtyNumericUpDown.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.qtyNumericUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.qtyNumericUpDown.Name = "qtyNumericUpDown";
            this.qtyNumericUpDown.Size = new System.Drawing.Size(100, 25);
            this.qtyNumericUpDown.TabIndex = 5;
            this.qtyNumericUpDown.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // qtyLabel
            // 
            this.qtyLabel.AutoSize = true;
            this.qtyLabel.Location = new System.Drawing.Point(10, 74);
            this.qtyLabel.Name = "qtyLabel";
            this.qtyLabel.Size = new System.Drawing.Size(56, 17);
            this.qtyLabel.TabIndex = 4;
            this.qtyLabel.Text = "Quantity";
            // 
            // descriptionTextBox
            // 
            this.descriptionTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.descriptionTextBox.Location = new System.Drawing.Point(124, 41);
            this.descriptionTextBox.Name = "descriptionTextBox";
            this.descriptionTextBox.Size = new System.Drawing.Size(411, 25);
            this.descriptionTextBox.TabIndex = 3;
            // 
            // descriptionLabel
            // 
            this.descriptionLabel.AutoSize = true;
            this.descriptionLabel.Location = new System.Drawing.Point(10, 44);
            this.descriptionLabel.Name = "descriptionLabel";
            this.descriptionLabel.Size = new System.Drawing.Size(74, 17);
            this.descriptionLabel.TabIndex = 2;
            this.descriptionLabel.Text = "Description";
            // 
            // drawingNumberTextBox
            // 
            this.drawingNumberTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.drawingNumberTextBox.Location = new System.Drawing.Point(124, 10);
            this.drawingNumberTextBox.Name = "drawingNumberTextBox";
            this.drawingNumberTextBox.Size = new System.Drawing.Size(411, 25);
            this.drawingNumberTextBox.TabIndex = 1;
            this.drawingNumberTextBox.TextChanged += new System.EventHandler(this.drawingNumberTextBox_TextChanged);
            // 
            // drawingNumberLabel
            // 
            this.drawingNumberLabel.AutoSize = true;
            this.drawingNumberLabel.Location = new System.Drawing.Point(10, 13);
            this.drawingNumberLabel.Name = "drawingNumberLabel";
            this.drawingNumberLabel.Size = new System.Drawing.Size(108, 17);
            this.drawingNumberLabel.TabIndex = 0;
            this.drawingNumberLabel.Text = "Drawing Number";
            // 
            // currentDrawingLabel
            // 
            this.currentDrawingLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.currentDrawingLabel.AutoEllipsis = true;
            this.currentDrawingLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.currentDrawingLabel.Location = new System.Drawing.Point(15, 217);
            this.currentDrawingLabel.Name = "currentDrawingLabel";
            this.currentDrawingLabel.Size = new System.Drawing.Size(557, 41);
            this.currentDrawingLabel.TabIndex = 9;
            this.currentDrawingLabel.Text = "Loading active drawing...";
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.Enabled = false;
            this.okButton.Location = new System.Drawing.Point(414, 273);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 30);
            this.okButton.TabIndex = 6;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(495, 273);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 30);
            this.cancelButton.TabIndex = 7;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(15, 280);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(44, 17);
            this.statusLabel.TabIndex = 8;
            this.statusLabel.Text = "Ready";
            // 
            // DrawingSelectionForm
            // 
            this.AcceptButton = this.okButton;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(584, 315);
            this.Controls.Add(this.currentDrawingLabel);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.newDrawingPanel);
            this.Controls.Add(this.newDrawingButton);
            this.Controls.Add(this.drawingComboBox);
            this.Controls.Add(this.drawingLabel);
            this.Controls.Add(this.equipmentComboBox);
            this.Controls.Add(this.equipmentLabel);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DrawingSelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Select Drawing - ExportDXF";
            this.newDrawingPanel.ResumeLayout(false);
            this.newDrawingPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.qtyNumericUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label equipmentLabel;
        private System.Windows.Forms.ComboBox equipmentComboBox;
        private System.Windows.Forms.Label drawingLabel;
        private System.Windows.Forms.ComboBox drawingComboBox;
        private System.Windows.Forms.Button newDrawingButton;
        private System.Windows.Forms.Panel newDrawingPanel;
        private System.Windows.Forms.TextBox drawingNumberTextBox;
        private System.Windows.Forms.Label drawingNumberLabel;
        private System.Windows.Forms.TextBox descriptionTextBox;
        private System.Windows.Forms.Label descriptionLabel;
        private System.Windows.Forms.NumericUpDown qtyNumericUpDown;
        private System.Windows.Forms.Label qtyLabel;
        private System.Windows.Forms.Label currentDrawingLabel;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label statusLabel;
    }
}
