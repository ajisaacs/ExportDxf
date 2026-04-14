namespace ExportDXF.Forms
{
    partial class MainForm
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
            runButton = new System.Windows.Forms.Button();
            label3 = new System.Windows.Forms.Label();
            viewFlipDeciderBox = new System.Windows.Forms.ComboBox();
            mainTabControl = new System.Windows.Forms.TabControl();
            logEventsTab = new System.Windows.Forms.TabPage();
            logEventsDataGrid = new System.Windows.Forms.DataGridView();
            bomTab = new System.Windows.Forms.TabPage();
            bomDataGrid = new System.Windows.Forms.DataGridView();
            cutTemplatesTab = new System.Windows.Forms.TabPage();
            cutTemplatesDataGrid = new System.Windows.Forms.DataGridView();
            templateLabel = new System.Windows.Forms.Label();
            txtFilenameTemplate = new System.Windows.Forms.TextBox();
            mainTabControl.SuspendLayout();
            logEventsTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)logEventsDataGrid).BeginInit();
            bomTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)bomDataGrid).BeginInit();
            cutTemplatesTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)cutTemplatesDataGrid).BeginInit();
            SuspendLayout();
            //
            // runButton
            //
            runButton.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            runButton.Location = new System.Drawing.Point(508, 12);
            runButton.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            runButton.Name = "runButton";
            runButton.Size = new System.Drawing.Size(65, 57);
            runButton.TabIndex = 11;
            runButton.Text = "Start";
            runButton.UseVisualStyleBackColor = true;
            runButton.Click += button1_Click;
            //
            // templateLabel
            //
            templateLabel.AutoSize = true;
            templateLabel.Location = new System.Drawing.Point(15, 15);
            templateLabel.Name = "templateLabel";
            templateLabel.Size = new System.Drawing.Size(116, 17);
            templateLabel.TabIndex = 2;
            templateLabel.Text = "Filename Template";
            //
            // txtFilenameTemplate
            //
            txtFilenameTemplate.Location = new System.Drawing.Point(137, 12);
            txtFilenameTemplate.Name = "txtFilenameTemplate";
            txtFilenameTemplate.Size = new System.Drawing.Size(365, 25);
            txtFilenameTemplate.TabIndex = 1;
            //
            // label3
            //
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point(26, 46);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(105, 17);
            label3.TabIndex = 2;
            label3.Text = "View flip decider";
            //
            // viewFlipDeciderBox
            //
            viewFlipDeciderBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            viewFlipDeciderBox.FormattingEnabled = true;
            viewFlipDeciderBox.Location = new System.Drawing.Point(137, 43);
            viewFlipDeciderBox.Name = "viewFlipDeciderBox";
            viewFlipDeciderBox.Size = new System.Drawing.Size(365, 25);
            viewFlipDeciderBox.TabIndex = 3;
            //
            // mainTabControl
            //
            mainTabControl.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            mainTabControl.Controls.Add(logEventsTab);
            mainTabControl.Controls.Add(bomTab);
            mainTabControl.Controls.Add(cutTemplatesTab);
            mainTabControl.Location = new System.Drawing.Point(15, 75);
            mainTabControl.Name = "mainTabControl";
            mainTabControl.Padding = new System.Drawing.Point(20, 5);
            mainTabControl.SelectedIndex = 0;
            mainTabControl.Size = new System.Drawing.Size(910, 522);
            mainTabControl.TabIndex = 12;
            //
            // logEventsTab
            //
            logEventsTab.Controls.Add(logEventsDataGrid);
            logEventsTab.Location = new System.Drawing.Point(4, 30);
            logEventsTab.Name = "logEventsTab";
            logEventsTab.Padding = new System.Windows.Forms.Padding(3);
            logEventsTab.Size = new System.Drawing.Size(902, 488);
            logEventsTab.TabIndex = 0;
            logEventsTab.Text = "Log Events";
            logEventsTab.UseVisualStyleBackColor = true;
            //
            // logEventsDataGrid
            //
            logEventsDataGrid.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            logEventsDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            logEventsDataGrid.GridColor = System.Drawing.Color.WhiteSmoke;
            logEventsDataGrid.Location = new System.Drawing.Point(6, 6);
            logEventsDataGrid.Name = "logEventsDataGrid";
            logEventsDataGrid.Size = new System.Drawing.Size(890, 476);
            logEventsDataGrid.TabIndex = 0;
            //
            // bomTab
            //
            bomTab.Controls.Add(bomDataGrid);
            bomTab.Location = new System.Drawing.Point(4, 28);
            bomTab.Name = "bomTab";
            bomTab.Padding = new System.Windows.Forms.Padding(3);
            bomTab.Size = new System.Drawing.Size(902, 490);
            bomTab.TabIndex = 1;
            bomTab.Text = "Bill Of Materials";
            bomTab.UseVisualStyleBackColor = true;
            //
            // bomDataGrid
            //
            bomDataGrid.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            bomDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            bomDataGrid.GridColor = System.Drawing.Color.WhiteSmoke;
            bomDataGrid.Location = new System.Drawing.Point(6, 6);
            bomDataGrid.Name = "bomDataGrid";
            bomDataGrid.Size = new System.Drawing.Size(890, 478);
            bomDataGrid.TabIndex = 1;
            //
            // cutTemplatesTab
            //
            cutTemplatesTab.Controls.Add(cutTemplatesDataGrid);
            cutTemplatesTab.Location = new System.Drawing.Point(4, 28);
            cutTemplatesTab.Name = "cutTemplatesTab";
            cutTemplatesTab.Padding = new System.Windows.Forms.Padding(3);
            cutTemplatesTab.Size = new System.Drawing.Size(902, 490);
            cutTemplatesTab.TabIndex = 2;
            cutTemplatesTab.Text = "Cut Templates";
            cutTemplatesTab.UseVisualStyleBackColor = true;
            //
            // cutTemplatesDataGrid
            //
            cutTemplatesDataGrid.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            cutTemplatesDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            cutTemplatesDataGrid.GridColor = System.Drawing.Color.WhiteSmoke;
            cutTemplatesDataGrid.Location = new System.Drawing.Point(6, 6);
            cutTemplatesDataGrid.Name = "cutTemplatesDataGrid";
            cutTemplatesDataGrid.Size = new System.Drawing.Size(890, 478);
            cutTemplatesDataGrid.TabIndex = 2;
            //
            // MainForm
            //
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            ClientSize = new System.Drawing.Size(937, 609);
            Controls.Add(txtFilenameTemplate);
            Controls.Add(mainTabControl);
            Controls.Add(viewFlipDeciderBox);
            Controls.Add(templateLabel);
            Controls.Add(label3);
            Controls.Add(runButton);
            Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            MaximizeBox = false;
            MinimumSize = new System.Drawing.Size(642, 455);
            Name = "MainForm";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "ExportDXF";
            mainTabControl.ResumeLayout(false);
            logEventsTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)logEventsDataGrid).EndInit();
            bomTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)bomDataGrid).EndInit();
            cutTemplatesTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)cutTemplatesDataGrid).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox viewFlipDeciderBox;
        private System.Windows.Forms.TabControl mainTabControl;
        private System.Windows.Forms.TabPage logEventsTab;
        private System.Windows.Forms.TabPage bomTab;
        private System.Windows.Forms.DataGridView logEventsDataGrid;
        private System.Windows.Forms.DataGridView bomDataGrid;
        private System.Windows.Forms.TabPage cutTemplatesTab;
        private System.Windows.Forms.DataGridView cutTemplatesDataGrid;
        private System.Windows.Forms.Label templateLabel;
        private System.Windows.Forms.TextBox txtFilenameTemplate;
    }
}
