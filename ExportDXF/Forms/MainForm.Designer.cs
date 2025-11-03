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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.runButton = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.viewFlipDeciderBox = new System.Windows.Forms.ComboBox();
            this.mainTabControl = new System.Windows.Forms.TabControl();
            this.logEventsTab = new System.Windows.Forms.TabPage();
            this.logEventsDataGrid = new System.Windows.Forms.DataGridView();
            this.bomTab = new System.Windows.Forms.TabPage();
            this.bomDataGrid = new System.Windows.Forms.DataGridView();
            this.cutTemplatesTab = new System.Windows.Forms.TabPage();
            this.cutTemplatesDataGrid = new System.Windows.Forms.DataGridView();
            this.dwgDetailsTab = new System.Windows.Forms.TabPage();
            this.drawingPdfViewer = new AxAcroPDFLib.AxAcroPDF();
            this.mainTabControl.SuspendLayout();
            this.logEventsTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.logEventsDataGrid)).BeginInit();
            this.bomTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bomDataGrid)).BeginInit();
            this.cutTemplatesTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.cutTemplatesDataGrid)).BeginInit();
            this.dwgDetailsTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.drawingPdfViewer)).BeginInit();
            this.SuspendLayout();
            // 
            // runButton
            // 
            this.runButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.runButton.Location = new System.Drawing.Point(790, 13);
            this.runButton.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(100, 30);
            this.runButton.TabIndex = 11;
            this.runButton.Text = "Start";
            this.runButton.UseVisualStyleBackColor = true;
            this.runButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 20);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 17);
            this.label3.TabIndex = 2;
            this.label3.Text = "View flip decider :";
            // 
            // viewFlipDeciderBox
            // 
            this.viewFlipDeciderBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.viewFlipDeciderBox.FormattingEnabled = true;
            this.viewFlipDeciderBox.Location = new System.Drawing.Point(130, 17);
            this.viewFlipDeciderBox.Name = "viewFlipDeciderBox";
            this.viewFlipDeciderBox.Size = new System.Drawing.Size(375, 25);
            this.viewFlipDeciderBox.TabIndex = 3;
            // 
            // mainTabControl
            // 
            this.mainTabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.mainTabControl.Controls.Add(this.logEventsTab);
            this.mainTabControl.Controls.Add(this.bomTab);
            this.mainTabControl.Controls.Add(this.cutTemplatesTab);
            this.mainTabControl.Controls.Add(this.dwgDetailsTab);
            this.mainTabControl.Location = new System.Drawing.Point(15, 50);
            this.mainTabControl.Name = "mainTabControl";
            this.mainTabControl.Padding = new System.Drawing.Point(20, 5);
            this.mainTabControl.SelectedIndex = 0;
            this.mainTabControl.Size = new System.Drawing.Size(879, 594);
            this.mainTabControl.TabIndex = 12;
            // 
            // logEventsTab
            // 
            this.logEventsTab.Controls.Add(this.logEventsDataGrid);
            this.logEventsTab.Location = new System.Drawing.Point(4, 30);
            this.logEventsTab.Name = "logEventsTab";
            this.logEventsTab.Padding = new System.Windows.Forms.Padding(3);
            this.logEventsTab.Size = new System.Drawing.Size(690, 484);
            this.logEventsTab.TabIndex = 0;
            this.logEventsTab.Text = "Log Events";
            this.logEventsTab.UseVisualStyleBackColor = true;
            // 
            // logEventsDataGrid
            // 
            this.logEventsDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.logEventsDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.logEventsDataGrid.Location = new System.Drawing.Point(6, 6);
            this.logEventsDataGrid.Name = "logEventsDataGrid";
            this.logEventsDataGrid.Size = new System.Drawing.Size(1038, 536);
            this.logEventsDataGrid.TabIndex = 0;
            // 
            // bomTab
            // 
            this.bomTab.Controls.Add(this.bomDataGrid);
            this.bomTab.Location = new System.Drawing.Point(4, 30);
            this.bomTab.Name = "bomTab";
            this.bomTab.Padding = new System.Windows.Forms.Padding(3);
            this.bomTab.Size = new System.Drawing.Size(690, 484);
            this.bomTab.TabIndex = 1;
            this.bomTab.Text = "Bill Of Materials";
            this.bomTab.UseVisualStyleBackColor = true;
            // 
            // bomDataGrid
            // 
            this.bomDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.bomDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.bomDataGrid.Location = new System.Drawing.Point(6, 6);
            this.bomDataGrid.Name = "bomDataGrid";
            this.bomDataGrid.Size = new System.Drawing.Size(1038, 536);
            this.bomDataGrid.TabIndex = 1;
            // 
            // cutTemplatesTab
            // 
            this.cutTemplatesTab.Controls.Add(this.cutTemplatesDataGrid);
            this.cutTemplatesTab.Location = new System.Drawing.Point(4, 30);
            this.cutTemplatesTab.Name = "cutTemplatesTab";
            this.cutTemplatesTab.Padding = new System.Windows.Forms.Padding(3);
            this.cutTemplatesTab.Size = new System.Drawing.Size(690, 484);
            this.cutTemplatesTab.TabIndex = 3;
            this.cutTemplatesTab.Text = "Cut Templates";
            this.cutTemplatesTab.UseVisualStyleBackColor = true;
            // 
            // cutTemplatesDataGrid
            // 
            this.cutTemplatesDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cutTemplatesDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.cutTemplatesDataGrid.Location = new System.Drawing.Point(6, 6);
            this.cutTemplatesDataGrid.Name = "cutTemplatesDataGrid";
            this.cutTemplatesDataGrid.Size = new System.Drawing.Size(678, 468);
            this.cutTemplatesDataGrid.TabIndex = 2;
            // 
            // dwgDetailsTab
            // 
            this.dwgDetailsTab.Controls.Add(this.drawingPdfViewer);
            this.dwgDetailsTab.Location = new System.Drawing.Point(4, 30);
            this.dwgDetailsTab.Name = "dwgDetailsTab";
            this.dwgDetailsTab.Padding = new System.Windows.Forms.Padding(3);
            this.dwgDetailsTab.Size = new System.Drawing.Size(871, 560);
            this.dwgDetailsTab.TabIndex = 2;
            this.dwgDetailsTab.Text = "Drawing Details";
            this.dwgDetailsTab.UseVisualStyleBackColor = true;
            // 
            // drawingPdfViewer
            // 
            this.drawingPdfViewer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.drawingPdfViewer.Enabled = true;
            this.drawingPdfViewer.Location = new System.Drawing.Point(6, 6);
            this.drawingPdfViewer.Name = "drawingPdfViewer";
            this.drawingPdfViewer.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("drawingPdfViewer.OcxState")));
            this.drawingPdfViewer.Size = new System.Drawing.Size(859, 548);
            this.drawingPdfViewer.TabIndex = 0;
            // 
            // MainForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(906, 656);
            this.Controls.Add(this.mainTabControl);
            this.Controls.Add(this.viewFlipDeciderBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.runButton);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(643, 355);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExportDXF";
            this.mainTabControl.ResumeLayout(false);
            this.logEventsTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.logEventsDataGrid)).EndInit();
            this.bomTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bomDataGrid)).EndInit();
            this.cutTemplatesTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.cutTemplatesDataGrid)).EndInit();
            this.dwgDetailsTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.drawingPdfViewer)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox viewFlipDeciderBox;
        private System.Windows.Forms.TabControl mainTabControl;
        private System.Windows.Forms.TabPage logEventsTab;
        private System.Windows.Forms.TabPage bomTab;
        private System.Windows.Forms.TabPage dwgDetailsTab;
        private System.Windows.Forms.DataGridView logEventsDataGrid;
        private AxAcroPDFLib.AxAcroPDF drawingPdfViewer;
        private System.Windows.Forms.DataGridView bomDataGrid;
        private System.Windows.Forms.TabPage cutTemplatesTab;
        private System.Windows.Forms.DataGridView cutTemplatesDataGrid;
    }
}

