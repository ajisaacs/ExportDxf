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
            equipmentBox = new System.Windows.Forms.ComboBox();
            label1 = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            drawingNoBox = new System.Windows.Forms.ComboBox();
            mainTabControl.SuspendLayout();
            logEventsTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)logEventsDataGrid).BeginInit();
            bomTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)bomDataGrid).BeginInit();
            SuspendLayout();
            // 
            // runButton
            // 
            runButton.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            runButton.Location = new System.Drawing.Point(656, 13);
            runButton.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            runButton.Name = "runButton";
            runButton.Size = new System.Drawing.Size(100, 30);
            runButton.TabIndex = 11;
            runButton.Text = "Start";
            runButton.UseVisualStyleBackColor = true;
            runButton.Click += button1_Click;
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
            viewFlipDeciderBox.Size = new System.Drawing.Size(502, 25);
            viewFlipDeciderBox.TabIndex = 3;
            // 
            // mainTabControl
            // 
            mainTabControl.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            mainTabControl.Controls.Add(logEventsTab);
            mainTabControl.Controls.Add(bomTab);
            mainTabControl.Location = new System.Drawing.Point(15, 74);
            mainTabControl.Name = "mainTabControl";
            mainTabControl.Padding = new System.Drawing.Point(20, 5);
            mainTabControl.SelectedIndex = 0;
            mainTabControl.Size = new System.Drawing.Size(741, 586);
            mainTabControl.TabIndex = 12;
            // 
            // logEventsTab
            // 
            logEventsTab.Controls.Add(logEventsDataGrid);
            logEventsTab.Location = new System.Drawing.Point(4, 30);
            logEventsTab.Name = "logEventsTab";
            logEventsTab.Padding = new System.Windows.Forms.Padding(3);
            logEventsTab.Size = new System.Drawing.Size(733, 552);
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
            logEventsDataGrid.Size = new System.Drawing.Size(721, 540);
            logEventsDataGrid.TabIndex = 0;
            // 
            // bomTab
            // 
            bomTab.Controls.Add(bomDataGrid);
            bomTab.Location = new System.Drawing.Point(4, 30);
            bomTab.Name = "bomTab";
            bomTab.Padding = new System.Windows.Forms.Padding(3);
            bomTab.Size = new System.Drawing.Size(733, 552);
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
            bomDataGrid.Size = new System.Drawing.Size(721, 540);
            bomDataGrid.TabIndex = 1;
            // 
            // equipmentBox
            // 
            equipmentBox.FormattingEnabled = true;
            equipmentBox.Location = new System.Drawing.Point(137, 12);
            equipmentBox.Name = "equipmentBox";
            equipmentBox.Size = new System.Drawing.Size(166, 25);
            equipmentBox.TabIndex = 13;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(61, 15);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(70, 17);
            label1.TabIndex = 2;
            label1.Text = "Equipment";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(354, 15);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(56, 17);
            label2.TabIndex = 2;
            label2.Text = "Drawing";
            // 
            // drawingNoBox
            // 
            drawingNoBox.FormattingEnabled = true;
            drawingNoBox.Location = new System.Drawing.Point(416, 12);
            drawingNoBox.Name = "drawingNoBox";
            drawingNoBox.Size = new System.Drawing.Size(223, 25);
            drawingNoBox.TabIndex = 13;
            // 
            // MainForm
            // 
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            ClientSize = new System.Drawing.Size(768, 672);
            Controls.Add(drawingNoBox);
            Controls.Add(equipmentBox);
            Controls.Add(mainTabControl);
            Controls.Add(viewFlipDeciderBox);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(label3);
            Controls.Add(runButton);
            Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            MaximizeBox = false;
            MinimumSize = new System.Drawing.Size(643, 355);
            Name = "MainForm";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "ExportDXF";
            mainTabControl.ResumeLayout(false);
            logEventsTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)logEventsDataGrid).EndInit();
            bomTab.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)bomDataGrid).EndInit();
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
        private System.Windows.Forms.ComboBox equipmentBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox drawingNoBox;
    }
}
