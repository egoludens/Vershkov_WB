namespace VershkovWB
{
    partial class WBReportProcessorForm
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
            this.sourceFileNameLabel = new System.Windows.Forms.Label();
            this.sourceFileNameText = new System.Windows.Forms.TextBox();
            this.sourceFileNameButton = new System.Windows.Forms.Button();
            this.openSourceFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.processWBReportButton = new System.Windows.Forms.Button();
            this.progressReportText = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // sourceFileNameLabel
            // 
            this.sourceFileNameLabel.Location = new System.Drawing.Point(10, 11);
            this.sourceFileNameLabel.Name = "sourceFileNameLabel";
            this.sourceFileNameLabel.Size = new System.Drawing.Size(119, 21);
            this.sourceFileNameLabel.TabIndex = 0;
            this.sourceFileNameLabel.Text = "Файл с отчетом WB:";
            this.sourceFileNameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // sourceFileNameText
            // 
            this.sourceFileNameText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.sourceFileNameText.Location = new System.Drawing.Point(125, 12);
            this.sourceFileNameText.Name = "sourceFileNameText";
            this.sourceFileNameText.Size = new System.Drawing.Size(345, 20);
            this.sourceFileNameText.TabIndex = 1;
            // 
            // sourceFileNameButton
            // 
            this.sourceFileNameButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.sourceFileNameButton.Location = new System.Drawing.Point(468, 11);
            this.sourceFileNameButton.Name = "sourceFileNameButton";
            this.sourceFileNameButton.Size = new System.Drawing.Size(25, 22);
            this.sourceFileNameButton.TabIndex = 2;
            this.sourceFileNameButton.Text = "...";
            this.sourceFileNameButton.UseVisualStyleBackColor = true;
            this.sourceFileNameButton.Click += new System.EventHandler(this.sourceFileNameButton_Click);
            // 
            // openSourceFileDialog
            // 
            this.openSourceFileDialog.Filter = "Файлы Excel (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы|*.*";
            this.openSourceFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.openSourceFileDialog_FileOk);
            // 
            // processWBReportButton
            // 
            this.processWBReportButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.processWBReportButton.Location = new System.Drawing.Point(10, 37);
            this.processWBReportButton.Name = "processWBReportButton";
            this.processWBReportButton.Size = new System.Drawing.Size(483, 35);
            this.processWBReportButton.TabIndex = 3;
            this.processWBReportButton.Text = "Обработать отчет Wildberries";
            this.processWBReportButton.UseVisualStyleBackColor = true;
            this.processWBReportButton.Click += new System.EventHandler(this.processWBReportButton_Click);
            // 
            // progressReportText
            // 
            this.progressReportText.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressReportText.Location = new System.Drawing.Point(12, 77);
            this.progressReportText.Multiline = true;
            this.progressReportText.Name = "progressReportText";
            this.progressReportText.ReadOnly = true;
            this.progressReportText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.progressReportText.Size = new System.Drawing.Size(480, 352);
            this.progressReportText.TabIndex = 4;
            // 
            // WBReportProcessorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(502, 441);
            this.Controls.Add(this.progressReportText);
            this.Controls.Add(this.processWBReportButton);
            this.Controls.Add(this.sourceFileNameButton);
            this.Controls.Add(this.sourceFileNameText);
            this.Controls.Add(this.sourceFileNameLabel);
            this.Name = "WBReportProcessorForm";
            this.Text = "WB Report Processor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label sourceFileNameLabel;
        private System.Windows.Forms.TextBox sourceFileNameText;
        private System.Windows.Forms.Button sourceFileNameButton;
        private System.Windows.Forms.OpenFileDialog openSourceFileDialog;
        private System.Windows.Forms.Button processWBReportButton;
        private System.Windows.Forms.TextBox progressReportText;
    }
}

