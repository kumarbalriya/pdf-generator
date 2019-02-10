namespace LetterRollout
{
    partial class Form1
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
            this.wordDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnWordFileBrowse = new System.Windows.Forms.Button();
            this.wordLabel = new System.Windows.Forms.Label();
            this.txtWordFileName = new System.Windows.Forms.TextBox();
            this.txtExcelTextbox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExcelFileBrowse = new System.Windows.Forms.Button();
            this.excelDialog = new System.Windows.Forms.OpenFileDialog();
            this.txtOutputFolder = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnOutputFolderBrowse = new System.Windows.Forms.Button();
            this.outputFolderDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.lstLetterGenerated = new System.Windows.Forms.ListBox();
            this.btnProcess = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lstMailSent = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // btnWordFileBrowse
            // 
            this.btnWordFileBrowse.Location = new System.Drawing.Point(641, 73);
            this.btnWordFileBrowse.Name = "btnWordFileBrowse";
            this.btnWordFileBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnWordFileBrowse.TabIndex = 0;
            this.btnWordFileBrowse.Text = "Browse";
            this.btnWordFileBrowse.UseVisualStyleBackColor = true;
            this.btnWordFileBrowse.Click += new System.EventHandler(this.button1_Click);
            // 
            // wordLabel
            // 
            this.wordLabel.AutoSize = true;
            this.wordLabel.Location = new System.Drawing.Point(98, 75);
            this.wordLabel.Name = "wordLabel";
            this.wordLabel.Size = new System.Drawing.Size(131, 17);
            this.wordLabel.TabIndex = 1;
            this.wordLabel.Text = "Word File Template";
            // 
            // txtWordFileName
            // 
            this.txtWordFileName.Location = new System.Drawing.Point(235, 73);
            this.txtWordFileName.Name = "txtWordFileName";
            this.txtWordFileName.ReadOnly = true;
            this.txtWordFileName.Size = new System.Drawing.Size(400, 22);
            this.txtWordFileName.TabIndex = 2;
            // 
            // txtExcelTextbox
            // 
            this.txtExcelTextbox.Location = new System.Drawing.Point(235, 116);
            this.txtExcelTextbox.Name = "txtExcelTextbox";
            this.txtExcelTextbox.ReadOnly = true;
            this.txtExcelTextbox.Size = new System.Drawing.Size(400, 22);
            this.txtExcelTextbox.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(98, 118);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "Excel File";
            // 
            // btnExcelFileBrowse
            // 
            this.btnExcelFileBrowse.Location = new System.Drawing.Point(641, 118);
            this.btnExcelFileBrowse.Name = "btnExcelFileBrowse";
            this.btnExcelFileBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnExcelFileBrowse.TabIndex = 3;
            this.btnExcelFileBrowse.Text = "Browse";
            this.btnExcelFileBrowse.UseVisualStyleBackColor = true;
            this.btnExcelFileBrowse.Click += new System.EventHandler(this.btnExcelFileBrowse_Click);
            // 
            // txtOutputFolder
            // 
            this.txtOutputFolder.Location = new System.Drawing.Point(235, 162);
            this.txtOutputFolder.Name = "txtOutputFolder";
            this.txtOutputFolder.ReadOnly = true;
            this.txtOutputFolder.Size = new System.Drawing.Size(400, 22);
            this.txtOutputFolder.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(98, 164);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 17);
            this.label2.TabIndex = 7;
            this.label2.Text = "Output Folder";
            // 
            // btnOutputFolderBrowse
            // 
            this.btnOutputFolderBrowse.Location = new System.Drawing.Point(641, 164);
            this.btnOutputFolderBrowse.Name = "btnOutputFolderBrowse";
            this.btnOutputFolderBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnOutputFolderBrowse.TabIndex = 6;
            this.btnOutputFolderBrowse.Text = "Browse";
            this.btnOutputFolderBrowse.UseVisualStyleBackColor = true;
            this.btnOutputFolderBrowse.Click += new System.EventHandler(this.btnOutputFolderBrowse_Click);
            // 
            // lstLetterGenerated
            // 
            this.lstLetterGenerated.FormattingEnabled = true;
            this.lstLetterGenerated.ItemHeight = 16;
            this.lstLetterGenerated.Location = new System.Drawing.Point(101, 241);
            this.lstLetterGenerated.Name = "lstLetterGenerated";
            this.lstLetterGenerated.Size = new System.Drawing.Size(284, 324);
            this.lstLetterGenerated.TabIndex = 9;
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(764, 72);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(184, 115);
            this.btnProcess.TabIndex = 10;
            this.btnProcess.Text = "Process";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(185, 221);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(117, 17);
            this.label3.TabIndex = 11;
            this.label3.Text = "Letter Generated";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(516, 221);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 17);
            this.label4.TabIndex = 13;
            this.label4.Text = "EMail Sent";
            // 
            // lstMailSent
            // 
            this.lstMailSent.FormattingEnabled = true;
            this.lstMailSent.ItemHeight = 16;
            this.lstMailSent.Location = new System.Drawing.Point(432, 241);
            this.lstMailSent.Name = "lstMailSent";
            this.lstMailSent.Size = new System.Drawing.Size(284, 324);
            this.lstMailSent.TabIndex = 12;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1167, 613);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lstMailSent);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.lstLetterGenerated);
            this.Controls.Add(this.txtOutputFolder);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnOutputFolderBrowse);
            this.Controls.Add(this.txtExcelTextbox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExcelFileBrowse);
            this.Controls.Add(this.txtWordFileName);
            this.Controls.Add(this.wordLabel);
            this.Controls.Add(this.btnWordFileBrowse);
            this.Name = "Form1";
            this.Text = "Form1";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog wordDialog;
        private System.Windows.Forms.Button btnWordFileBrowse;
        private System.Windows.Forms.Label wordLabel;
        private System.Windows.Forms.TextBox txtWordFileName;
        private System.Windows.Forms.TextBox txtExcelTextbox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnExcelFileBrowse;
        private System.Windows.Forms.OpenFileDialog excelDialog;
        private System.Windows.Forms.TextBox txtOutputFolder;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnOutputFolderBrowse;
        private System.Windows.Forms.FolderBrowserDialog outputFolderDialog;
        private System.Windows.Forms.ListBox lstLetterGenerated;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox lstMailSent;
    }
}

