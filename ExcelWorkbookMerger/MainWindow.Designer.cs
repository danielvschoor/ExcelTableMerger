using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;

namespace ExcelWorkbookMerger
{
    partial class MainWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private List<String> ExcelFilesToProcess = new List<string>();
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
            components = new Container();
            fileBrowseButton = new Button();
            fileListBox = new ListBox();
            processButton = new Button();
            progressBar1 = new ProgressBar();
            directoryTextBox = new TextBox();
            label1 = new Label();
            label2 = new Label();
            directoryBrowseButton = new Button();
            label3 = new Label();
            label4 = new Label();
            backgroundWorker1 = new BackgroundWorker();
            timer1 = new Timer(components);
            label5 = new Label();
            cancelButton = new Button();
            label6 = new Label();
            onlyMergeLatestSheetCheckBox = new CheckBox();
            label7 = new Label();
            panel1 = new Panel();
            enableDebugSheet = new CheckBox();
            toolTip1 = new ToolTip(components);
            clearFileListButton = new Button();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // fileBrowseButton
            // 
            fileBrowseButton.Anchor = AnchorStyles.Left;
            fileBrowseButton.Location = new System.Drawing.Point(14, 439);
            fileBrowseButton.Margin = new Padding(4, 3, 4, 3);
            fileBrowseButton.Name = "fileBrowseButton";
            fileBrowseButton.Size = new System.Drawing.Size(88, 27);
            fileBrowseButton.TabIndex = 0;
            fileBrowseButton.Text = "Select Files";
            fileBrowseButton.UseVisualStyleBackColor = true;
            fileBrowseButton.Click += button1_Click;
            // 
            // fileListBox
            // 
            fileListBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            fileListBox.FormattingEnabled = true;
            fileListBox.HorizontalScrollbar = true;
            fileListBox.Location = new System.Drawing.Point(13, 84);
            fileListBox.Margin = new Padding(4, 3, 4, 3);
            fileListBox.Name = "fileListBox";
            fileListBox.Size = new System.Drawing.Size(450, 349);
            fileListBox.TabIndex = 1;
            // 
            // processButton
            // 
            processButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            processButton.BackColor = System.Drawing.SystemColors.Control;
            processButton.Cursor = Cursors.Hand;
            processButton.FlatStyle = FlatStyle.Popup;
            processButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            processButton.Location = new System.Drawing.Point(13, 574);
            processButton.Margin = new Padding(4, 3, 4, 3);
            processButton.Name = "processButton";
            processButton.Size = new System.Drawing.Size(450, 46);
            processButton.TabIndex = 2;
            processButton.Text = "Process";
            processButton.UseVisualStyleBackColor = false;
            processButton.Click += button1_Click_1;
            // 
            // progressBar1
            // 
            progressBar1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            progressBar1.Location = new System.Drawing.Point(14, 297);
            progressBar1.Margin = new Padding(4, 3, 4, 3);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new System.Drawing.Size(449, 35);
            progressBar1.TabIndex = 4;
            // 
            // directoryTextBox
            // 
            directoryTextBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            directoryTextBox.Location = new System.Drawing.Point(13, 509);
            directoryTextBox.Margin = new Padding(4, 3, 4, 3);
            directoryTextBox.Name = "directoryTextBox";
            directoryTextBox.Size = new System.Drawing.Size(450, 23);
            directoryTextBox.TabIndex = 5;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Left;
            label1.Location = new System.Drawing.Point(13, 54);
            label1.Margin = new Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(117, 27);
            label1.TabIndex = 6;
            label1.Text = "Files:";
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Left;
            label2.Location = new System.Drawing.Point(13, 479);
            label2.Margin = new Padding(4, 0, 4, 0);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(117, 27);
            label2.TabIndex = 7;
            label2.Text = "Output Directory:";
            // 
            // directoryBrowseButton
            // 
            directoryBrowseButton.Anchor = AnchorStyles.Left;
            directoryBrowseButton.Location = new System.Drawing.Point(14, 538);
            directoryBrowseButton.Margin = new Padding(4, 3, 4, 3);
            directoryBrowseButton.Name = "directoryBrowseButton";
            directoryBrowseButton.Size = new System.Drawing.Size(117, 27);
            directoryBrowseButton.TabIndex = 8;
            directoryBrowseButton.Text = "Select Directory";
            directoryBrowseButton.UseVisualStyleBackColor = true;
            directoryBrowseButton.Click += button1_Click_2;
            // 
            // label3
            // 
            label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            label3.Location = new System.Drawing.Point(13, 9);
            label3.Margin = new Padding(4, 0, 4, 0);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(333, 42);
            label3.TabIndex = 9;
            label3.Text = "Excel Table Merger";
            // 
            // label4
            // 
            label4.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            label4.Font = new System.Drawing.Font("Segoe UI", 7F);
            label4.Location = new System.Drawing.Point(590, 633);
            label4.Margin = new Padding(4, 0, 4, 0);
            label4.Name = "label4";
            label4.Size = new System.Drawing.Size(147, 18);
            label4.TabIndex = 10;
            label4.Text = "van Schoor-Els Technology\r\n";
            // 
            // backgroundWorker1
            // 
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker1.RunWorkerCompleted += backgroundWorker1_RunWorkerCompleted;
            // 
            // timer1
            // 
            timer1.Tick += timer1_Tick;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new System.Drawing.Point(14, 447);
            label5.Name = "label5";
            label5.Size = new System.Drawing.Size(0, 15);
            label5.TabIndex = 11;
            // 
            // cancelButton
            // 
            cancelButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            cancelButton.BackColor = System.Drawing.SystemColors.Control;
            cancelButton.Cursor = Cursors.Hand;
            cancelButton.FlatStyle = FlatStyle.Popup;
            cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            cancelButton.Location = new System.Drawing.Point(13, 574);
            cancelButton.Margin = new Padding(4, 3, 4, 3);
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(450, 46);
            cancelButton.TabIndex = 12;
            cancelButton.Text = "Cancel";
            cancelButton.UseVisualStyleBackColor = false;
            cancelButton.Click += cancelButton_Click;
            // 
            // label6
            // 
            label6.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            label6.Font = new System.Drawing.Font("Segoe UI", 7F);
            label6.Location = new System.Drawing.Point(1, 633);
            label6.Margin = new Padding(4, 0, 4, 0);
            label6.Name = "label6";
            label6.Size = new System.Drawing.Size(147, 18);
            label6.TabIndex = 13;
            label6.Text = "Version: " + typeof(Program).Assembly.GetName().Version;
            // 
            // onlyMergeLatestSheetCheckBox
            // 
            onlyMergeLatestSheetCheckBox.Location = new System.Drawing.Point(3, 33);
            onlyMergeLatestSheetCheckBox.Name = "onlyMergeLatestSheetCheckBox";
            onlyMergeLatestSheetCheckBox.Size = new System.Drawing.Size(240, 24);
            onlyMergeLatestSheetCheckBox.TabIndex = 14;
            onlyMergeLatestSheetCheckBox.Text = "Only merge latest sheet";
            toolTip1.SetToolTip(onlyMergeLatestSheetCheckBox, "Only merge the latest sheet in each of the worksheets.\r\n\r\nThe sheet names MUST be dates in the format \"MMMyy\" \r\nor \"ddMMMyy\" (FEB24 or 21FEB24)");
            onlyMergeLatestSheetCheckBox.UseVisualStyleBackColor = true;
            onlyMergeLatestSheetCheckBox.Checked = true;
            // 
            // label7
            // 
            label7.Anchor = AnchorStyles.Right;
            label7.Location = new System.Drawing.Point(470, 54);
            label7.Margin = new Padding(4, 0, 4, 0);
            label7.Name = "label7";
            label7.Size = new System.Drawing.Size(117, 27);
            label7.TabIndex = 15;
            label7.Text = "Options:";
            // 
            // panel1
            // 
            panel1.Anchor = AnchorStyles.Right;
            panel1.BorderStyle = BorderStyle.FixedSingle;
            panel1.Controls.Add(enableDebugSheet);
            panel1.Controls.Add(onlyMergeLatestSheetCheckBox);
            panel1.Location = new System.Drawing.Point(470, 84);
            panel1.Margin = new Padding(4, 3, 4, 3);
            panel1.Name = "panel1";
            panel1.Size = new System.Drawing.Size(244, 349);
            panel1.TabIndex = 16;
            // 
            // checkBox1
            // 
            enableDebugSheet.Location = new System.Drawing.Point(3, 3);
            enableDebugSheet.Name = "checkBox1";
            enableDebugSheet.Size = new System.Drawing.Size(240, 24);
            enableDebugSheet.TabIndex = 15;
            enableDebugSheet.Text = "Enable debug sheet";
            toolTip1.SetToolTip(enableDebugSheet, "Create a new sheet in the merged excel file called \"Debug\". \r\nwhich will contain debug information useful for troubleshooting issues");
            enableDebugSheet.UseVisualStyleBackColor = true;
            // 
            // clearFileListButton
            // 
            clearFileListButton.Anchor = AnchorStyles.Right;
            clearFileListButton.Location = new System.Drawing.Point(375, 439);
            clearFileListButton.Margin = new Padding(4, 3, 4, 3);
            clearFileListButton.Name = "clearFileListButton";
            clearFileListButton.Size = new System.Drawing.Size(88, 27);
            clearFileListButton.TabIndex = 17;
            clearFileListButton.Text = "Clear Selection";
            clearFileListButton.UseVisualStyleBackColor = true;
            clearFileListButton.Click += clearFileListButton_Click;
            // 
            // MainWindow
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = System.Drawing.SystemColors.Control;
            ClientSize = new System.Drawing.Size(714, 648);
            Controls.Add(clearFileListButton);
            Controls.Add(panel1);
            Controls.Add(label7);
            Controls.Add(label6);
            Controls.Add(cancelButton);
            Controls.Add(label5);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(directoryBrowseButton);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(directoryTextBox);
            Controls.Add(progressBar1);
            Controls.Add(processButton);
            Controls.Add(fileListBox);
            Controls.Add(fileBrowseButton);
            Location = new System.Drawing.Point(15, 15);
            Margin = new Padding(4, 3, 4, 3);
            MinimumSize = new System.Drawing.Size(380, 680);
            Name = "MainWindow";
            panel1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.CheckBox onlyMergeLatestSheetCheckBox;
        private System.Windows.Forms.Label label7;

        private System.Windows.Forms.Label label6;


        private System.Windows.Forms.Label label4;

        private System.Windows.Forms.Label label3;

        private System.Windows.Forms.Button directoryBrowseButton;

        private System.Windows.Forms.Label label2;

        private System.Windows.Forms.TextBox directoryTextBox;
        private System.Windows.Forms.Label label1;

        private System.Windows.Forms.ProgressBar progressBar1;

        private System.Windows.Forms.Button processButton;

        private System.Windows.Forms.ListBox fileListBox;

        private System.Windows.Forms.Button fileBrowseButton;

        #endregion

        private System.ComponentModel.BackgroundWorker backgroundWorker1;

        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button cancelButton;
        private Panel panel1;
        private CheckBox enableDebugSheet;
        private ToolTip toolTip1;
        private Button clearFileListButton;
    }
}