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
            components = new System.ComponentModel.Container();
            fileBrowseButton = new System.Windows.Forms.Button();
            fileListBox = new System.Windows.Forms.ListBox();
            processButton = new System.Windows.Forms.Button();
            progressBar1 = new System.Windows.Forms.ProgressBar();
            directoryTextBox = new System.Windows.Forms.TextBox();
            label1 = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            directoryBrowseButton = new System.Windows.Forms.Button();
            label3 = new System.Windows.Forms.Label();
            label4 = new System.Windows.Forms.Label();
            backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            timer1 = new System.Windows.Forms.Timer(components);
            label5 = new System.Windows.Forms.Label();
            cancelButton = new System.Windows.Forms.Button();
            label6 = new System.Windows.Forms.Label();
            SuspendLayout();
            // 
            // fileBrowseButton
            // 
            fileBrowseButton.Anchor = System.Windows.Forms.AnchorStyles.Left;
            fileBrowseButton.Location = new System.Drawing.Point(14, 435);
            fileBrowseButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            fileBrowseButton.Name = "fileBrowseButton";
            fileBrowseButton.Size = new System.Drawing.Size(88, 27);
            fileBrowseButton.TabIndex = 0;
            fileBrowseButton.Text = "Select Files";
            fileBrowseButton.UseVisualStyleBackColor = true;
            fileBrowseButton.Click += button1_Click;
            // 
            // fileListBox
            // 
            fileListBox.Anchor = ((System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right));
            fileListBox.FormattingEnabled = true;
            fileListBox.HorizontalScrollbar = true;
            fileListBox.Location = new System.Drawing.Point(13, 80);
            fileListBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            fileListBox.Name = "fileListBox";
            fileListBox.Size = new System.Drawing.Size(333, 349);
            fileListBox.TabIndex = 1;
            // 
            // processButton
            // 
            processButton.Anchor = ((System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right));
            processButton.BackColor = System.Drawing.SystemColors.Control;
            processButton.Cursor = System.Windows.Forms.Cursors.Hand;
            processButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            processButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            processButton.Location = new System.Drawing.Point(13, 567);
            processButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            processButton.Name = "processButton";
            processButton.Size = new System.Drawing.Size(334, 46);
            processButton.TabIndex = 2;
            processButton.Text = "Process";
            processButton.UseVisualStyleBackColor = false;
            processButton.Click += button1_Click_1;
            // 
            // progressBar1
            // 
            progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right));
            progressBar1.Location = new System.Drawing.Point(14, 297);
            progressBar1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new System.Drawing.Size(334, 28);
            progressBar1.TabIndex = 4;
            // 
            // directoryTextBox
            // 
            directoryTextBox.Anchor = ((System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right));
            directoryTextBox.Location = new System.Drawing.Point(13, 505);
            directoryTextBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            directoryTextBox.Name = "directoryTextBox";
            directoryTextBox.Size = new System.Drawing.Size(333, 23);
            directoryTextBox.TabIndex = 5;
            // 
            // label1
            // 
            label1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            label1.Location = new System.Drawing.Point(13, 50);
            label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(117, 27);
            label1.TabIndex = 6;
            label1.Text = "Files:";
            // 
            // label2
            // 
            label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            label2.Location = new System.Drawing.Point(13, 475);
            label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(117, 27);
            label2.TabIndex = 7;
            label2.Text = "Output Directory:";
            // 
            // directoryBrowseButton
            // 
            directoryBrowseButton.Anchor = System.Windows.Forms.AnchorStyles.Left;
            directoryBrowseButton.Location = new System.Drawing.Point(14, 534);
            directoryBrowseButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
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
            label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(333, 42);
            label3.TabIndex = 9;
            label3.Text = "Excel Table Merger";
            // 
            // label4
            // 
            label4.Anchor = ((System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right));
            label4.Font = new System.Drawing.Font("Segoe UI", 7F);
            label4.Location = new System.Drawing.Point(240, 626);
            label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label4.Name = "label4";
            label4.Size = new System.Drawing.Size(147, 18);
            label4.TabIndex = 10;
            label4.Text = $"{Application.CompanyName}\r\n";
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
            cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right));
            cancelButton.BackColor = System.Drawing.SystemColors.Control;
            cancelButton.Cursor = System.Windows.Forms.Cursors.Hand;
            cancelButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            cancelButton.Location = new System.Drawing.Point(13, 567);
            cancelButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(334, 46);
            cancelButton.TabIndex = 12;
            cancelButton.Text = "Cancel";
            cancelButton.UseVisualStyleBackColor = false;
            cancelButton.Click += cancelButton_Click;
            // 
            // label6
            // 
            label6.Anchor = ((System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left));
            label6.Font = new System.Drawing.Font("Segoe UI", 7F);
            label6.Location = new System.Drawing.Point(1, 626);
            label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            label6.Name = "label6";
            label6.Size = new System.Drawing.Size(147, 18);
            label6.TabIndex = 13;
            label6.Text = $"Version: {Application.ProductVersion}";

            // 
            // MainWindow
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackColor = System.Drawing.SystemColors.Control;
            ClientSize = new System.Drawing.Size(364, 641);
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
            Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            MinimumSize = new System.Drawing.Size(380, 680);
            ResumeLayout(false);
            PerformLayout();
        }

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
    }
}