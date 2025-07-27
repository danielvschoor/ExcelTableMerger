using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelWorkbookMerger.Logic;
using ExcelWorkbookMerger.Models;

namespace ExcelWorkbookMerger;

public partial class MainWindow : Form
{
    private readonly Stopwatch _stopwatch = new();

    public MainWindow()
    {
        InitializeComponent();
        progressBar1.Hide();
        cancelButton.Hide();
        backgroundWorker1.WorkerReportsProgress = true;
        backgroundWorker1.WorkerSupportsCancellation = true;
        fileListBox.SelectionMode = SelectionMode.None;
    }

    private void button1_Click(object sender, EventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            CheckFileExists = true,
            CheckPathExists = true,
            DefaultExt = "csv",
            Filter = "Excel Files|*.xls;*.xlsx",
            Title = "Browse Excel Files",
            Multiselect = true,
            RestoreDirectory = true
        };
        if (dialog.ShowDialog() != DialogResult.OK)
            return;
        var basePath = string.Empty;

        foreach (var file in dialog.FileNames)
        {
            var fi = new FileInfo(file);
            basePath = fi.DirectoryName;
            fileListBox.Items.Add(file);
        }

        if (basePath != string.Empty) directoryTextBox.Text = basePath;
    }

    private void button1_Click_1(object sender, EventArgs e)
    {
        if (backgroundWorker1.IsBusy)
        {
            MessageBox.Show("Already busy processing");
            return;
        }

        progressBar1.Step = 100 / fileListBox.Items.Count;

        fileListBox.Hide();
        processButton.Hide();
        fileBrowseButton.Hide();
        directoryBrowseButton.Hide();
        directoryTextBox.Hide();
        label1.Hide();
        label2.Hide();
        cancelButton.Show();
        progressBar1.Show();
        label5.Show();
        clearFileListButton.Hide();
        if (!progressBar1.IsHandleCreated)
        {
            var _ = progressBar1.Handle;
        }

        timer1.Start();

        // Start the asynchronous operation.
        backgroundWorker1.RunWorkerAsync();
    }

    private void button1_Click_2(object sender, EventArgs e)
    {
        var dialog = new FolderBrowserDialog
        {
            ShowNewFolderButton = true
        };
        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        directoryTextBox.Text = dialog.SelectedPath;
    }

    private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
    {
        _stopwatch.Restart();
        if (sender is BackgroundWorker worker)
        {
            var mergeSettings = new MergeSettings()
            {
                OnlyMergeLatestSheet = onlyMergeLatestSheetCheckBox.Checked,
                EnableDebugSheet = enableDebugSheet.Checked
            };
            ExcelMerger.MergeSheets(directoryTextBox.Text, fileListBox.Items.Cast<string>(), worker, e, mergeSettings);
        }

        _stopwatch.Stop();
    }

    private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
        progressBar1.PerformStep();
    }

    private void backgroundWorker1_RunWorkerCompleted(
        object sender, RunWorkerCompletedEventArgs e)
    {
        timer1.Stop();
        // First, handle the case where an exception was thrown.
        if (e.Error != null)
        {
            using var dlg = new ErrorDialog(e.Error.Message);
            dlg.ShowDialog();
        }
        else if (e.Cancelled)
        {
            // Next, handle the case where the user canceled 
            // the operation.
            // Note that due to a race condition in 
            // the DoWork event handler, the Cancelled
            // flag may not have been set, even though
            // CancelAsync was called.
            MessageBox.Show(@"Processing Canceled", @"Canceled", MessageBoxButtons.OK);
        }
        else
        {
            // Finally, handle the case where the operation 
            // succeeded.
            var answer = MessageBox.Show($@"Completed in {_stopwatch.Elapsed.TotalSeconds} seconds. Open merged file?",
                "Completed",
                MessageBoxButtons.YesNo);

            if (answer == DialogResult.Yes)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = ExcelMerger.GetMergedFilePath(directoryTextBox.Text),
                    UseShellExecute = true,
                });
            }
            else
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = directoryTextBox.Text,
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
        }


        progressBar1.Hide();
        progressBar1.Value = 0;
        label5.Hide();
        cancelButton.Hide();
        fileListBox.Show();
        processButton.Show();
        fileBrowseButton.Show();
        directoryBrowseButton.Show();
        directoryTextBox.Show();
        label1.Show();
        label2.Show();
        clearFileListButton.Show();
    }

    private void timer1_Tick(object sender, EventArgs e)
    {
        label5.Text = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss\.f");
    }

    private void cancelButton_Click(object sender, EventArgs e)
    {
        if (backgroundWorker1.WorkerSupportsCancellation)
            // Cancel the asynchronous operation.
            backgroundWorker1.CancelAsync();
    }

    private void clearFileListButton_Click(object sender, EventArgs e)
    {
        fileListBox.Items.Clear();
    }
}