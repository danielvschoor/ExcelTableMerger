using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelWorkbookMerger.ExtensionMethods;
using ExcelWorkbookMerger.Models;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;

namespace ExcelWorkbookMerger.Logic;

public static partial class ExcelMerger
{
    private const int StepSize = 1;
    private const string FileName = @"MergedExcel.xlsx";
    private static readonly ConcurrentDictionary<TableExport, ConcurrentQueue<DataTable>> DataTables = new();
    private static readonly ConcurrentQueue<Exception> Exceptions = new();
    private static ExcelPackage? _mergedWorkbook;

    private static void ClearFields()
    {
        DataTables.Clear();
        Exceptions.Clear();
        _mergedWorkbook = null;
    }

    public static void MergeSheets(string path, IEnumerable<string> files, BackgroundWorker worker,
        DoWorkEventArgs workArgs)
    {
        ClearFields();
        var newFile = CheckResultsFile(path);
        if (newFile == null)
            return;

        if (ShouldCancel(worker, workArgs)) return;

        _mergedWorkbook = new ExcelPackage(newFile);


        ProcessDataTables(files, worker, workArgs);

        CheckExceptions();

        if (ExportToWorkBook(worker, workArgs)) _mergedWorkbook.SaveAs(newFile);
    }

    private static bool ExportToWorkBook(BackgroundWorker worker, DoWorkEventArgs workArgs)
    {
        foreach (var ((fileName,newTableName,originalTableName), dataTables) in DataTables)
        {
            try
            {
                if (ShouldCancel(worker, workArgs)) return false;

                var finalDt = new DataTable();
                foreach (var tempDt in dataTables)
                {
                    if (ShouldCancel(worker, workArgs)) return false;

                    finalDt.Merge(tempDt);
                }

                var finalWorksheet = _mergedWorkbook?.Workbook.Worksheets.FirstOrDefault(x => x.Name == newTableName.Truncate(31)) ??
                                     _mergedWorkbook?.Workbook.Worksheets.Add(newTableName);
                if (finalWorksheet == null) return false;
                if (finalWorksheet.Tables.All(x => x.Name != newTableName))
                    finalWorksheet.Tables.Add(new ExcelAddressBase(1, 1, finalDt.Rows.Count + 1,
                        finalDt.Columns.Count), newTableName);

                finalWorksheet.Cells["A1"].LoadFromDataTable(finalDt, true);
            }
            catch (Exception ex)
            {
                throw new Exception($"Exception occurred while exporting workbook.\nException originated from file:\n{fileName}.\n\nException: {ex}");
            }
        }

        return !ShouldCancel(worker, workArgs);
    }

    private static void ProcessDataTables(IEnumerable<string> files, BackgroundWorker worker, DoWorkEventArgs workArgs)
    {
        Parallel.ForEach(files, (file, state) =>
        {
            try
            {
                if (ShouldCancel(worker, workArgs) || state.IsStopped)
                {
                    state.Stop();
                    return;
                }

                var excelPackage = new ExcelPackage(new FileInfo(file));
                foreach (var sheet in excelPackage.Workbook.Worksheets)
                {
                    foreach (var table in sheet.Tables)
                    {
                        try
                        {
                            var newTableName = table.Name.Split('_')[0];
                            newTableName = DigitRegex().Replace(newTableName, string.Empty);
                            newTableName = newTableName + '_' + sheet.Name;

                            var tempDt = new DataTable();
                            tempDt.Columns.AddRange(table.Columns.Select(x => new DataColumn(x.Name)).ToArray());
                            var tableRange = table.Range;
                            sheet.Cells[tableRange.Start.Row, tableRange.Start.Column, tableRange.End.Row,
                                tableRange.End.Column].ToDataTable(x =>
                            {
                                x.EmptyRowStrategy = EmptyRowsStrategy.Ignore;
                                x.FirstRowIsColumnNames = true;
                                x.ExcelErrorParsingStrategy = ExcelErrorParsingStrategy.HandleExcelErrorsAsBlankCells;
                            }, tempDt);
                            var dtList = DataTables.GetOrAdd(new TableExport
                                {
                                    FileName = file,
                                    NewTableName = newTableName,
                                    OriginalTableName = table.Name
                                }
                                , _ => new ConcurrentQueue<DataTable>());
                            dtList.Enqueue(tempDt);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(
                                $"Exception in File: {file}\nSheet:{sheet}\nTable: {table.Name}:\n{ex}");
                        }
                    }
                }

                worker.ReportProgress(StepSize);
            }
            catch (Exception exception)
            {
                var newException = new Exception($"Exception occurred while processing file {file}: {exception}");
                Exceptions.Enqueue(newException);
            }
        });
    }

    private static FileInfo? CheckResultsFile(string path)
    {
        var mergedFilePath = Path.Combine(path, FileName);
        var newFile = new FileInfo(mergedFilePath);

        if (!newFile.Exists)
            return newFile;
        var message = $"The file \"{FileName}\" already exists at {path}. Delete file?";
        if (MessageBox.Show(message,
                @"File exists", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
            return null;

        newFile.Delete(); // ensures we create a new workbook
        newFile = new FileInfo(mergedFilePath);
        return newFile;
    }

    private static void CheckExceptions()
    {
        if (!Exceptions.IsEmpty) throw new AggregateException(Exceptions);
    }

    private static bool ShouldCancel(BackgroundWorker worker, DoWorkEventArgs workArgs)
    {
        if (!worker.CancellationPending)
            return false;
        workArgs.Cancel = true;
        return true;
    }

    [GeneratedRegex("[\\d-]")]
    private static partial Regex DigitRegex();
}