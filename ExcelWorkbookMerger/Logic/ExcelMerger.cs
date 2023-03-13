using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelWorkbookMerger.Models;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using static OfficeOpenXml.ExcelErrorValue;

namespace ExcelWorkbookMerger.Logic;

public static partial class ExcelMerger
{
    private const int StepSize = 1;
    private const string FileName = @"MergedExcel.xlsx";
    private static readonly ConcurrentDictionary<string, TableExport> TablesToExport = new();
    private static readonly ConcurrentDictionary<string, ConcurrentQueue<DataTable>> DataTables = new();
    private static readonly ConcurrentQueue<Exception> Exceptions = new();
    private static ExcelPackage? _mergedWorkbook;

    private static void ClearFields()
    {
        TablesToExport.Clear();
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
        foreach (KeyValuePair<string, TableExport> excelTable in TablesToExport)
        {
            try
            {
                if (ShouldCancel(worker, workArgs)) return false;

                var finalDt = new DataTable();
                foreach (var tempDt in DataTables[excelTable.Key])
                {
                    if (ShouldCancel(worker, workArgs)) return false;

                    finalDt.Merge(tempDt);
                }

                var finalWorksheet = _mergedWorkbook?.Workbook.Worksheets.FirstOrDefault(x => x.Name == excelTable.Value.NewTableName) ??
                                     _mergedWorkbook?.Workbook.Worksheets.Add(excelTable.Value.NewTableName);
                if (finalWorksheet == null) return false;
                if (finalWorksheet.Tables.All(x => x.Name != excelTable.Value.NewTableName))
                    finalWorksheet.Tables.Add(new ExcelAddressBase(1, 1, finalDt.Rows.Count + 1,
                        finalDt.Columns.Count), excelTable.Value.NewTableName);

                finalWorksheet.Cells["A1"].LoadFromDataTable(finalDt, true);
            }catch (Exception ex)
            {
                throw new Exception($"Exception occurred while exporting workbook.\nException originated from file:\n{excelTable.Value.FileName}.\n\nException: {ex}");
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
                if (ShouldCancel(worker, workArgs)|| state.IsStopped)
                {
                    state.Stop();
                    return;
                }

                var excelPackage = new ExcelPackage(new FileInfo(file));

                foreach (var sheet in excelPackage.Workbook.Worksheets)
                {
                    foreach (var tableName in sheet.Tables.Select(x => x.Name))
                    {
                        try
                        {
                            var newTableName = tableName.Split('_')[0];
                            newTableName = DigitRegex().Replace(newTableName, string.Empty);
                            newTableName = newTableName + '_' + sheet.Name;
                            TablesToExport.TryAdd(tableName, new() { FileName = file, NewTableName = newTableName });
                        }catch(Exception ex)
                        {
                            throw new Exception($"Exception in File: {file}\nSheet:{sheet}\nTable: {tableName}:\n{ex}");
                        }
                    }
                    foreach (var excelTableName in TablesToExport.Keys)
                    {
                        if (ShouldCancel(worker, workArgs) || state.IsStopped)
                        {
                            state.Stop();
                            return;
                        }

                        var table = sheet.Tables.FirstOrDefault(x => x.Name == excelTableName);
                        if (table == null) continue;


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
                        var dtList = DataTables.GetOrAdd(excelTableName, _ => new ConcurrentQueue<DataTable>());
                        dtList.Enqueue(tempDt);
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