using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
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

    private static readonly Lock Lock = new();
    
    private static readonly DataTable DebugSheetDataTable = new()
    {
        TableName = "DebugSheets",
        Columns = {"File", "SheetsParsed", "SheetsNotParsed", "SheetsUsedForMerge"},
    };

    private static readonly DataTable DebugLogDataTable = new()
    {
        TableName = "DebugLogs",
        Columns = {"File", "Sheet", "LogLevel", "LogMessage"},
    };

    private static void AddLogMessage(LogLevel loggingLevel, string message, string? file = null,
        string? sheet = null)
    {
        if (!(MergeSettings?.EnableDebugSheet ?? false))
        {
            return;
        }
        lock (Lock)
        {
            DebugLogDataTable.Rows.Add(file, sheet, loggingLevel.ToString(), message);
        }
    }

    private static void AddDebugSheetInfo(string file, string sheetsParsed, string sheetsNotParsed,
        string sheetsUsedForMerge)
    {
        if (!(MergeSettings?.EnableDebugSheet ?? false))
        {
            return;
        }

        lock (Lock)
        {
            DebugSheetDataTable.Rows.Add(file, sheetsParsed, sheetsNotParsed, sheetsUsedForMerge);
        }
    }

    private static MergeSettings? MergeSettings { get; set; }
    private static ExcelPackage? _mergedWorkbook;

    private static void ClearFields()
    {
        DataTables.Clear();
        Exceptions.Clear();
        _mergedWorkbook = null;
        DebugLogDataTable.Clear();
        DebugSheetDataTable.Clear();
    }

    public static void MergeSheets(string path, IEnumerable<string> files, BackgroundWorker worker,
        DoWorkEventArgs workArgs, MergeSettings mergeSettings)
    {
        ClearFields();
        MergeSettings = mergeSettings;

        var newFile = CheckResultsFile(path);

        if (newFile == null)
            return;

        if (ShouldCancel(worker, workArgs)) return;

        _mergedWorkbook = new ExcelPackage(newFile);

        ProcessDataTables(files, worker, workArgs);

        CheckExceptions();

        if (ExportToWorkBook(worker, workArgs))
        {
            _mergedWorkbook.SaveAs(newFile);
        }
    }

    private static bool ExportToWorkBook(BackgroundWorker worker, DoWorkEventArgs workArgs)
    {
        foreach (var ((fileName, newTableName, _), dataTables) in DataTables)
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

                var finalWorksheet =
                    _mergedWorkbook?.Workbook.Worksheets.FirstOrDefault(x => x.Name == newTableName.Truncate(31)) ??
                    _mergedWorkbook?.Workbook.Worksheets.Add(newTableName);
                if (finalWorksheet == null) return false;
                if (finalWorksheet.Tables.All(x => x.Name != newTableName))
                    finalWorksheet.Tables.Add(new ExcelAddressBase(1, 1, finalDt.Rows.Count + 1,
                        finalDt.Columns.Count), newTableName);

                finalWorksheet.Cells["A1"].LoadFromDataTable(finalDt, true);
            }
            catch (Exception ex)
            {
                throw new Exception(
                    $"Exception occurred while exporting workbook.\nException originated from file:\n{fileName}.\n\nException: {ex}");
            }
        }

        if (!(MergeSettings?.EnableDebugSheet ?? false))
        {
            return !ShouldCancel(worker, workArgs);
        }

        var debugWorksheet =
            _mergedWorkbook?.Workbook.Worksheets.FirstOrDefault(x => x.Name == "DebugLogs") ??
            _mergedWorkbook?.Workbook.Worksheets.Add("DebugLogs");

        if (debugWorksheet == null)
        {
            return !ShouldCancel(worker, workArgs);
        }

        debugWorksheet.Cells["A1"].LoadFromDataTable(DebugLogDataTable, true);
        debugWorksheet = _mergedWorkbook?.Workbook.Worksheets.FirstOrDefault(x => x.Name == "DebugSheets") ??
                         _mergedWorkbook?.Workbook.Worksheets.Add("DebugSheets");
        if (debugWorksheet == null)
        {
            return !ShouldCancel(worker, workArgs);
        }

        debugWorksheet.Cells["A1"].LoadFromDataTable(DebugSheetDataTable, true);

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

                var sheetsParsed = new List<string>();
                var sheetsNotParsed = new List<string>();
                var sheetsUsedForMerge = new List<string>();
                IEnumerable<ExcelWorksheet> sheetsToProcess;

                if (MergeSettings!.OnlyMergeLatestSheet)
                {
                    ExcelWorksheet? latestSheet = null;
                    DateTime? latestDate = null;

                    foreach (var sheet in excelPackage.Workbook.Worksheets)
                    {
                        if (!DateTime.TryParseExact(sheet.Name, "MMMyy", CultureInfo.InvariantCulture,
                                DateTimeStyles.None, out var date) && !DateTime.TryParseExact(sheet.Name, "ddMMMyy",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None, out date))
                        {
                            sheetsNotParsed.Add(sheet.Name);
                            AddLogMessage(LogLevel.Warning, "Could not parse sheet name", file, sheet.Name);
                            continue;
                        }

                        sheetsParsed.Add(sheet.Name);

                        if (latestDate != null && !(date.Date > latestDate))
                        {
                            continue;
                        }

                        latestDate = date.Date;
                        latestSheet = sheet;
                    }

                    if (latestSheet == null)
                    {
                        throw new Exception(
                            $"Unable to find latest sheet for {file}. Sheets should be named as MMMyy (FEB24)");
                    }

                    AddLogMessage(LogLevel.Info, "Found latest sheet", file, latestSheet.Name);

                    sheetsToProcess = [excelPackage.Workbook.Worksheets.First(x => x.Name == latestSheet.Name)];
                }
                else
                {
                    sheetsToProcess = excelPackage.Workbook.Worksheets;
                }


                foreach (var sheet in sheetsToProcess)
                {
                    sheetsUsedForMerge.Add(sheet.Name);
                    foreach (var table in sheet.Tables)
                    {
                        try
                        {
                            var newTableName = table.Name.Split('_')[0];
                            newTableName = DigitRegex().Replace(newTableName, string.Empty);
                            if (!MergeSettings.OnlyMergeLatestSheet)
                            {
                                newTableName = newTableName + '_' + sheet.Name;
                            }

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

                AddDebugSheetInfo(file, string.Join(',', sheetsParsed), string.Join(',', sheetsNotParsed),
                    string.Join(',', sheetsUsedForMerge));

                worker.ReportProgress(StepSize);
            }
            catch (Exception exception)
            {
                var newException = new Exception($"Exception occurred while processing file {file}: {exception}");
                Exceptions.Enqueue(newException);
            }
        });
    }

    public static string GetMergedFilePath(string path)
    {
        return Path.Combine(path, FileName);
    }
    
    private static FileInfo? CheckResultsFile(string path)
    {
        var mergedFilePath = GetMergedFilePath(path);
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