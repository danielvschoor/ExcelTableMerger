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
    private static readonly ConcurrentDictionary<string, ConcurrentQueue<TableExport>> DataTables = new();
    private static readonly ConcurrentQueue<Exception> Exceptions = new();

    private static bool IsDebugEnabled => MergeSettings?.EnableDebugSheet == true;
    private static readonly Lock DebugLock = new();
    private static readonly Lock LogLock = new();
    private static readonly Lock ProgressLock = new();


    private static readonly DataTable DebugSheetDataTable = new()
    {
        TableName = "DebugSheets",
        Columns = { "File", "SheetsParsed", "SheetsNotParsed", "SheetsUsedForMerge" },
    };

    private static readonly DataTable DebugLogDataTable = new()
    {
        TableName = "DebugLogs",
        Columns = { "File", "Sheet", "LogLevel", "LogMessage" },
    };

    private static void AddLogMessage(LogLevel loggingLevel, string message, string? file = null,
        string? sheet = null)
    {
        if (!IsDebugEnabled)
        {
            return;
        }

        using (LogLock.EnterScope())
        {
            DebugLogDataTable.Rows.Add(file, sheet, loggingLevel.ToString(), message);
        }
    }

    private static void AddDebugSheetInfo(string file, string sheetsParsed, string sheetsNotParsed,
        string sheetsUsedForMerge)
    {
        if (!IsDebugEnabled)
        {
            return;
        }

        using (DebugLock.EnterScope())
        {
            DebugSheetDataTable.Rows.Add(file, sheetsParsed, sheetsNotParsed, sheetsUsedForMerge);
        }
    }

    private static MergeSettings? MergeSettings { get; set; }

    private static void ClearFields()
    {
        DataTables.Clear();
        Exceptions.Clear();
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

        ProcessDataTables(files, worker, workArgs);

        CheckExceptions();

        using var mergedWorkbook = new ExcelPackage(newFile);
        if (ExportToWorkBook(worker, workArgs, mergedWorkbook))
        {
            mergedWorkbook.SaveAs(newFile);
        }
    }

    private static bool ExportToWorkBook(BackgroundWorker worker, DoWorkEventArgs workArgs, ExcelPackage mergedWorkbook)
    {
        foreach (var (newTableName, innerTables) in DataTables
                     .OrderBy(x => x.Key))
        {
            try
            {
                if (ShouldCancel(worker, workArgs)) return false;

                var finalDt = new DataTable();

                AddLogMessage(LogLevel.Warning, $"{newTableName} tableCount={innerTables.Count}", null, newTableName);

                foreach (var tempDt in innerTables)
                {
                    if (ShouldCancel(worker, workArgs)) return false;

                    finalDt.Merge(tempDt.DataTable);
                }

                var safeTableName = newTableName.Truncate(31);
                var finalWorksheet =
                    mergedWorkbook.Workbook.Worksheets.FirstOrDefault(x => x.Name == safeTableName) ??
                    mergedWorkbook.Workbook.Worksheets.Add(safeTableName);
                
                // If not table exists, create it
                if (finalWorksheet.Tables.All(x => x.Name != safeTableName))
                {
                    finalWorksheet.Tables.Add(new ExcelAddressBase(1, 1, finalDt.Rows.Count + 1,
                        finalDt.Columns.Count), safeTableName);
                }
                
                finalWorksheet.Cells["A1"].LoadFromDataTable(finalDt, true);
            }
            catch (Exception ex)
            {
                throw new Exception(
                    $"Exception occurred while exporting workbook.\nException: {ex}");
            }
        }

        if (!IsDebugEnabled)
        {
            return !ShouldCancel(worker, workArgs);
        }

        var debugWorksheet =
            mergedWorkbook.Workbook.Worksheets.FirstOrDefault(x => x.Name == "DebugLogs") ??
            mergedWorkbook.Workbook.Worksheets.Add("DebugLogs");

        if (debugWorksheet == null)
        {
            return !ShouldCancel(worker, workArgs);
        }

        debugWorksheet.Cells["A1"].LoadFromDataTable(DebugLogDataTable, true);
        debugWorksheet = mergedWorkbook.Workbook.Worksheets.FirstOrDefault(x => x.Name == "DebugSheets") ??
                         mergedWorkbook.Workbook.Worksheets.Add("DebugSheets");
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

                var sheetsParsed = new List<string>();
                var sheetsNotParsed = new List<string>();
                var sheetsUsedForMerge = new List<string>();
                IEnumerable<ExcelWorksheet> sheetsToProcess;

                using var excelPackage = new ExcelPackage(new FileInfo(file));
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

                    sheetsToProcess = [latestSheet];
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
                        AddLogMessage(LogLevel.Warning, $"Found table {table.Name}", file, sheet.Name);
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

                            AddLogMessage(LogLevel.Warning, $"{table.Name} has {tempDt.Rows.Count}rows ", file,
                                sheet.Name);

                            var dtList = DataTables.GetOrAdd(newTableName
                                , _ => new ConcurrentQueue<TableExport>());

                            dtList.Enqueue(new TableExport
                            {
                                FileName = file, DataTable = tempDt
                            });
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(
                                $"Exception in File: {file}\nSheet:{sheet}\nTable: {table.Name}:\n{ex}", ex);
                        }
                    }
                }

                AddDebugSheetInfo(file, string.Join(',', sheetsParsed), string.Join(',', sheetsNotParsed),
                    string.Join(',', sheetsUsedForMerge));
                ReportProgress(worker, StepSize);
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

    private static void ReportProgress(BackgroundWorker worker, int progress)
    {
        using (ProgressLock.EnterScope())
        {
            worker.ReportProgress(progress);
        }
    }

    [GeneratedRegex("[\\d-]")]
    private static partial Regex DigitRegex();
}