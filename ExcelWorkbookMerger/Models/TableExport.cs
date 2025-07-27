using System.Data;

namespace ExcelWorkbookMerger.Models;

public record TableExport
{
    public string FileName { get; init; } = string.Empty;
    
    public DataTable DataTable { get; init; } = new();
}