namespace ExcelWorkbookMerger.Models
{
    public record TableExport
    {
        public override int GetHashCode() => NewTableName.GetHashCode();
        
        public required string FileName { get; init; }
        public required string NewTableName { get; init; }
        public required string OriginalTableName { get; init; }

        public void Deconstruct(out string fileName, out string newTableName, out string originalTableName)
        {
            fileName = FileName;
            newTableName = NewTableName;
            originalTableName = OriginalTableName;
        }
    }
    
}
