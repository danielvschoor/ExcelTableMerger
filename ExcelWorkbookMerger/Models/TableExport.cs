using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWorkbookMerger.Models
{
    public class TableExport
    {
        public override int GetHashCode() => NewTableName.GetHashCode();
        
        public string FileName { get; set; }
        public string NewTableName { get; set; }
        public string OriginalTableName { get; set; }

        public void Deconstruct(out string fileName, out string newTableName, out string originalTableName)
        {
            fileName = FileName;
            newTableName = NewTableName;
            originalTableName = OriginalTableName;
        }
    }
    
}
