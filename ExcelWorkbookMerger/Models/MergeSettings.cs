using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWorkbookMerger.Models
{
    public class MergeSettings
    {
        public required bool EnableDebugSheet { get; set; }
        public required bool OnlyMergeLatestSheet { get; set; }
    }
}
