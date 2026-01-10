using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerger.Models
{
    public class ExcelSheetModel
    {
        public required string SheetName { get; set; }
        public List<ExcelRowModel> Rows { get; set; } = new();

        // 记录该表中最长的一行有多少列
        public int MaxColumnCount { get; set; } = 0;
    }
}
