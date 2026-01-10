using CommunityToolkit.Mvvm.ComponentModel;
using ExcelMerger.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace ExcelMerger.ViewModels
{
    public partial class SheetComparisonViewModel : ObservableObject
    {
        public string SheetName { get; set; }

        // 该工作表左侧的行
        public ObservableCollection<ExcelRowModel> LeftRows { get; set; } = new();

        // 该工作表右侧的行
        public ObservableCollection<ExcelRowModel> RightRows { get; set; } = new();

        [ObservableProperty]
        private int _maxColumns;

        // 存储该 Sheet 的关键字列索引列表。例如：[0, 1] 表示前两列组合成复合 Key
        [ObservableProperty]
        private List<int> _keyIndices = new();

        // 用于生成 DataGrid 列头的配置列表
        public ObservableCollection<ColumnConfig> ColumnHeaderList { get; set; } = new();
    }
}
