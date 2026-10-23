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
        // 列模式下该列表存储的是关键“行”索引（转置后的列索引）
        [ObservableProperty]
        private List<int> _keyIndices = new();

        // 用于生成 DataGrid 列头的配置列表
        public ObservableCollection<ColumnConfig> ColumnHeaderList { get; set; } = new();

        // 列模式标志（构建时确定）：true 表示按列对比，两个 DataGrid 上下堆叠显示
        public bool IsColumnMode { get; set; }

        // 列模式：两个文件原始的“列模型”（转置后的列数据），供重新比对使用
        public List<ExcelRowModel> LeftColumnModels { get; set; } = new();
        public List<ExcelRowModel> RightColumnModels { get; set; } = new();

        // 列模式：每个对齐列的状态，供水平 Minimap 绑定
        public ObservableCollection<ColumnStatusModel> ColumnStatuses { get; set; } = new();

        // 列模式：每个对齐列的表头文本（使用原始列号）。null 表示行模式（默认按顺序编号）
        public List<string>? ColumnHeaderLabels { get; set; }
    }
}
