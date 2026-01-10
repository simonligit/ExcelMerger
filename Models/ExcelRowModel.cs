using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerger.Models
{
    // ObservableObject 让这个类具备“属性更改通知”的能力
    public partial class ExcelRowModel : ObservableObject
    {
        // 存储这一行所有的列数据
        [ObservableProperty]
        private List<string> _cells = new();

        // 用于标识这一行在对比中的状态：例如 "Same", "Modified", "Added", "Deleted"
        [ObservableProperty]
        private string? _status;

        // 用于在界面上控制高亮颜色（我们稍后会绑定到这个属性）
        [ObservableProperty]
        private bool _isDifferent;

        // 存储发生变化的列索引
        [ObservableProperty]
        private HashSet<int> _diffColumnIndices = new();

        // 存储该行在原始文件中的行号（从 1 开始）
        [ObservableProperty]
        private int _originalIndex;
        
        // 获取用于逻辑匹配的“关键特征”：默认取第一个非空单元格
        public string GetKey() => Cells.FirstOrDefault(c => !string.IsNullOrWhiteSpace(c))?.Trim() ?? string.Empty;

        // 获取用于内容对比的“全行特征”：将所有单元格拼接
        public string GetFullContent() => string.Join("|", Cells.Select(c => c?.Trim() ?? string.Empty));

        // 判断两行是否内容完全一致
        public bool IsSameContent(ExcelRowModel other)
        {
            if (other == null) return false;
            // 如果列数都不一样，肯定不同
            if (this.Cells.Count != other.Cells.Count) return false;
            return GetFullContent() == other.GetFullContent();
        }
    }
}