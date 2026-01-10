using ExcelMerger.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelMerger.Converters
{
    public class CellDiffConverter : IValueConverter, IMultiValueConverter
    {
        // --- IMultiValueConverter 成员 (用于 Foreground 的 MultiBinding) ---
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            // values[0]: ExcelRowModel 对象
            // values[1]: 当前列的索引 (int)
            if (values.Length >= 2 && values[0] is ExcelRowModel row && values[1] is int colIndex)
            {
                if (row.Status == "Modified" && row.DiffColumnIndices.Contains(colIndex))
                {
                    return new SolidColorBrush(Colors.Red); // 发现差异则返回红色
                }
            }
            return DependencyProperty.UnsetValue; // 否则不设置（沿用 RowStyle）
        }

        // --- IValueConverter 成员 (用于 FontWeight 的普通 Binding) ---
        // 注意：普通 Binding 只能传回 RowModel，无法直接传回列索引
        // 所以在设置 FontWeight 时，我们最好也改用 MultiBinding
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }

        // 统一的逻辑处理函数
        private object GetValidationResult(ExcelRowModel row, int colIndex, string type)
        {
            if (row.Status == "Modified" && row.DiffColumnIndices.Contains(colIndex))
            {
                if (type == "Weight") return FontWeights.Bold;
                return Brushes.Red;
            }

            if (type == "Weight") return FontWeights.Normal;
            return DependencyProperty.UnsetValue;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) => null;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => null;
    }
}
