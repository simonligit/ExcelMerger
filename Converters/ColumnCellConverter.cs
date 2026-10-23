using ExcelMerger.Models;
using ExcelMerger.Utils;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelMerger.Converters
{
    /// <summary>
    /// 列模式下的单元格转换器：根据该列的对齐状态（ColumnStatusMap）设置整列背景色，
    /// 并对修改列中发生变化的单元格设置红色加粗文字（与行模式的 CellDiffConverter 逻辑对应）
    /// </summary>
    public class ColumnCellConverter : IMultiValueConverter
    {
        // values[0]: ExcelRowModel 对象（显示行）
        // values[1]: 当前列的索引 (int)
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length < 2 || values[0] is not ExcelRowModel row || values[1] is not int colIndex)
                return DependencyProperty.UnsetValue;

            row.ColumnStatusMap.TryGetValue(colIndex, out var status);
            bool cellDiffers = status == StatusCodes.Modified && row.DiffColumnIndices.Contains(colIndex);

            switch (parameter as string)
            {
                case "Background":
                    return status switch
                    {
                        StatusCodes.Added => new SolidColorBrush(Colors.LightGreen),
                        StatusCodes.Deleted => new SolidColorBrush(Colors.LightPink),
                        StatusCodes.Modified => new SolidColorBrush(Colors.Moccasin),
                        StatusCodes.Empty => new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)),
                        _ => DependencyProperty.UnsetValue
                    };
                case "Foreground":
                    return cellDiffers ? new SolidColorBrush(Colors.Red) : DependencyProperty.UnsetValue;
                case "Weight":
                    return cellDiffers ? FontWeights.Bold : DependencyProperty.UnsetValue;
                default:
                    return DependencyProperty.UnsetValue;
            }
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) => null;
    }
}
