using System;
using System.Globalization;
using System.Windows.Data;

namespace ExcelMerger.Converters
{
    /// <summary>
    /// 布尔取反转换器：用于两个互斥选项（如 行模式/列模式 单选框）的双向绑定
    /// </summary>
    public class InverseBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b ? !b : value;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b ? !b : value;
    }
}
