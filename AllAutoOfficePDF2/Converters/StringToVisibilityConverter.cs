using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace AllAutoOfficePDF2.Converters
{
    /// <summary>
    /// 文字列からVisibilityへの変換コンバーター
    /// </summary>
    public class StringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string str && !string.IsNullOrEmpty(str))
            {
                return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}