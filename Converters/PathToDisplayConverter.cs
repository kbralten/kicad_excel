using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;

namespace KiCadExcelBridge.Converters
{
    public class PathToDisplayConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return string.Empty;
            var s = value.ToString();
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;

            try
            {
                // Use the file name for display; fall back to the full value
                return Path.GetFileName(s);
            }
            catch
            {
                return s;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // We don't convert back from display name to full path
            return Binding.DoNothing;
        }
    }
}
