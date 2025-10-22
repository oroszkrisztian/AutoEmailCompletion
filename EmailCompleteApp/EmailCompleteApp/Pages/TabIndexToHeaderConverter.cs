using System;
using System.Globalization;
using System.Windows.Data;

namespace EmailCompleteApp.Pages
{
    public class TabIndexToHeaderConverter : IValueConverter
    {
        public object? Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                var index = value is int i ? i : System.Convert.ToInt32(value);
                return index switch
                {
                    0 => "Client",
                    1 => "Transportator",
                    _ => ""
                };
            }
            catch
            {
                return "";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => Binding.DoNothing;
    }
}
