using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace TankManager
{
    /// <summary>
    /// Число ↔ текст для полей ввода: принимает и запятую, и точку, пробелы между разрядами,
    /// пустое поле = 0. Параметр — формат отображения (по умолчанию без лишних нулей)
    /// </summary>
    public class DecimalTextConverter : IValueConverter
    {
        private const string DefaultFormat = "0.####";

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double number)
                return number.ToString(parameter as string ?? DefaultFormat, culture);
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string text = (value as string ?? string.Empty)
                .Replace(" ", string.Empty)
                .Replace(" ", string.Empty)
                .Replace(',', '.');

            if (text.Length == 0)
                return 0d;

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) &&
                !double.IsNaN(result) && !double.IsInfinity(result))
                return result;

            // Ошибка проверки: поле подсвечивается, значение в настройках не меняется
            return DependencyProperty.UnsetValue;
        }
    }
}
