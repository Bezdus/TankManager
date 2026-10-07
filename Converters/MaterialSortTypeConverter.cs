using System;
using System.Globalization;
using System.Windows.Data;
using TankManager.Core.Models;

namespace TankManager
{
    public class MaterialSortTypeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is MaterialSortType sortType)
            {
                switch (sortType)
                {
                    case MaterialSortType.ByName:
                        return "по названию";
                    case MaterialSortType.ByMass:
                        return "по массе";
                    case MaterialSortType.ByLength:
                        return "по длине";
                    case MaterialSortType.ByCost:
                        return "по стоимости";
                    case MaterialSortType.ByDate:
                        return "по дате";
                    case MaterialSortType.ByAssemblyOrder:
                        return "как в сборке";
                    default:
                        return "сортировка";
                }
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
