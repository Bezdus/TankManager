using System;
using System.Globalization;
using System.Windows.Data;
using TankManager.Core.Models;

namespace TankManager
{
    public class ProductTypeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ProductType productType)
            {
                switch (productType)
                {
                    case ProductType.Part:
                        return "Деталь";
                    case ProductType.PurchasedPart:
                        return "Покупное изделие";
                    case ProductType.SheetMaterial:
                        return "Листовой прокат";
                    case ProductType.TubularProduct:
                        return "Трубный прокат";
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
