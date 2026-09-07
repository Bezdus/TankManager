using System.Windows;
using TankManager.Core.Models;

namespace TankManager.Views
{
    public partial class PricingSettingsDialog : Window
    {
        public PricingSettings PricingSettings { get; private set; }

        public PricingSettingsDialog(PricingSettings settings)
        {
            InitializeComponent();
            PricingSettings = new PricingSettings
            {
                SheetMetalPricePerKg = settings.SheetMetalPricePerKg,
                OtherMetalPricePerKg = settings.OtherMetalPricePerKg,
                LaserCuttingPricePerMm = settings.LaserCuttingPricePerMm,
                EngravingPricePerMm = settings.EngravingPricePerMm,
                BendingPricePerOperation = settings.BendingPricePerOperation,
                RollingPricePerMm = settings.RollingPricePerMm,
                FlangingPricePerOperation = settings.FlangingPricePerOperation
            };

            foreach (var entry in settings.TubularPricing)
            {
                PricingSettings.TubularPricing.Add(new TubularPricingEntry
                {
                    Size = entry.Size,
                    PricePerMeter = entry.PricePerMeter
                });
            }

            DataContext = PricingSettings;
            TubularGrid.ItemsSource = PricingSettings.TubularPricing;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}