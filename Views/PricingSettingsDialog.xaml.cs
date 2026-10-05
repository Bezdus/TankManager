using System.Windows;
using System.Windows.Controls;
using TankManager.Core.Models;

namespace TankManager.Views
{
    public partial class PricingSettingsDialog : Window
    {
        public PricingSettings PricingSettings { get; private set; }

        public PricingSettingsDialog(PricingSettings settings, bool isReadOnly = false)
        {
            InitializeComponent();
            PricingSettings = new PricingSettings
            {
                SheetMetalPricePerKg = settings.SheetMetalPricePerKg,
                OtherMetalPricePerKg = settings.OtherMetalPricePerKg,
                LaserCuttingPricePerMm = settings.LaserCuttingPricePerMm,
                EngravingPricePerMm = settings.EngravingPricePerMm,
                BendingPricePerOperation = settings.BendingPricePerOperation,
                RollingPricePerKg = settings.RollingPricePerKg,
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

            if (isReadOnly)
                MakeReadOnly();
        }

        /// <summary>
        /// Только просмотр: расценки общие, меняет их конструктор
        /// </summary>
        private void MakeReadOnly()
        {
            Title = "Расценки (просмотр)";
            ReadOnlyNote.Visibility = Visibility.Visible;
            SaveButton.Visibility = Visibility.Collapsed;
            CancelButton.Content = "Закрыть";
            CancelButton.IsDefault = true;
            CancelButton.Margin = new Thickness(0);

            TubularGrid.IsReadOnly = true;
            TubularGrid.CanUserAddRows = false;
            TubularGrid.CanUserDeleteRows = false;
            SetTextBoxesReadOnly(this);
        }

        private static void SetTextBoxesReadOnly(DependencyObject parent)
        {
            foreach (var child in LogicalTreeHelper.GetChildren(parent))
            {
                if (child is TextBox textBox)
                    textBox.IsReadOnly = true;

                if (child is DependencyObject dependencyObject)
                    SetTextBoxesReadOnly(dependencyObject);
            }
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