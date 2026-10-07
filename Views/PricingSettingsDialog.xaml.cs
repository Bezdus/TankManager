using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
                LaserCuttingPricePerMeter = settings.LaserCuttingPricePerMeter,
                EngravingPricePerMm = settings.EngravingPricePerMm,
                BendingPricePerOperation = settings.BendingPricePerOperation,
                RollingPricePerKg = settings.RollingPricePerKg,
                FlangingPricePerOperation = settings.FlangingPricePerOperation,
                ModifiedBy = settings.ModifiedBy,
                ModifiedByName = settings.ModifiedByName,
                ModifiedUtcTicks = settings.ModifiedUtcTicks
            };

            foreach (var entry in settings.TubularPricing ?? Enumerable.Empty<TubularPricingEntry>())
            {
                PricingSettings.TubularPricing.Add(new TubularPricingEntry
                {
                    Size = entry.Size,
                    PricePerMeter = entry.PricePerMeter
                });
            }

            foreach (var entry in settings.LaserCuttingPricing ?? Enumerable.Empty<LaserCuttingPricingEntry>())
            {
                PricingSettings.LaserCuttingPricing.Add(new LaserCuttingPricingEntry
                {
                    Thickness = entry.Thickness,
                    PricePerMeter = entry.PricePerMeter
                });
            }

            DataContext = PricingSettings;
            TubularGrid.ItemsSource = PricingSettings.TubularPricing;
            LaserGrid.ItemsSource = PricingSettings.LaserCuttingPricing;

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
            LaserGrid.IsReadOnly = true;
            LaserGrid.CanUserAddRows = false;
            LaserGrid.CanUserDeleteRows = false;
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
            NumberInput.CommitFocusedTextBox();

            foreach (var grid in new[] { TubularGrid, LaserGrid })
            {
                grid.CommitEdit(DataGridEditingUnit.Cell, true);
                grid.CommitEdit(DataGridEditingUnit.Row, true);
            }

            if (NumberInput.HasValidationErrors(this))
            {
                NumberInput.ShowValidationWarning(this);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void NumberBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
            => NumberInput.OnGotKeyboardFocus(sender);

        private void NumberBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            => NumberInput.OnPreviewMouseLeftButtonDown(sender, e);

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}