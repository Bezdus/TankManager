using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TankManager.Core.Models;

namespace TankManager.Views
{
    /// <summary>
    /// Ручная правка операций изготовления детали. Работает с копиями операций:
    /// результат забирается из Operations только при «Сохранить»
    /// </summary>
    public partial class OperationsEditorDialog : Window
    {
        // Параметры операций из КОМПАС: их изменение помечает операцию как отредактированную
        private static readonly HashSet<string> ParameterProperties = new HashSet<string>
        {
            nameof(LaserCuttingOperation.CutLength),
            nameof(LaserCuttingOperation.EngravingLength),
            nameof(BendingOperation.BendAngle),
            nameof(BendingOperation.BendLength),
            nameof(RollingOperation.RollDiameter),
            nameof(RollingOperation.Radius),
            nameof(RollingOperation.Length),
            nameof(FlangingOperation.Diameter)
        };

        // Свойства, изменение которых не влияет на стоимость
        private static readonly HashSet<string> NonCostProperties = new HashSet<string>
        {
            nameof(ManufacturingOperationBase.Cost),
            nameof(ManufacturingOperationBase.IsModified),
            nameof(ManufacturingOperationBase.IsManual),
            nameof(ManufacturingOperationBase.IsEdited),
            nameof(ManufacturingOperationBase.Name)
        };

        private readonly PricingSettings _settings;
        private readonly double _partMass;
        private readonly double _sheetThickness;
        private readonly List<ManufacturingOperationBase> _original;

        public ObservableCollection<ManufacturingOperationBase> Operations { get; }

        public OperationsEditorDialog(PartModel part, PricingSettings settings, int instancesCount)
        {
            InitializeComponent();

            _settings = settings;
            _partMass = part.Mass;
            _sheetThickness = part.SheetThickness;
            _original = part.Operations.Select(op => op.Clone()).ToList();

            Operations = new ObservableCollection<ManufacturingOperationBase>();
            foreach (var op in _original)
                AddWorkingCopy(op.Clone());

            PartTitle.Text = string.IsNullOrEmpty(part.Marking) ? part.Name : $"{part.Marking} — {part.Name}";
            InstancesNote.Text = instancesCount > 1
                ? $"Изменения применятся ко всем экземплярам детали в сборке ({instancesCount} шт.) и сразу сохранятся на сервер"
                : "Изменения сразу сохранятся на сервер";

            OperationsList.ItemsSource = Operations;
            Operations.CollectionChanged += (s, e) => UpdateTotal();
            UpdateTotal();
        }

        private void AddWorkingCopy(ManufacturingOperationBase op)
        {
            if (op is RollingOperation rolling)
                rolling.PartMass = _partMass;

            if (op is LaserCuttingOperation laser)
                laser.MaterialThickness = _sheetThickness;

            op.CalculateCost(_settings);
            op.PropertyChanged += Operation_PropertyChanged;
            Operations.Add(op);
        }

        private void Operation_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var op = (ManufacturingOperationBase)sender;

            if (!op.IsManual && ParameterProperties.Contains(e.PropertyName))
                op.IsEdited = true;

            if (NonCostProperties.Contains(e.PropertyName))
                return;

            op.CalculateCost(_settings);
            UpdateTotal();
        }

        private void UpdateTotal()
        {
            double total = Operations.Sum(op => op.Cost);
            TotalText.Text = $"Операции: {total:N2} ₽";
            EmptyNote.Visibility = Operations.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            var menu = AddButton.ContextMenu;
            menu.PlacementTarget = AddButton;
            menu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
            menu.IsOpen = true;
        }

        private void AddOperation_Click(object sender, RoutedEventArgs e)
        {
            var tag = (sender as MenuItem)?.Tag as string;
            if (!Enum.TryParse(tag, out ManufacturingOperationType type))
                return;

            ManufacturingOperationBase op;
            switch (type)
            {
                case ManufacturingOperationType.LaserCutting: op = new LaserCuttingOperation(); break;
                case ManufacturingOperationType.Bending: op = new BendingOperation(); break;
                case ManufacturingOperationType.Rolling: op = new RollingOperation(); break;
                case ManufacturingOperationType.Flanging: op = new FlangingOperation(); break;
                default: op = new CustomOperation { Name = "Сварка" }; break;
            }

            op.Origin = OperationOrigin.Manual;
            AddWorkingCopy(op);
        }

        private void RemoveOperation_Click(object sender, RoutedEventArgs e)
        {
            var op = (sender as FrameworkElement)?.Tag as ManufacturingOperationBase;
            if (op == null)
                return;

            if (op.IsManual)
            {
                op.PropertyChanged -= Operation_PropertyChanged;
                Operations.Remove(op);
                UpdateTotal();
            }
            else
            {
                // Операцию из КОМПАС не удаляем, а исключаем: иначе она вернётся при обновлении из КОМПАС
                op.IsExcluded = !op.IsExcluded;
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var op in Operations)
                op.PropertyChanged -= Operation_PropertyChanged;
            Operations.Clear();

            // Операции из КОМПАС без правок. Исходной геометрии отредактированных операций здесь уже
            // нет — с флагом IsEdited = false её вернёт следующее «Обновить из КОМПАС»
            foreach (var op in _original.Where(o => !o.IsManual))
            {
                var copy = op.Clone();
                copy.IsEdited = false;
                copy.IsExcluded = false;
                copy.HasManualCost = false;
                copy.ManualCost = 0;
                copy.Quantity = 1;
                AddWorkingCopy(copy);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            NumberInput.CommitFocusedTextBox();
            if (NumberInput.HasValidationErrors(this))
            {
                NumberInput.ShowValidationWarning(this);
                return;
            }

            foreach (var op in Operations)
            {
                op.PropertyChanged -= Operation_PropertyChanged;
                if (op is CustomOperation custom && string.IsNullOrWhiteSpace(custom.Name))
                    custom.Name = "Прочая операция";
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
