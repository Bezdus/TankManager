using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using TankManager.Core.Models;
using TankManager.Core.ViewModels;

namespace TankManager
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            this.DataContext = _viewModel;
        }

        private void FileDropZone_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string filePath = files[0];
                    if (IsKompasFile(filePath))
                    {
                        _viewModel.FilePath = filePath;
                    }
                    else
                    {
                        MessageBox.Show("Пожалуйста, выберите файл КОМПАС (.a3d)", 
                            "Неверный формат файла", 
                            MessageBoxButton.OK, 
                            MessageBoxImage.Warning);
                    }
                }
            }
        }

        private void FileDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void ClearFilter_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.SelectedMaterial = null;
        }

        private void MaterialsListBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Получаем элемент под курсором
            var listBox = sender as ListBox;
            if (listBox == null)
                return;

            // Проверяем, кликнули ли мы на ListBoxItem
            var clickedElement = e.OriginalSource as DependencyObject;
            var listBoxItem = FindParent<ListBoxItem>(clickedElement);

            if (listBoxItem != null)
            {
                // Клик по элементу списка
                if (listBoxItem.Content is MaterialInfo clickedMaterial)
                {
                    // Если кликнули по уже выбранному материалу - сбрасываем фильтр
                    if (_viewModel.SelectedMaterial != null && 
                        clickedMaterial.Name == _viewModel.SelectedMaterial.Name)
                    {
                        _viewModel.SelectedMaterial = null;
                        e.Handled = true;
                    }
                }
            }
            else
            {
                // Клик в пустую область - сбрасываем фильтр
                _viewModel.SelectedMaterial = null;
            }
        }

        // Вспомогательный метод для поиска родительского элемента определенного типа
        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            while (child != null)
            {
                if (child is T parent)
                    return parent;
                
                child = System.Windows.Media.VisualTreeHelper.GetParent(child);
            }
            return null;
        }

        private bool IsKompasFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;
            
            return Path.GetExtension(filePath).Equals(".a3d", StringComparison.OrdinalIgnoreCase);
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _viewModel?.Dispose();
        }
    }

    /// <summary>
    /// Конвертер для проверки количества элементов в группе
    /// </summary>
    public class GroupCountToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int count)
            {
                // Если параметр "Multiple" - показываем только для count > 1
                // Если параметр "Single" - показываем только для count == 1
                string mode = parameter as string;
                
                if (mode == "Multiple")
                    return count > 1 ? Visibility.Visible : Visibility.Collapsed;
                else if (mode == "Single")
                    return count == 1 ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Статический класс команд для работы с Expander
    /// </summary>
    public static class ExpanderCommands
    {
        public static ICommand ToggleCommand { get; } = new RelayCommand<System.Windows.Controls.Expander>(expander =>
        {
            if (expander != null)
            {
                expander.IsExpanded = !expander.IsExpanded;
            }
        });
    }

    /// <summary>
    /// Простая реализация RelayCommand для использования в коде
    /// </summary>
    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, bool> _canExecute;

        public RelayCommand(Action<T> execute, Func<T, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute((T)parameter);
        }

        public void Execute(object parameter)
        {
            _execute((T)parameter);
        }
    }
}
