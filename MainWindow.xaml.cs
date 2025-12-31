using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using TankManager.Core.Models;
using TankManager.Core.ViewModels;

namespace TankManager
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;

        public MainWindow()
        {
            _viewModel = new MainViewModel();
            this.DataContext = _viewModel;
            
            InitializeComponent();
        }

        private void LoadDocumentButton_Click(object sender, RoutedEventArgs e)
        {
            LoadOptionsPopup.IsOpen = true;
        }

        private async void LoadFromKompas_Click(object sender, RoutedEventArgs e)
        {
            LoadOptionsPopup.IsOpen = false;
            await _viewModel.LoadFromActiveDocumentAsync();
        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            LoadOptionsPopup.IsOpen = false;

            var openFileDialog = new OpenFileDialog
            {
                Filter = "Файлы КОМПАС (*.a3d)|*.a3d|Все файлы (*.*)|*.*",
                Title = "Выберите файл КОМПАС"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _viewModel.FilePath = openFileDialog.FileName;
            }
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
                        MessageBox.Show(
                            "Пожалуйста, выберите файл КОМПАС (.a3d)", 
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
            _viewModel.ClearMaterialFilter();
        }

        private void SheetMaterialsListBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            HandleMaterialListBoxClick(sender, e, 
                () => _viewModel.SelectedSheetMaterial, 
                () => _viewModel.SelectedSheetMaterial = null);
        }

        private void TubularProductsListBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            HandleMaterialListBoxClick(sender, e, 
                () => _viewModel.SelectedTubularProduct, 
                () => _viewModel.SelectedTubularProduct = null);
        }

        private void HandleMaterialListBoxClick(object sender, MouseButtonEventArgs e, 
            Func<MaterialInfo> getSelected, Action clearSelection)
        {
            var listBox = sender as ListBox;
            if (listBox == null)
                return;

            var clickedElement = e.OriginalSource as DependencyObject;
            var listBoxItem = FindParent<ListBoxItem>(clickedElement);

            if (listBoxItem != null)
            {
                if (listBoxItem.Content is MaterialInfo clickedMaterial)
                {
                    var selected = getSelected();
                    // Сброс фильтра при клике по уже выбранному материалу
                    if (selected != null && clickedMaterial.Name == selected.Name)
                    {
                        clearSelection();
                        e.Handled = true;
                    }
                }
            }
            else
            {
                // Сброс фильтра при клике в пустую область
                clearSelection();
            }
        }

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

        private void OverlayBackground_MouseDown(object sender, MouseButtonEventArgs e)
        {
            _viewModel.IsProductsPanelOpen = false;
        }

        private void DrawingPreview_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var part = _viewModel.CurrentlySelectedPart;
            if (part?.DrawingPreview == null)
                return;

            var previewWindow = new Window
            {
                Title = $"Чертёж: {part.Name} {part.Marking}",
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Width = SystemParameters.WorkArea.Width * 0.9,
                Height = SystemParameters.WorkArea.Height * 0.9,
                Background = System.Windows.Media.Brushes.White,
                WindowState = WindowState.Maximized
            };

            var scrollViewer = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Background = System.Windows.Media.Brushes.LightGray
            };

            var image = new Image
            {
                Source = part.DrawingPreview,
                Stretch = System.Windows.Media.Stretch.None,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(20)
            };

            // Закрытие по Escape
            previewWindow.KeyDown += (s, args) =>
            {
                if (args.Key == Key.Escape)
                    previewWindow.Close();
            };

            // Двойной клик для закрытия
            image.MouseLeftButtonDown += (s, args) =>
            {
                if (args.ClickCount == 2)
                    previewWindow.Close();
            };

            scrollViewer.Content = image;
            previewWindow.Content = scrollViewer;
            
            // Очистка ресурсов при закрытии окна
            previewWindow.Closed += (s, args) =>
            {
                image.Source = null;
                scrollViewer.Content = null;
                previewWindow.Content = null;
            };
            
            previewWindow.ShowDialog();
        }
    }

    /// <summary>
    /// Статический класс команд для работы с Expander
    /// </summary>
    public static class ExpanderCommands
    {
        public static ICommand ToggleCommand { get; } = new RelayCommand<Expander>(expander =>
        {
            if (expander != null)
            {
                expander.IsExpanded = !expander.IsExpanded;
            }
        });
    }

    /// <summary>
    /// Простая реализация RelayCommand
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
