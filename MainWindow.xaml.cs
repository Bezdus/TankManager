using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using TankManager.Core.Models;
using TankManager.Core.Services;
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

            // Устанавливаем размер окна - 60% экрана
            this.Width = SystemParameters.PrimaryScreenWidth * 0.6;
            this.Height = SystemParameters.PrimaryScreenHeight * 0.6;

            InitializeComponent();
            
            // Подписываемся на изменение видимости SnackBar для запуска анимации
            _viewModel.PropertyChanged += ViewModel_PropertyChanged;
        }

        private void ViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MainViewModel.IsSnackbarVisible))
            {
                if (_viewModel.IsSnackbarVisible)
                {
                    var showStoryboard = (Storyboard)FindResource("ShowSnackbarStoryboard");
                    showStoryboard?.Begin();
                }
                else
                {
                    var hideStoryboard = (Storyboard)FindResource("HideSnackbarStoryboard");
                    hideStoryboard?.Begin();
                }
            }
            else if (e.PropertyName == nameof(MainViewModel.IsProductsPanelOpen))
            {
                if (_viewModel.IsProductsPanelOpen)
                {
                    ProductsPanelOverlay.Visibility = Visibility.Visible;
                    ProductsPanel.Visibility = Visibility.Visible;
                    var showStoryboard = (Storyboard)FindResource("ShowProductsPanelStoryboard");
                    showStoryboard?.Begin();
                }
                else
                {
                    var hideStoryboard = (Storyboard)FindResource("HideProductsPanelStoryboard");
                    if (hideStoryboard != null)
                    {
                        var storyboardCopy = hideStoryboard.Clone();
                        storyboardCopy.Completed += (s, args) =>
                        {
                            ProductsPanelOverlay.Visibility = Visibility.Collapsed;
                            ProductsPanel.Visibility = Visibility.Collapsed;
                        };
                        storyboardCopy.Begin(this);
                    }
                }
            }
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

        private void ProductHeader_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _viewModel.IsProductSelected = true;
            e.Handled = true;
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
            
            // Если это не Visual (например, Run), пытаемся получить родительский Visual
            if (clickedElement != null && !(clickedElement is Visual || clickedElement is System.Windows.Media.Media3D.Visual3D))
            {
                // Для элементов типа Run, получаем Parent через логическое дерево
                if (clickedElement is FrameworkContentElement fce)
                {
                    clickedElement = fce.Parent as DependencyObject;
                }
            }
            
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
            // Защита от null
            if (child == null)
                return null;
                
            while (child != null)
            {
                if (child is T parent)
                    return parent;
                
                // Проверяем, является ли объект Visual перед попыткой получить родителя
                if (child is Visual || child is System.Windows.Media.Media3D.Visual3D)
                {
                    child = System.Windows.Media.VisualTreeHelper.GetParent(child);
                }
                else
                {
                    // Если не Visual, прерываем поиск
                    break;
                }
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

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            base.OnPreviewKeyDown(e);
            if (e.Key == Key.Escape && DrawingPopupOverlay.Visibility == Visibility.Visible)
            {
                CloseDrawingPopup();
                e.Handled = true;
            }
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

            DrawingPopupTitle.Text = $"Чертёж: {part.Name} {part.Marking}";
            DrawingPopupImage.Source = part.DrawingPreview;
            DrawingPopupOverlay.Visibility = Visibility.Visible;
        }

        private void CloseDrawingPopup()
        {
            DrawingPopupOverlay.Visibility = Visibility.Collapsed;
            DrawingPopupImage.Source = null;
        }

        private void DrawingPopupOverlay_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            CloseDrawingPopup();
        }

        private void DrawingPopupContent_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
        }

        private void DrawingPopupClose_Click(object sender, RoutedEventArgs e)
        {
            CloseDrawingPopup();
        }

        private void DeleteProduct_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true; // Предотвращаем всплытие события к ListBoxItem
            
            var button = sender as Button;
            if (button?.DataContext != null)
            {
                _viewModel.DeleteProductCommand?.Execute(button.DataContext);
            }
        }

        private void DeleteProductLocal_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;

            var button = sender as Button;
            if (button == null)
                return;

            var listBoxItem = FindParent<ListBoxItem>(button);
            var productInfo = (listBoxItem?.DataContext ?? button.DataContext) as ProductFileInfo;

            if (productInfo != null)
            {
                _viewModel.SelectedSavedProduct = productInfo;
                _viewModel.DeleteProductLocalCommand?.Execute(null);
            }
        }

        private void DeleteProductEverywhere_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;

            var button = sender as Button;
            if (button == null)
                return;

            var listBoxItem = FindParent<ListBoxItem>(button);
            var productInfo = (listBoxItem?.DataContext ?? button.DataContext) as ProductFileInfo;

            if (productInfo != null)
            {
                _viewModel.SelectedSavedProduct = productInfo;
                _viewModel.DeleteProductEverywhereCommand?.Execute(null);
            }
        }

        private void DeleteButton_MouseEnter(object sender, MouseEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                button.Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(255, 224, 224)); // #FFFFE0E0
                button.BorderBrush = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(255, 107, 107)); // #FFFF6B6B
            }
        }

        private void DeleteButton_MouseLeave(object sender, MouseEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                button.Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(255, 255, 255)); // #FFFFFFFF (PrimaryBackground)
                button.BorderBrush = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(224, 224, 224)); // #FFE0E0E0 (BorderColor)
            }
        }

        private void DeleteButtonEverywhere_MouseEnter(object sender, MouseEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                button.Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(255, 200, 200)); // более насыщенный красный
                button.BorderBrush = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(220, 60, 60));
            }
        }

        private void DeleteButtonEverywhere_MouseLeave(object sender, MouseEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                button.Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(255, 255, 255));
                button.BorderBrush = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(224, 224, 224));
            }
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
