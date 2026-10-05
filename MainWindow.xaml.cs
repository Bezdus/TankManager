using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
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
        private readonly WindowSettingsService _windowSettingsService = new WindowSettingsService(new FileLogger());
        private double[] _savedColumnStars;

        public MainWindow()
        {
            _viewModel = new MainViewModel();
            this.DataContext = _viewModel;

            InitializeComponent();

            RestoreWindowSettings();
            UpdateTitle();

            // Подписываемся на изменение видимости SnackBar для запуска анимации
            _viewModel.PropertyChanged += ViewModel_PropertyChanged;

            Loaded += (s, e) => _viewModel.OnWindowLoaded();
        }

        #region Заголовок и настройки окна

        private void UpdateTitle()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            var title = $"Tank Manager v{version.Major}.{version.Minor}.{version.Build}";

            var productName = _viewModel.CurrentProduct?.Name;
            if (!string.IsNullOrWhiteSpace(productName))
                title += $" — {productName}";

            Title = title;
        }

        private void RestoreWindowSettings()
        {
            var settings = _windowSettingsService.Load();
            if (settings == null)
            {
                // Первый запуск: 60% экрана по центру
                Width = Math.Max(MinWidth, SystemParameters.PrimaryScreenWidth * 0.6);
                Height = Math.Max(MinHeight, SystemParameters.PrimaryScreenHeight * 0.6);
                return;
            }

            if (settings.Width >= MinWidth && settings.Height >= MinHeight && IsOnScreen(settings))
            {
                WindowStartupLocation = WindowStartupLocation.Manual;
                Left = settings.Left;
                Top = settings.Top;
                Width = settings.Width;
                Height = settings.Height;
            }

            if (settings.IsMaximized)
                WindowState = WindowState.Maximized;

            var stars = settings.ColumnStars;
            if (stars != null && stars.Length == 3 && Array.TrueForAll(stars, s => s > 0))
            {
                _savedColumnStars = stars;
                LeftColumn.Width = new GridLength(stars[0], GridUnitType.Star);
                CenterColumn.Width = new GridLength(stars[1], GridUnitType.Star);
                RightColumn.Width = new GridLength(stars[2], GridUnitType.Star);
            }
        }

        /// <summary>
        /// Проверяет, что заголовок окна попадает в видимую область рабочего стола
        /// </summary>
        private static bool IsOnScreen(WindowSettings settings)
        {
            var left = SystemParameters.VirtualScreenLeft;
            var top = SystemParameters.VirtualScreenTop;
            var right = left + SystemParameters.VirtualScreenWidth;
            var bottom = top + SystemParameters.VirtualScreenHeight;

            return settings.Left + 100 <= right
                && settings.Left + settings.Width - 100 >= left
                && settings.Top >= top
                && settings.Top + 30 <= bottom;
        }

        private void SaveWindowSettings()
        {
            var bounds = WindowState == WindowState.Normal
                ? new Rect(Left, Top, ActualWidth, ActualHeight)
                : RestoreBounds;

            // Колонки имеют нулевую ширину, пока изделие не загружено — тогда оставляем прежние пропорции
            var stars = new[] { LeftColumn.ActualWidth, CenterColumn.ActualWidth, RightColumn.ActualWidth };
            if (!Array.TrueForAll(stars, s => s > 0))
                stars = _savedColumnStars;

            _windowSettingsService.Save(new WindowSettings
            {
                Left = bounds.Left,
                Top = bounds.Top,
                Width = bounds.Width,
                Height = bounds.Height,
                IsMaximized = WindowState == WindowState.Maximized,
                ColumnStars = stars
            });
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            if (!e.Cancel)
                SaveWindowSettings();
        }

        #endregion

        private void ViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MainViewModel.CurrentProduct))
            {
                UpdateTitle();
            }
            else if (e.PropertyName == nameof(MainViewModel.IsSnackbarVisible))
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
            DropOverlay.Visibility = Visibility.Collapsed;

            var filePath = GetDraggedFile(e);
            if (filePath == null)
                return;

            if (_viewModel.IsViewerMode)
            {
                _viewModel.ShowSnackbar(MainViewModel.ViewerModeLoadMessage, SnackbarKind.Warning);
            }
            else if (IsKompasFile(filePath))
            {
                _viewModel.FilePath = filePath;
            }
            else
            {
                _viewModel.ShowSnackbar("Неверный формат файла: нужна сборка КОМПАС (.a3d)", SnackbarKind.Error);
            }
        }

        private void FileDropZone_DragEnter(object sender, DragEventArgs e)
        {
            var filePath = GetDraggedFile(e);
            if (filePath == null)
                return;

            DropOverlayText.Text = _viewModel.IsViewerMode
                ? MainViewModel.ViewerModeLoadMessage
                : IsKompasFile(filePath)
                    ? "Отпустите, чтобы загрузить сборку"
                    : "Поддерживаются только файлы КОМПАС (.a3d)";
            DropOverlay.Visibility = Visibility.Visible;
        }

        private void FileDropZone_DragOver(object sender, DragEventArgs e)
        {
            var filePath = GetDraggedFile(e);
            e.Effects = filePath != null && IsKompasFile(filePath) && !_viewModel.IsViewerMode
                ? DragDropEffects.Copy
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void DropOverlay_DragLeave(object sender, DragEventArgs e)
        {
            DropOverlay.Visibility = Visibility.Collapsed;
        }

        private static string GetDraggedFile(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return null;

            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            return files != null && files.Length > 0 ? files[0] : null;
        }

        private void SnackbarClose_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.IsSnackbarVisible = false;
        }

        private void OtherMaterialsListBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            HandleMaterialListBoxClick(sender, e,
                () => _viewModel.SelectedOtherMaterial,
                () => _viewModel.SelectedOtherMaterial = null);
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
            if (e.Handled)
                return;

            var ctrl = Keyboard.Modifiers == ModifierKeys.Control;

            if (ctrl && e.Key == Key.O)
            {
                // В режиме просмотра «открыть» — это список изделий
                if (_viewModel.IsViewerMode)
                    _viewModel.IsProductsPanelOpen = true;
                else
                    LoadOptionsPopup.IsOpen = true;
                e.Handled = true;
            }
            else if (ctrl && e.Key == Key.F && _viewModel.HasProduct)
            {
                SearchBox.Focus();
                SearchBox.SelectAll();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                if (DrawingPopupOverlay.Visibility == Visibility.Visible)
                {
                    CloseDrawingPopup();
                    e.Handled = true;
                }
                else if (_viewModel.IsProductsPanelOpen)
                {
                    _viewModel.IsProductsPanelOpen = false;
                    e.Handled = true;
                }
                else if (SearchBox.IsKeyboardFocusWithin && !string.IsNullOrEmpty(_viewModel.SearchText))
                {
                    _viewModel.SearchText = string.Empty;
                    e.Handled = true;
                }
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
