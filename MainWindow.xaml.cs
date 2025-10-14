using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
            InitializeComponent();
            _viewModel = new MainViewModel();
            this.DataContext = _viewModel;
        }

        private void FileDropZone_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                var kompasFile = files?.FirstOrDefault();

                if (IsKompasFile(kompasFile))
                {
                    var fileName = Path.GetFullPath(kompasFile);
                    _viewModel.FilePath = fileName;
                }
            }
            e.Handled = true;
        }

        private void FileDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                e.Effects = IsKompasFile(files?.FirstOrDefault()) 
                    ? DragDropEffects.Copy 
                    : DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void ClearFilter_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SelectedMaterial = null;
            }
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
                var clickedMaterial = listBoxItem.Content as string;
                
                // Если кликнули по уже выбранному материалу - сбрасываем фильтр
                if (clickedMaterial != null && clickedMaterial == _viewModel.SelectedMaterial)
                {
                    _viewModel.SelectedMaterial = null;
                    e.Handled = true;
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
}
