using System;
using System.IO;
using System.Linq;
using System.Windows;
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
