using Microsoft.Win32;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace TankManager.Views
{
    /// <summary>
    /// Логика взаимодействия для LoadingScreen.xaml
    /// </summary>
    public partial class LoadingScreen : Window
    {
        public string SelectedFilePath { get; private set; }
        public bool LoadFromActiveDocument { get; private set; }

        public LoadingScreen()
        {
            InitializeComponent();
        }

        private async void LoadFromKompas_Click(object sender, MouseButtonEventArgs e)
        {
            try
            {
                StatusText.Text = "Загрузка документа из КОМПАС...";
                LoadingOverlay.Visibility = Visibility.Visible;

                // Проверяем наличие активного документа
                await Task.Run(() =>
                {
                    // Небольшая задержка для визуального эффекта
                    System.Threading.Thread.Sleep(300);
                });

                LoadFromActiveDocument = true;
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
                StatusText.Text = "Выберите способ загрузки или перетащите файл .a3d в окно";
                MessageBox.Show($"Не удалось загрузить документ из КОМПАС: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SelectFile_Click(object sender, MouseButtonEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Выберите файл КОМПАС",
                Filter = "Файлы КОМПАС (*.a3d)|*.a3d|Все файлы (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                SelectedFilePath = dialog.FileName;
                LoadFromActiveDocument = false;
                StatusText.Text = $"Выбран файл: {Path.GetFileName(SelectedFilePath)}";
                DialogResult = true;
                Close();
            }
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            // Скрываем оверлей
            DragDropOverlay.Visibility = Visibility.Collapsed;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string filePath = files[0];
                    if (IsKompasFile(filePath))
                    {
                        SelectedFilePath = filePath;
                        LoadFromActiveDocument = false;
                        StatusText.Text = $"Загружен файл: {Path.GetFileName(SelectedFilePath)}";
                        DialogResult = true;
                        Close();
                    }
                    else
                    {
                        MessageBox.Show("Пожалуйста, выберите файл КОМПАС (.a3d)",
                            "Неверный формат файла",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        StatusText.Text = "Ошибка: неверный формат файла";
                    }
                }
            }
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                
                // Показываем оверлей на всё окно
                DragDropOverlay.Visibility = Visibility.Visible;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void Window_DragLeave(object sender, DragEventArgs e)
        {
            // Скрываем оверлей когда курсор покидает окно
            DragDropOverlay.Visibility = Visibility.Collapsed;
        }

        private bool IsKompasFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            return Path.GetExtension(filePath).Equals(".a3d", StringComparison.OrdinalIgnoreCase);
        }
    }
}
