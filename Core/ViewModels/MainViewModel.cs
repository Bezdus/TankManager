using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using TankManager.Core.Models;
using TankManager.Core.Services;

namespace TankManager.Core.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged, IDisposable
    {
        private readonly IKompasService _kompasService;
        
        public ObservableCollection<PartModel> Details { get; }
        public ObservableCollection<string> Materials { get; }
        public ObservableCollection<PartModel> StandardParts { get; }

        public MainViewModel() : this(new KompasService())
        {
        }

        public MainViewModel(IKompasService kompasService)
        {
            _kompasService = kompasService ?? throw new ArgumentNullException(nameof(kompasService));
            Details = new ObservableCollection<PartModel>();
            Materials = new ObservableCollection<string>();
            StandardParts = new ObservableCollection<PartModel>();
        }

        private string _filePath;
        public string FilePath
        {
            get { return _filePath; }
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    OnPropertyChanged(nameof(FilePath));
                    LoadDocument();
                }
            }
        }



        private PartModel _selectedDetail;
        public PartModel SelectedDetail
        {
            get { return _selectedDetail; }
            set
            {
                if (_selectedDetail != value)
                {
                    _selectedDetail = value;
                    OnPropertyChanged(nameof(SelectedDetail));
                    _kompasService.ShowDetailInKompas(_selectedDetail);
                }
            }
        }

        private bool _isLoading;
        public bool IsLoading
        {
            get { return _isLoading; }
            set
            {
                if (_isLoading != value)
                {
                    _isLoading = value;
                    OnPropertyChanged(nameof(IsLoading));
                }
            }
        }

        private string _statusMessage;
        public string StatusMessage
        {
            get { return _statusMessage; }
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged(nameof(StatusMessage));
                }
            }
        }

        private void LoadDocument()
        {
            if (string.IsNullOrEmpty(_filePath))
                return;

            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа...";
                Details.Clear();
                Materials.Clear();
                StandardParts.Clear();

                List<PartModel> parts = _kompasService.LoadDocument(_filePath);
                
                foreach (PartModel part in parts)
                {
                    Details.Add(part);
                    
                    // Собираем уникальные материалы
                    if (!string.IsNullOrEmpty(part.Material) && !Materials.Contains(part.Material))
                    {
                        Materials.Add(part.Material);
                    }
                    
                    // Фильтруем стандартные детали (например, по типу)
                    if (part.DetailType?.Contains("Стандартное") == true || 
                        part.DetailType?.Contains("Standard") == true)
                    {
                        StandardParts.Add(part);
                    }
                }

                StatusMessage = $"Загружено деталей: {Details.Count}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
                MessageBox.Show($"Ошибка при загрузке файла: {ex.Message}",
                              "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void Dispose()
        {
            _kompasService?.Dispose();
        }
    }
}
