using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using CommunityToolkit.Mvvm.Input;
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
        public ICollectionView DetailsView { get; }
        public ICollectionView StandardPartsView { get; }

        public ICommand ShowInKompasCommand { get; }

        public MainViewModel() : this(new KompasService())
        {
        }

        public MainViewModel(IKompasService kompasService)
        {
            _kompasService = kompasService ?? throw new ArgumentNullException(nameof(kompasService));

            Details = new ObservableCollection<PartModel>();
            Materials = new ObservableCollection<string>();
            StandardParts = new ObservableCollection<PartModel>();

            // Используем обычный CollectionViewSource с пользовательской группировкой
            DetailsView = CollectionViewSource.GetDefaultView(Details);
            DetailsView.Filter = FilterDetails;
            DetailsView.GroupDescriptions.Add(new PartNameAndMarkingGroupDescription());

            StandardPartsView = CollectionViewSource.GetDefaultView(StandardParts);
            StandardPartsView.Filter = FilterDetails;
            StandardPartsView.GroupDescriptions.Add(new PartNameAndMarkingGroupDescription());

            Details.CollectionChanged += (s, e) => UpdateCalculations();
            StandardParts.CollectionChanged += (s, e) => UpdateCalculations();

            ShowInKompasCommand = new RelayCommand(ShowDetailInKompas, () => CurrentlySelectedPart != null);
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
                    _ = LoadDocumentAsync(value);
                }
            }
        }

        private string _selectedMaterial;
        public string SelectedMaterial
        {
            get { return _selectedMaterial; }
            set
            {
                if (_selectedMaterial != value)
                {
                    _selectedMaterial = value;
                    OnPropertyChanged(nameof(SelectedMaterial));
                    DetailsView?.Refresh();
                    StandardPartsView?.Refresh();
                    UpdateCalculations();
                }
            }
        }

        private PartModel _currentlySelectedPart;
        public PartModel CurrentlySelectedPart
        {
            get { return _currentlySelectedPart; }
            private set
            {
                if (_currentlySelectedPart != value)
                {
                    _currentlySelectedPart = value;
                    OnPropertyChanged(nameof(CurrentlySelectedPart));
                    ((RelayCommand)ShowInKompasCommand).NotifyCanExecuteChanged();
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
                    if (value != null)
                    {
                        SelectedStandardPart = null;
                        CurrentlySelectedPart = value;
                    }
                }
            }
        }

        private PartModel _selectedStandardPart;
        public PartModel SelectedStandardPart
        {
            get { return _selectedStandardPart; }
            set
            {
                if (_selectedStandardPart != value)
                {
                    _selectedStandardPart = value;
                    OnPropertyChanged(nameof(SelectedStandardPart));
                    if (value != null)
                    {
                        SelectedDetail = null;
                        CurrentlySelectedPart = value;
                    }
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

        private double _totalMassMultipleParts;
        public double TotalMassMultipleParts
        {
            get { return _totalMassMultipleParts; }
            set
            {
                if (Math.Abs(_totalMassMultipleParts - value) > 0.0001)
                {
                    _totalMassMultipleParts = value;
                    OnPropertyChanged(nameof(TotalMassMultipleParts));
                }
            }
        }

        private int _uniquePartsCount;
        public int UniquePartsCount
        {
            get { return _uniquePartsCount; }
            set
            {
                if (_uniquePartsCount != value)
                {
                    _uniquePartsCount = value;
                    OnPropertyChanged(nameof(UniquePartsCount));
                }
            }
        }

        private bool FilterDetails(object obj)
        {
            if (string.IsNullOrEmpty(_selectedMaterial))
                return true;

            if (obj is PartModel part)
            {
                return part.Material == _selectedMaterial;
            }

            return false;
        }

        private void UpdateCalculations()
        {
            var visibleParts = DetailsView.Cast<PartModel>().ToList();

            var groupedParts = visibleParts
                .GroupBy(p => new { p.Name, p.Marking })
                .Where(g => g.Count() > 1)
                .ToList();

            TotalMassMultipleParts = groupedParts
                .Sum(g => g.Sum(p => p.Mass));

            UniquePartsCount = groupedParts.Count;
        }

        private void ShowDetailInKompas()
        {
            if (CurrentlySelectedPart != null)
            {
                try
                {
                    _kompasService.ShowDetailInKompas(CurrentlySelectedPart);
                    StatusMessage = $"Показана деталь: {CurrentlySelectedPart.Name}";
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Ошибка: {ex.Message}";
                    MessageBox.Show($"Не удалось показать деталь в KOMPAS: {ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async Task LoadDocumentAsync(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return;

            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа...";

                Details.Clear();
                Materials.Clear();
                StandardParts.Clear();

                var parts = await Task.Run(() => _kompasService.LoadDocument(filePath));

                foreach (var part in parts)
                {
                    Details.Add(part);

                    if (!string.IsNullOrEmpty(part.Material) && !Materials.Contains(part.Material))
                    {
                        Materials.Add(part.Material);
                    }

                    if (part.DetailType == "Покупная деталь")
                    {
                        StandardParts.Add(part);
                    }
                }

                UpdateCalculations();
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

        // Кастомная группировка по имени и маркировке
        private class PartNameAndMarkingGroupDescription : GroupDescription
        {
            public override object GroupNameFromItem(object item, int level, CultureInfo culture)
            {
                if (item is PartModel part)
                {
                    // Группируем по комбинации имени и маркировки
                    var name = part.Name ?? string.Empty;
                    var marking = part.Marking ?? string.Empty;
                    return $"{name} | {marking}";
                }
                return string.Empty;
            }

            public override bool NamesMatch(object groupName, object itemName)
            {
                return object.Equals(groupName, itemName);
            }
        }
    }
}