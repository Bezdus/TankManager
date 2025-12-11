using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
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
        private readonly ProductStorageService _storageService = new ProductStorageService();
        private bool _isUpdatingCalculations = false;

        private Product _currentProduct;
        public Product CurrentProduct
        {
            get => _currentProduct;
            private set
            {
                if (_currentProduct != value)
                {
                    _currentProduct = value;
                    OnPropertyChanged(nameof(CurrentProduct));
                    OnPropertyChanged(nameof(Details));
                    OnPropertyChanged(nameof(Materials));
                    OnPropertyChanged(nameof(StandardParts));
                    
                    // Пересоздаём представления для новых коллекций
                    InitializeCollectionViews();
                    
                    // Автосохранение при установке нового продукта
                    AutoSaveProduct();
                }
            }
        }

        // Делегируем коллекции к Product
        public ObservableCollection<PartModel> Details => CurrentProduct?.Details;
        public ObservableCollection<MaterialInfo> Materials => CurrentProduct?.Materials;
        public ObservableCollection<PartModel> StandardParts => CurrentProduct?.StandardParts;

        public ICollectionView DetailsView { get; private set; }
        public ICollectionView StandardPartsView { get; private set; }
        public ICollectionView MaterialsView { get; private set; }

        // Список сохранённых продуктов
        private ObservableCollection<ProductFileInfo> _savedProducts;
        public ObservableCollection<ProductFileInfo> SavedProducts
        {
            get => _savedProducts;
            private set
            {
                _savedProducts = value;
                OnPropertyChanged(nameof(SavedProducts));
            }
        }

        // Выбранный продукт в списке
        private ProductFileInfo _selectedSavedProduct;
        public ProductFileInfo SelectedSavedProduct
        {
            get => _selectedSavedProduct;
            set
            {
                if (_selectedSavedProduct != value)
                {
                    _selectedSavedProduct = value;
                    OnPropertyChanged(nameof(SelectedSavedProduct));
                    ((RelayCommand)DeleteProductCommand)?.NotifyCanExecuteChanged();
                }
            }
        }

        // Видимость панели продуктов
        private bool _isProductsPanelOpen;
        public bool IsProductsPanelOpen
        {
            get => _isProductsPanelOpen;
            set
            {
                if (_isProductsPanelOpen != value)
                {
                    _isProductsPanelOpen = value;
                    OnPropertyChanged(nameof(IsProductsPanelOpen));
                    
                    if (value)
                    {
                        RefreshSavedProducts();
                    }
                }
            }
        }

        public ICommand ShowInKompasCommand { get; }
        public ICommand ToggleMaterialSortCommand { get; }
        public ICommand LoadFromActiveDocumentCommand { get; }
        public ICommand ClearSearchCommand { get; }
        public ICommand LoadProductCommand { get; }
        public ICommand DeleteProductCommand { get; }
        public ICommand ToggleProductsPanelCommand { get; }
        public ICommand SwitchToProductCommand { get; }

        public MainViewModel() : this(new KompasService())
        {
        }

        public MainViewModel(IKompasService kompasService)
        {
            Debug.WriteLine("MainViewModel.Constructor: Начало");
            
            try
            {
                _kompasService = kompasService ?? throw new ArgumentNullException(nameof(kompasService));
                Debug.WriteLine("MainViewModel.Constructor: KompasService установлен");

                SavedProducts = new ObservableCollection<ProductFileInfo>();
                CurrentProduct = new Product();
                Debug.WriteLine("MainViewModel.Constructor: Product создан");

                ShowInKompasCommand = new RelayCommand(ShowDetailInKompas, () => CurrentlySelectedPart != null);
                ToggleMaterialSortCommand = new RelayCommand(ToggleMaterialSort);
                LoadFromActiveDocumentCommand = new RelayCommand(async () => await LoadFromActiveDocumentAsync());
                ClearSearchCommand = new RelayCommand(ClearSearch);
                LoadProductCommand = new RelayCommand<string>(LoadProduct);
                DeleteProductCommand = new RelayCommand(DeleteSelectedProduct, () => SelectedSavedProduct != null);
                ToggleProductsPanelCommand = new RelayCommand(() => IsProductsPanelOpen = !IsProductsPanelOpen);
                SwitchToProductCommand = new RelayCommand<ProductFileInfo>(SwitchToProduct);
                Debug.WriteLine("MainViewModel.Constructor: Команды созданы");
                
                // Загрузка последнего продукта при старте (без автосохранения)
                var lastProduct = _storageService.LoadLast();
                if (lastProduct != null && !string.IsNullOrEmpty(lastProduct.Name))
                {
                    _currentProduct = lastProduct; // Прямое присваивание без вызова сеттера
                    OnPropertyChanged(nameof(CurrentProduct));
                    OnPropertyChanged(nameof(Details));
                    OnPropertyChanged(nameof(Materials));
                    OnPropertyChanged(nameof(StandardParts));
                    InitializeCollectionViews();
                    UpdateCalculations();
                    Debug.WriteLine("MainViewModel.Constructor: Загружен последний Product");
                }

                Debug.WriteLine("MainViewModel.Constructor: Завершено успешно");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MainViewModel.Constructor: ИСКЛЮЧЕНИЕ - {ex.Message}");
                Debug.WriteLine($"MainViewModel.Constructor: StackTrace - {ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Автоматическое сохранение продукта в список (если ещё не существует)
        /// </summary>
        private void AutoSaveProduct()
        {
            if (CurrentProduct == null || string.IsNullOrEmpty(CurrentProduct.Name))
                return;

            // Проверяем, существует ли уже такой продукт
            var existingProducts = _storageService.GetSavedProducts();
            var existing = existingProducts.FirstOrDefault(p => 
                p.ProductName == CurrentProduct.Name && 
                p.Marking == CurrentProduct.Marking);

            if (existing == null)
            {
                // Сохраняем только если такого продукта ещё нет
                string filePath = _storageService.Save(CurrentProduct);
                StatusMessage = $"Автосохранено: {Path.GetFileName(filePath)}";
                RefreshSavedProducts();
            }
        }

        private void InitializeCollectionViews()
        {
            if (Details != null)
            {
                DetailsView = CollectionViewSource.GetDefaultView(Details);
                DetailsView.Filter = FilterDetails;
                DetailsView.GroupDescriptions.Clear();
                DetailsView.GroupDescriptions.Add(new PartNameAndMarkingGroupDescription());
                OnPropertyChanged(nameof(DetailsView));
            }

            if (StandardParts != null)
            {
                StandardPartsView = CollectionViewSource.GetDefaultView(StandardParts);
                StandardPartsView.Filter = FilterDetails;
                StandardPartsView.GroupDescriptions.Clear();
                StandardPartsView.GroupDescriptions.Add(new PartNameAndMarkingGroupDescription());
                OnPropertyChanged(nameof(StandardPartsView));
            }

            if (Materials != null)
            {
                MaterialsView = CollectionViewSource.GetDefaultView(Materials);
                MaterialsView.SortDescriptions.Clear();
                MaterialsView.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));
                OnPropertyChanged(nameof(MaterialsView));
            }
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

        private string _searchText;
        public string SearchText
        {
            get { return _searchText; }
            set
            {
                if (_searchText != value)
                {
                    _searchText = value;
                    OnPropertyChanged(nameof(SearchText));
                    
                    DetailsView?.Refresh();
                    StandardPartsView?.Refresh();
                    UpdateCalculations();
                }
            }
        }

        private MaterialInfo _selectedMaterial;
        public MaterialInfo SelectedMaterial
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

        private bool _sortMaterialsByMass = false;
        public bool SortMaterialsByMass
        {
            get { return _sortMaterialsByMass; }
            set
            {
                if (_sortMaterialsByMass != value)
                {
                    _sortMaterialsByMass = value;
                    OnPropertyChanged(nameof(SortMaterialsByMass));
                    OnPropertyChanged(nameof(MaterialSortText));
                    ApplyMaterialSort();
                }
            }
        }

        public string MaterialSortText
        {
            get { return _sortMaterialsByMass ? "Сортировка: по массе ↓" : "Сортировка: по названию ↑"; }
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
            if (obj is PartModel part)
            {
                if (_selectedMaterial != null && part.Material != _selectedMaterial.Name)
                {
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(_searchText))
                {
                    string searchLower = _searchText.ToLower();
                    
                    bool nameMatch = part.Name != null && part.Name.ToLower().Contains(searchLower);
                    bool markingMatch = part.Marking != null && part.Marking.ToLower().Contains(searchLower);
                    
                    if (!nameMatch && !markingMatch)
                    {
                        return false;
                    }
                }

                return true;
            }

            return false;
        }

        private void ToggleMaterialSort()
        {
            SortMaterialsByMass = !SortMaterialsByMass;
        }

        private void ApplyMaterialSort()
        {
            if (MaterialsView == null) return;

            MaterialsView.SortDescriptions.Clear();
            
            if (_sortMaterialsByMass)
            {
                MaterialsView.SortDescriptions.Add(new SortDescription("TotalMass", ListSortDirection.Descending));
            }
            else
            {
                MaterialsView.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));
            }
        }

        private void UpdateCalculations()
        {
            if (_isUpdatingCalculations || DetailsView == null)
                return;

            _isUpdatingCalculations = true;

            try
            {
                var visibleParts = DetailsView.Cast<PartModel>().ToList();

                var groupedParts = visibleParts
                    .GroupBy(p => new { p.Name, p.Marking })
                    .Where(g => g.Count() > 1)
                    .ToList();

                TotalMassMultipleParts = groupedParts.Sum(g => g.Sum(p => p.Mass));
                UniquePartsCount = groupedParts.Count;

                UpdateMaterialWeights();
            }
            finally
            {
                _isUpdatingCalculations = false;
            }
        }

        private void UpdateMaterialWeights()
        {
            if (Details == null || Materials == null) return;

            var materialWeights = Details
                .Where(p => !string.IsNullOrEmpty(p.Material))
                .GroupBy(p => p.Material)
                .ToDictionary(g => g.Key, g => g.Sum(p => p.Mass));

            var existingMaterials = new HashSet<string>(Materials.Select(m => m.Name));

            foreach (var material in Materials)
            {
                if (materialWeights.TryGetValue(material.Name, out double totalMass))
                {
                    material.TotalMass = totalMass;
                }
                else
                {
                    material.TotalMass = 0;
                }
            }

            foreach (var kvp in materialWeights)
            {
                if (!existingMaterials.Contains(kvp.Key))
                {
                    Materials.Add(new MaterialInfo
                    {
                        Name = kvp.Key,
                        TotalMass = kvp.Value
                    });
                }
            }

            for (int i = Materials.Count - 1; i >= 0; i--)
            {
                if (!materialWeights.ContainsKey(Materials[i].Name))
                {
                    Materials.RemoveAt(i);
                }
            }
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

                var product = await Task.Run(() => _kompasService.LoadDocument(filePath));
                CurrentProduct = product;

                UpdateCalculations();
                StatusMessage = $"Загружено изделие: {CurrentProduct.Name}, деталей: {Details.Count}";
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

        public async Task LoadFromActiveDocumentAsync()
        {
            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа из КОМПАС...";

                var product = await Task.Run(() => _kompasService.LoadActiveDocument());
                CurrentProduct = product;

                UpdateCalculations();
                StatusMessage = $"Загружено изделие: {CurrentProduct.Name}, деталей: {Details.Count}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
                MessageBox.Show($"Ошибка при загрузке из КОМПАС: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ClearSearch()
        {
            SearchText = string.Empty;
        }

        private void LoadProduct(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return;

            var product = _storageService.Load(fileName);
            if (product != null)
            {
                // Прямое присваивание чтобы избежать повторного автосохранения
                _currentProduct = product;
                OnPropertyChanged(nameof(CurrentProduct));
                OnPropertyChanged(nameof(Details));
                OnPropertyChanged(nameof(Materials));
                OnPropertyChanged(nameof(StandardParts));
                InitializeCollectionViews();
                UpdateCalculations();
                StatusMessage = $"Загружено: {product.Name}";
            }
        }

        private void SwitchToProduct(ProductFileInfo productInfo)
        {
            if (productInfo == null) return;

            var product = _storageService.Load(productInfo.FileName);
            if (product != null)
            {
                // Прямое присваивание чтобы избежать повторного автосохранения
                _currentProduct = product;
                OnPropertyChanged(nameof(CurrentProduct));
                OnPropertyChanged(nameof(Details));
                OnPropertyChanged(nameof(Materials));
                OnPropertyChanged(nameof(StandardParts));
                InitializeCollectionViews();
                UpdateCalculations();
                StatusMessage = $"Переключено на: {product.Name}";
                IsProductsPanelOpen = false;
            }
        }

        private void DeleteSelectedProduct()
        {
            if (SelectedSavedProduct == null) return;

            var result = MessageBox.Show(
                $"Удалить \"{SelectedSavedProduct.ProductName}\"?",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                if (_storageService.Delete(SelectedSavedProduct.FileName))
                {
                    StatusMessage = $"Удалено: {SelectedSavedProduct.ProductName}";
                    RefreshSavedProducts();
                }
            }
        }

        public void RefreshSavedProducts()
        {
            SavedProducts.Clear();
            foreach (var product in _storageService.GetSavedProducts())
            {
                SavedProducts.Add(product);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void Dispose()
        {
            _storageService.SaveAsLast(CurrentProduct);
            CurrentProduct?.Clear();
            _kompasService?.Dispose();
        }

        private class PartNameAndMarkingGroupDescription : GroupDescription
        {
            public override object GroupNameFromItem(object item, int level, CultureInfo culture)
            {
                if (item is PartModel part)
                {
                    return $"{part.Name}|{part.Marking}";
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