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

        // Кэш загруженных продуктов (ключ = FilePath)
        private readonly Dictionary<string, Product> _linkedProductsCache = new Dictionary<string, Product>(StringComparer.OrdinalIgnoreCase);

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

                    // Сброс выбранных элементов при смене продукта
                    SelectedDetail = null;
                    SelectedStandardPart = null;
                    CurrentlySelectedPart = null;

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

        // Флаг связи с KOMPAS
        private bool _isLinkedToKompas;
        public bool IsLinkedToKompas
        {
            get => _isLinkedToKompas;
            private set
            {
                if (_isLinkedToKompas != value)
                {
                    _isLinkedToKompas = value;
                    OnPropertyChanged(nameof(IsLinkedToKompas));
                    OnPropertyChanged(nameof(KompasLinkStatus));
                }
            }
        }

        public string KompasLinkStatus => IsLinkedToKompas 
            ? "🔗 Связан с KOMPАС" 
            : "⚠️ Нет связи с KOMPАС";

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

                ShowInKompasCommand = new RelayCommand(ShowDetailInKompas, () => CurrentlySelectedPart != null && IsLinkedToKompas);
                ToggleMaterialSortCommand = new RelayCommand(ToggleMaterialSort);
                LoadFromActiveDocumentCommand = new RelayCommand(async () => await LoadFromActiveDocumentAsync());
                ClearSearchCommand = new RelayCommand(ClearSearch);
                LoadProductCommand = new RelayCommand<string>(fileName => _ = LoadProductAsync(fileName));
                DeleteProductCommand = new RelayCommand(DeleteSelectedProduct, () => SelectedSavedProduct != null);
                ToggleProductsPanelCommand = new RelayCommand(() => IsProductsPanelOpen = !IsProductsPanelOpen);
                SwitchToProductCommand = new RelayCommand<ProductFileInfo>(info => _ = SwitchToProductAsync(info));
                Debug.WriteLine("MainViewModel.Constructor: Команды созданы");
                
                // Загрузка последнего продукта при старте с автоматическим связыванием
                _ = LoadLastProductAsync();

                Debug.WriteLine("MainViewModel.Constructor: ЗавершеноSuccessfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MainViewModel.Constructor: ИСКЛЮЧЕНИЕ - {ex.Message}");
                Debug.WriteLine($"MainViewModel.Constructor: StackTrace - {ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Загрузка последнего продукта при старте с автоматическим связыванием с KOMPАС
        /// </summary>
        private async Task LoadLastProductAsync()
        {
            var lastProduct = _storageService.LoadLast();
            if (lastProduct == null || string.IsNullOrEmpty(lastProduct.Name))
                return;

            await LoadAndLinkProductAsync(lastProduct, "Загружен последний Product");
        }

        /// <summary>
        /// Загрузить продукт и автоматически связать с KOMPАС (с кэшированием)
        /// </summary>
        private async Task LoadAndLinkProductAsync(Product savedProduct, string successMessage)
        {
            string filePath = savedProduct.FilePath;

            // Проверяем кэш — если уже загружали этот продукт, используем его
            if (!string.IsNullOrEmpty(filePath) && _linkedProductsCache.TryGetValue(filePath, out var cachedProduct))
            {
                // Мгновенное переключение из кэша
                SetCurrentProduct(cachedProduct, isLinked: true);
                StatusMessage = $"{successMessage} (из кэша)";
                Debug.WriteLine($"Продукт загружен из кэша: {filePath}");
                return;
            }

            // Сначала показываем сохранённые данные (мгновенно)
            SetCurrentProduct(savedProduct, isLinked: false);

            // Затем пытаемся связать с KOMPАС в фоне
            if (!string.IsNullOrEmpty(filePath))
            {
                await TryLinkToKompasAsync(filePath);
            }

            StatusMessage = IsLinkedToKompas 
                ? $"{successMessage} (связан с KOMPАС)" 
                : $"{successMessage} (без связи с KOMPАС)";
        }

        /// <summary>
        /// Установить текущий продукт без вызова сеттера (для избежания автосохранения при переключении)
        /// </summary>
        private void SetCurrentProduct(Product product, bool isLinked)
        {
            _currentProduct = product;
            _isLinkedToKompas = isLinked;

            SelectedDetail = null;
            SelectedStandardPart = null;
            CurrentlySelectedPart = null;

            OnPropertyChanged(nameof(CurrentProduct));
            OnPropertyChanged(nameof(Details));
            OnPropertyChanged(nameof(Materials));
            OnPropertyChanged(nameof(StandardParts));
            OnPropertyChanged(nameof(IsLinkedToKompas));
            OnPropertyChanged(nameof(KompasLinkStatus));
            InitializeCollectionViews();
            UpdateCalculations();
            ((RelayCommand)ShowInKompasCommand)?.NotifyCanExecuteChanged();
        }

        /// <summary>
        /// Попытаться связать с KOMPАС — загрузить документ по пути (с кэшированием)
        /// </summary>
        private async Task TryLinkToKompasAsync(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return;

            try
            {
                IsLoading = true;
                StatusMessage = "Связывание с KOMPАС...";

                if (File.Exists(filePath))
                {
                    // Загружаем документ в KOMPАС
                    var linkedProduct = await Task.Run(() => _kompasService.LoadDocument(filePath));
                    
                    if (linkedProduct != null)
                    {
                        // Сохраняем в кэш
                        _linkedProductsCache[filePath] = linkedProduct;
                        
                        // Обновляем данные из KOMPАС
                        SetCurrentProduct(linkedProduct, isLinked: true);
                        Debug.WriteLine($"Продукт загружен и закэширован: {filePath}");
                    }
                }
                else
                {
                    Debug.WriteLine($"Файл не найден: {filePath}");
                    StatusMessage = $"Файл не найден: {Path.GetFileName(filePath)}";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка связывания с KOMPАС: {ex.Message}");
                IsLinkedToKompas = false;
            }
            finally
            {
                IsLoading = false;
                ((RelayCommand)ShowInKompasCommand)?.NotifyCanExecuteChanged();
            }
        }

        /// <summary>
        /// Очистить кэш загруженных продуктов (например, при обновлении данных)
        /// </summary>
        public void ClearProductCache()
        {
            _linkedProductsCache.Clear();
            Debug.WriteLine("Кэш продуктов очищен");
        }

        /// <summary>
        /// Очистить конкретный продукт из кэша (при изменении файла)
        /// </summary>
        public void InvalidateProductCache(string filePath)
        {
            if (!string.IsNullOrEmpty(filePath) && _linkedProductsCache.Remove(filePath))
            {
                Debug.WriteLine($"Продукт удалён из кэша: {filePath}");
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
            if (CurrentlySelectedPart == null) return;

            if (!IsLinkedToKompas || CurrentProduct?.Context == null)
            {
                StatusMessage = "Нет связи с KOMPАС. Дождитесь загрузки документа.";
                return;
            }

            try
            {
                _kompasService.ShowDetailInKompas(CurrentlySelectedPart, CurrentProduct);
                StatusMessage = $"Показана деталь: {CurrentlySelectedPart.Name}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
                MessageBox.Show($"Не удалось показать деталь в KOMPАС: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
                
                // Сохраняем в кэш
                _linkedProductsCache[filePath] = product;
                
                CurrentProduct = product;
                IsLinkedToKompas = true;

                UpdateCalculations();
                ((RelayCommand)ShowInKompasCommand)?.NotifyCanExecuteChanged();
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
                
                // Сохраняем в кэш по FilePath
                if (!string.IsNullOrEmpty(product.FilePath))
                {
                    _linkedProductsCache[product.FilePath] = product;
                }
                
                CurrentProduct = product;
                IsLinkedToKompas = true;

                UpdateCalculations();
                ((RelayCommand)ShowInKompasCommand)?.NotifyCanExecuteChanged();
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

        private async Task LoadProductAsync(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return;

            var product = _storageService.Load(fileName);
            if (product != null)
            {
                await LoadAndLinkProductAsync(product, $"Загружено: {product.Name}");
            }
        }

        private async Task SwitchToProductAsync(ProductFileInfo productInfo)
        {
            if (productInfo == null) return;

            var product = _storageService.Load(productInfo.FileName);
            if (product != null)
            {
                IsProductsPanelOpen = false;
                await LoadAndLinkProductAsync(product, $"Переключено на: {product.Name}");
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
            _linkedProductsCache.Clear();
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