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
using Microsoft.WindowsAPICodePack.Dialogs;
using TankManager.Core.Models;
using TankManager.Core.Services;

namespace TankManager.Core.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged, IDisposable
    {
        #region Fields

        private readonly IKompasService _kompasService;
        private readonly ProductStorageService _storageService = new ProductStorageService();
        private readonly ExcelService _excelService = new ExcelService();
        private readonly Dictionary<string, Product> _linkedProductsCache = new Dictionary<string, Product>(StringComparer.OrdinalIgnoreCase);
        
        private bool _isUpdatingCalculations;
        private Product _currentProduct;
        private bool _isLinkedToKompas;
        private ObservableCollection<ProductFileInfo> _savedProducts;
        private ProductFileInfo _selectedSavedProduct;
        private bool _isProductsPanelOpen;
        private string _filePath;
        private string _searchText;
        private MaterialInfo _selectedSheetMaterial;
        private MaterialInfo _selectedTubularProduct;
        private MaterialSortType _sheetMaterialsSortType = MaterialSortType.ByMass;
        private MaterialSortType _tubularProductsSortType = MaterialSortType.ByLength;
        private MaterialSortType _otherMaterialsSortType = MaterialSortType.ByMass;
        private PartModel _currentlySelectedPart;
        private PartModel _selectedDetail;
        private PartModel _selectedStandardPart;
        private bool _isLoading;
        private string _statusMessage;
        private double _totalMassMultipleParts;
        private int _uniquePartsCount;
        private bool _isSnackbarVisible;
        private string _snackbarMessage;
        private System.Threading.Timer _snackbarTimer;

        #endregion

        #region Properties - Product

        public Product CurrentProduct
        {
            get => _currentProduct;
            private set
            {
                if (_currentProduct == value) return;
                
                _currentProduct = value;
                NotifyProductChanged();
                ResetSelections();
                InitializeCollectionViews();
                NotifySaveCommandCanExecuteChanged();
            }
        }

        public ObservableCollection<PartModel> Details => CurrentProduct?.Details;
        public ObservableCollection<MaterialInfo> SheetMaterials => CurrentProduct?.SheetMaterials;
        public ObservableCollection<MaterialInfo> TubularProducts => CurrentProduct?.TubularProducts;
        public ObservableCollection<PartModel> StandardParts => CurrentProduct?.StandardParts;
        public ObservableCollection<MaterialInfo> OtherMaterials => CurrentProduct?.OtherMaterials;

        #endregion

        #region Properties - Collection Views

        public ICollectionView DetailsView { get; private set; }
        public ICollectionView StandardPartsView { get; private set; }
        public ICollectionView SheetMaterialsView { get; private set; }
        public ICollectionView TubularProductsView { get; private set; }
        public ICollectionView OtherMaterialsView { get; private set; }

        #endregion

        #region Properties - KOMPAS Link

        public bool IsLinkedToKompas
        {
            get => _isLinkedToKompas;
            private set => SetProperty(ref _isLinkedToKompas, value, nameof(IsLinkedToKompas), nameof(KompasLinkStatus));
        }

        public string KompasLinkStatus => IsLinkedToKompas ? "🔗 Связан с КОМПАС" : "⚠️ Нет связи с КОМПАС";

        #endregion

        #region Properties - Saved Products

        public ObservableCollection<ProductFileInfo> SavedProducts
        {
            get => _savedProducts;
            private set => SetProperty(ref _savedProducts, value, nameof(SavedProducts));
        }

        public ProductFileInfo SelectedSavedProduct
        {
            get => _selectedSavedProduct;
            set
            {
                if (SetProperty(ref _selectedSavedProduct, value, nameof(SelectedSavedProduct)))
                    ((RelayCommand)DeleteProductCommand)?.NotifyCanExecuteChanged();
            }
        }

        public bool IsProductsPanelOpen
        {
            get => _isProductsPanelOpen;
            set
            {
                if (SetProperty(ref _isProductsPanelOpen, value, nameof(IsProductsPanelOpen)) && value)
                    RefreshSavedProducts();
            }
        }

        #endregion

        #region Properties - Filters & Search

        public string FilePath
        {
            get => _filePath;
            set
            {
                if (SetProperty(ref _filePath, value, nameof(FilePath)))
                    _ = LoadDocumentAsync(value);
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value, nameof(SearchText)))
                {
                    RefreshViews();
                    UpdateCalculations();
                }
            }
        }

        public MaterialInfo SelectedSheetMaterial
        {
            get => _selectedSheetMaterial;
            set
            {
                if (SetProperty(ref _selectedSheetMaterial, value, nameof(SelectedSheetMaterial)))
                {
                    if (value != null) _selectedTubularProduct = null;
                    OnMaterialFilterChanged();
                }
            }
        }

        public MaterialInfo SelectedTubularProduct
        {
            get => _selectedTubularProduct;
            set
            {
                if (SetProperty(ref _selectedTubularProduct, value, nameof(SelectedTubularProduct)))
                {
                    if (value != null) _selectedSheetMaterial = null;
                    OnMaterialFilterChanged();
                }
            }
        }

        public MaterialInfo SelectedMaterialFilter => SelectedSheetMaterial ?? SelectedTubularProduct;

        public MaterialSortType SheetMaterialsSortType
        {
            get => _sheetMaterialsSortType;
            set
            {
                if (SetProperty(ref _sheetMaterialsSortType, value, nameof(SheetMaterialsSortType), nameof(SheetMaterialsSortText)))
                {
                    ApplyMaterialSort(SheetMaterialsView, value);
                }
            }
        }

        public MaterialSortType TubularProductsSortType
        {
            get => _tubularProductsSortType;
            set
            {
                if (SetProperty(ref _tubularProductsSortType, value, nameof(TubularProductsSortType), nameof(TubularProductsSortText)))
                {
                    ApplyMaterialSort(TubularProductsView, value);
                }
            }
        }

        public MaterialSortType OtherMaterialsSortType
        {
            get => _otherMaterialsSortType;
            set
            {
                if (SetProperty(ref _otherMaterialsSortType, value, nameof(OtherMaterialsSortType), nameof(OtherMaterialsSortText)))
                {
                    ApplyMaterialSort(OtherMaterialsView, value);
                }
            }
        }

        public string SheetMaterialsSortText
        {
            get
            {
                switch (_sheetMaterialsSortType)
                {
                    case MaterialSortType.ByName:
                        return "по названию ↑";
                    case MaterialSortType.ByMass:
                        return "по массе ↓";
                    default:
                        return "сортировка";
                }
            }
        }

        public string TubularProductsSortText
        {
            get
            {
                switch (_tubularProductsSortType)
                {
                    case MaterialSortType.ByName:
                        return "по названию ↑";
                    case MaterialSortType.ByLength:
                        return "по длине ↓";
                    case MaterialSortType.ByMass:
                        return "по массе ↓";
                    default:
                        return "сортировка";
                }
            }
        }

        public string OtherMaterialsSortText
        {
            get
            {
                switch (_otherMaterialsSortType)
                {
                    case MaterialSortType.ByName:
                        return "по названию ↑";
                    case MaterialSortType.ByMass:
                        return "по массе ↓";
                    default:
                        return "сортировка";
                }
            }
        }

        #endregion

        #region Properties - Selection

        public PartModel CurrentlySelectedPart
        {
            get => _currentlySelectedPart;
            private set
            {
                if (SetProperty(ref _currentlySelectedPart, value, nameof(CurrentlySelectedPart)))
                    ((RelayCommand)ShowInKompasCommand)?.NotifyCanExecuteChanged();
            }
        }

        public PartModel SelectedDetail
        {
            get => _selectedDetail;
            set
            {
                if (SetProperty(ref _selectedDetail, value, nameof(SelectedDetail)) && value != null)
                {
                    SelectedStandardPart = null;
                    CurrentlySelectedPart = value;
                    _ = LoadDrawingPreviewForSelectedPartAsync();
                }
            }
        }

        public PartModel SelectedStandardPart
        {
            get => _selectedStandardPart;
            set
            {
                if (SetProperty(ref _selectedStandardPart, value, nameof(SelectedStandardPart)) && value != null)
                {
                    SelectedDetail = null;
                    CurrentlySelectedPart = value;
                    _ = LoadDrawingPreviewForSelectedPartAsync();
                }
            }
        }

        #endregion

        #region Properties - Status

        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value, nameof(IsLoading));
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value, nameof(StatusMessage));
        }

        public double TotalMassMultipleParts
        {
            get => _totalMassMultipleParts;
            set
            {
                if (Math.Abs(_totalMassMultipleParts - value) > 0.0001)
                    SetProperty(ref _totalMassMultipleParts, value, nameof(TotalMassMultipleParts));
            }
        }

        public int UniquePartsCount
        {
            get => _uniquePartsCount;
            set => SetProperty(ref _uniquePartsCount, value, nameof(UniquePartsCount));
        }

        public bool IsSnackbarVisible
        {
            get => _isSnackbarVisible;
            set => SetProperty(ref _isSnackbarVisible, value, nameof(IsSnackbarVisible));
        }

        public string SnackbarMessage
        {
            get => _snackbarMessage;
            set => SetProperty(ref _snackbarMessage, value, nameof(SnackbarMessage));
        }

        #endregion

        #region Properties - Server Storage

        /// <summary>
        /// Путь к серверной папке для хранения изделий
        /// </summary>
        public string ServerStorageFolder
        {
            get => _storageService.ServerStorageFolder;
            set
            {
                if (_storageService.ServerStorageFolder != value)
                {
                    _storageService.ServerStorageFolder = value;
                    OnPropertyChanged(nameof(ServerStorageFolder));
                    OnPropertyChanged(nameof(ServerStorageFolderDisplay));
                    OnPropertyChanged(nameof(HasServerStorageFolder));
                    OnPropertyChanged(nameof(IsServerAvailable));
                    ((RelayCommand)ClearServerStorageFolderCommand)?.NotifyCanExecuteChanged();
                    ((RelayCommand)SyncFromServerCommand)?.NotifyCanExecuteChanged();
                }
            }
        }

        /// <summary>
        /// Отображаемый путь к серверной папке (сокращённый)
        /// </summary>
        public string ServerStorageFolderDisplay
        {
            get
            {
                if (string.IsNullOrEmpty(ServerStorageFolder))
                    return "Не указана";
                    
                // Сокращаем путь если слишком длинный
                if (ServerStorageFolder.Length > 40)
                    return "..." + ServerStorageFolder.Substring(ServerStorageFolder.Length - 37);
                    
                return ServerStorageFolder;
            }
        }

        /// <summary>
        /// Указана ли серверная папка
        /// </summary>
        public bool HasServerStorageFolder => _storageService.HasServerFolder;

        /// <summary>
        /// Доступна ли серверная папка
        /// </summary>
        public bool IsServerAvailable => _storageService.IsServerAvailable;

        #endregion

        #region Commands

        public ICommand ShowInKompasCommand { get; private set; }
        public ICommand LoadFromActiveDocumentCommand { get; private set; }
        public ICommand ClearSearchCommand { get; private set; }
        public ICommand LoadProductCommand { get; private set; }
        public ICommand DeleteProductCommand { get; private set; }
        public ICommand ToggleProductsPanelCommand { get; private set; }
        public ICommand SwitchToProductCommand { get; private set; }
        public ICommand CopyAllToClipboardCommand { get; private set; }
        public ICommand CopySheetToClipboardCommand { get; private set; }
        public ICommand CopyTubularProductsToClipboardCommand { get; private set; }
        public ICommand CopyStandartPartsToClipboardCommand { get; private set; }
        public ICommand CopyOtherMaterialsToClipboardCommand { get; private set; }
        public ICommand CopyAllDataToClipboardCommand { get; private set; }
        public ICommand CheckForUpdatesCommand { get; private set; }
        public ICommand LinkToKompasCommand { get; private set; }
        public ICommand SaveProductCommand { get; private set; }
        public ICommand RefreshFromKompasCommand { get; private set; }
        public ICommand SelectServerStorageFolderCommand { get; private set; }
        public ICommand ClearServerStorageFolderCommand { get; private set; }
        public ICommand SyncFromServerCommand { get; private set; }

        #endregion

        #region Constructors

        public MainViewModel() : this(new KompasService()) { }

        public MainViewModel(IKompasService kompasService)
        {
            _kompasService = kompasService ?? throw new ArgumentNullException(nameof(kompasService));
            
            SavedProducts = new ObservableCollection<ProductFileInfo>();
            CurrentProduct = new Product();

            InitializeCommands();
        }

        private void InitializeCommands()
        {
            ShowInKompasCommand = new RelayCommand(ShowDetailInKompas, () => CurrentlySelectedPart != null && IsLinkedToKompas);
            LoadFromActiveDocumentCommand = new RelayCommand(async () => await LoadFromActiveDocumentAsync());
            ClearSearchCommand = new RelayCommand(() => SearchText = string.Empty);
            LoadProductCommand = new RelayCommand<string>(fileName => _ = LoadProductAsync(fileName));
            DeleteProductCommand = new RelayCommand(DeleteSelectedProduct, () => SelectedSavedProduct != null);
            ToggleProductsPanelCommand = new RelayCommand(() => IsProductsPanelOpen = !IsProductsPanelOpen);
            SwitchToProductCommand = new RelayCommand<ProductFileInfo>(info => _ = SwitchToProductAsync(info));
            CopyAllToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyPartsToClipboard, Details), () => Details?.Any() == true);
            CopySheetToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyMaterialsToClipboard, SheetMaterials), () => SheetMaterials?.Any() == true);
            CopyTubularProductsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyTubularProductsToClipboard, TubularProducts), () => TubularProducts?.Any() == true);
            CopyStandartPartsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyPartsToClipboard, StandardParts), () => StandardParts?.Any() == true);
            CopyOtherMaterialsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyMaterialsToClipboard, OtherMaterials), () => OtherMaterials?.Any() == true);
            CopyAllDataToClipboardCommand = new RelayCommand(CopyAllDataToClipboard, () => StandardParts?.Any() == true || SheetMaterials?.Any() == true || TubularProducts?.Any() == true || OtherMaterials?.Any() == true);
            CheckForUpdatesCommand = new RelayCommand(() => UpdateService.CheckForUpdates(showNoUpdateMessage: true));
            LinkToKompasCommand = new RelayCommand(async () => await LinkToKompasAsync(), () => !IsLinkedToKompas && !string.IsNullOrEmpty(CurrentProduct?.FilePath));
            SaveProductCommand = new RelayCommand(async () => await SaveProductAsync(), () => CurrentProduct != null && !string.IsNullOrEmpty(CurrentProduct.Name) && IsLinkedToKompas);
            RefreshFromKompasCommand = new RelayCommand(async () => await RefreshFromKompasAsync(), () => IsLinkedToKompas && !string.IsNullOrEmpty(CurrentProduct?.FilePath));
            SelectServerStorageFolderCommand = new RelayCommand(SelectServerStorageFolder);
            ClearServerStorageFolderCommand = new RelayCommand(ClearServerStorageFolder, () => HasServerStorageFolder);
            SyncFromServerCommand = new RelayCommand(async () => await SyncFromServerAsync(), () => IsServerAvailable);
        }

        #endregion

        #region Product Loading

        private async Task LoadAndLinkProductAsync(Product savedProduct, string successMessage)
        {
            var filePath = savedProduct.FilePath;

            if (!string.IsNullOrEmpty(filePath) && _linkedProductsCache.TryGetValue(filePath, out var cachedProduct))
            {
                SetCurrentProduct(cachedProduct, isLinked: true);
                StatusMessage = $"{successMessage} (из кэша)";
                return;
            }

            SetCurrentProduct(savedProduct, isLinked: false);

            StatusMessage = $"{successMessage} (без связи с КОМПАС)";
            NotifyLinkCommandCanExecuteChanged();
        }

        private async Task LinkToKompasAsync()
        {
            var filePath = CurrentProduct?.FilePath;
            if (string.IsNullOrEmpty(filePath)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Связывание с КОМПАС...";

                if (!File.Exists(filePath))
                {
                    StatusMessage = $"Файл не найден: {Path.GetFileName(filePath)}";
                    MessageBox.Show($"Файл не найден:\n{filePath}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var linkedProduct = await Task.Run(() => _kompasService.LoadDocument(filePath));
                if (linkedProduct != null)
                {
                    _linkedProductsCache[filePath] = linkedProduct;
                    SetCurrentProduct(linkedProduct, isLinked: true);
                    StatusMessage = $"Связано с КОМПАС: {CurrentProduct.Name}";
                }
                else
                {
                    MessageBox.Show(
                        "Не удалось установить связь с КОМПАС.\n\n" +
                        "Возможные причины:\n" +
                        "• КОМПАС-3D не запущен\n" +
                        "• Документ не удалось открыть\n\n" +
                        "Пожалуйста, запустите КОМПАС-3D и попробуйте снова.",
                        "Связь с КОМПАС",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    StatusMessage = "Не удалось связаться с КОМПАС. Убедитесь, что КОМПАС запущен.";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка связывания с КОМПАС: {ex.Message}");
                
                string errorMessage = "Не удалось установить связь с КОМПАС.\n\n";
                
                if (ex.Message.Contains("Не удалось подключиться к KOMPAS-3D") || 
                    ex is InvalidOperationException)
                {
                    errorMessage += "Возможные причины:\n" +
                                  "• КОМПАС-3D не запущен\n" +
                                  "• Нет прав доступа к приложению\n\n" +
                                  "Пожалуйста, запустите КОМПАС-3D и попробуйте снова.";
                }
                else
                {
                    errorMessage += $"Ошибка: {ex.Message}";
                }
                
                MessageBox.Show(errorMessage, "Связь с КОМПАС", MessageBoxButton.OK, MessageBoxImage.Warning);
                StatusMessage = $"Ошибка связи с КОМПАС: {ex.Message}";
                IsLinkedToKompas = false;
                NotifySaveCommandCanExecuteChanged();
            }
            finally
            {
                IsLoading = false;
                NotifyCopyCommandsCanExecuteChanged();
                NotifyLinkCommandCanExecuteChanged();
                NotifyRefreshCommandCanExecuteChanged();
            }
        }

        private async Task RefreshFromKompasAsync()
        {
            var filePath = CurrentProduct?.FilePath;
            if (string.IsNullOrEmpty(filePath) || !IsLinkedToKompas) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Обновление данных из КОМПАС...";

                // Удаляем из кэша, чтобы загрузить актуальные данные
                _linkedProductsCache.Remove(filePath);

                var refreshedProduct = await Task.Run(() => _kompasService.LoadDocument(filePath));
                if (refreshedProduct != null)
                {
                    _linkedProductsCache[filePath] = refreshedProduct;
                    
                    // Инвалидируем кэш превью чертежей для всех деталей
                    InvalidateDrawingPreviewsCache(refreshedProduct);
                    
                    SetCurrentProduct(refreshedProduct, isLinked: true);
                    StatusMessage = $"Данные обновлены: {CurrentProduct.Name}, деталей: {Details?.Count ?? 0}";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка обновления из КОМПАС: {ex.Message}");
                StatusMessage = $"Ошибка обновления: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
                NotifyCopyCommandsCanExecuteChanged();
            }
        }

        /// <summary>
        /// Инвалидирует кэш превью чертежей для всех деталей продукта
        /// </summary>
        private void InvalidateDrawingPreviewsCache(Product product)
        {
            if (product?.Details == null) return;

            foreach (var detail in product.Details)
            {
                detail.InvalidateDrawingPreviewCache();
            }
        }

        private async Task TryLinkToKompasAsync(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Связывание с КОМПАС...";

                if (!File.Exists(filePath))
                {
                    StatusMessage = $"Файл не найден: {Path.GetFileName(filePath)}";
                    return;
                }

                var linkedProduct = await Task.Run(() => _kompasService.LoadDocument(filePath));
                if (linkedProduct != null)
                {
                    _linkedProductsCache[filePath] = linkedProduct;
                    SetCurrentProduct(linkedProduct, isLinked: true);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка связывания с КОМПАС: {ex.Message}");
                IsLinkedToKompas = false;
            }
            finally
            {
                IsLoading = false;
                NotifyCopyCommandsCanExecuteChanged();
            }
        }

        private async Task LoadDocumentAsync(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа...";

                var product = await Task.Run(() => _kompasService.LoadDocument(filePath));
                _linkedProductsCache[filePath] = product;
                CurrentProduct = product;
                IsLinkedToKompas = true;

                UpdateCalculations();
                StatusMessage = $"Загружено изделие: {CurrentProduct.Name}, деталей: {Details.Count}";
            }
            catch (Exception ex)
            {
                ShowError("Ошибка при загрузке файла", ex);
            }
            finally
            {
                IsLoading = false;
                NotifyCopyCommandsCanExecuteChanged();
                NotifyRefreshCommandCanExecuteChanged();
            }
        }

        public async Task LoadFromActiveDocumentAsync()
        {
            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа из КОМПАС...";

                var product = await Task.Run(() => _kompasService.LoadActiveDocument());

                if (!string.IsNullOrEmpty(product.FilePath))
                    _linkedProductsCache[product.FilePath] = product;

                CurrentProduct = product;
                IsLinkedToKompas = true;

                UpdateCalculations();
                StatusMessage = $"Загружено изделие: {CurrentProduct.Name}, деталей: {Details.Count}";
            }
            catch (Exception ex)
            {
                ShowError("Ошибка при загрузке из КОМПАС", ex);
            }
            finally
            {
                IsLoading = false;
                NotifyCopyCommandsCanExecuteChanged();
                NotifyRefreshCommandCanExecuteChanged();
            }
        }

        private async Task LoadProductAsync(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return;

            var product = _storageService.Load(fileName);
            if (product != null)
                await LoadAndLinkProductAsync(product, $"Загружено: {product.Name}");
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

        #endregion

        #region Product Management

        private void SetCurrentProduct(Product product, bool isLinked)
        {
            // Очищаем превью у старого продукта перед переключением
            if (_currentProduct != null && _currentProduct != product)
            {
                if (_currentProduct.Details != null)
                {
                    foreach (var detail in _currentProduct.Details)
                    {
                        detail.FilePreview = null;
                    }
                }
                
                if (_currentProduct.StandardParts != null)
                {
                    foreach (var part in _currentProduct.StandardParts)
                    {
                        part.FilePreview = null;
                    }
                }
            }
            
            _currentProduct = product;
            _isLinkedToKompas = isLinked;

            ResetSelections();
            NotifyProductChanged();
            OnPropertyChanged(nameof(IsLinkedToKompas));
            OnPropertyChanged(nameof(KompasLinkStatus));
            InitializeCollectionViews();
            UpdateCalculations();
            NotifyCopyCommandsCanExecuteChanged();
            NotifySaveCommandCanExecuteChanged();
            NotifyRefreshCommandCanExecuteChanged();
            NotifyLinkCommandCanExecuteChanged();
            
            // Принудительная сборка мусора после смены продукта
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private async Task SaveProductAsync()
        {
            if (CurrentProduct == null || string.IsNullOrEmpty(CurrentProduct.Name)) return;

            try
            {
                // Если есть связь с КОМПАС - предлагаем загрузить превью чертежей
                if (IsLinkedToKompas && CurrentProduct?.Context != null)
                {
                    var detailsWithoutPreview = Details?
                        .Where(d => string.IsNullOrEmpty(d.CdfFilePath) && !d.IsBodyBased && !string.IsNullOrEmpty(d.FilePath))
                        .GroupBy(d => d.FilePath)
                        .Count() ?? 0;

                    if (detailsWithoutPreview > 0)
                    {
                        var result = MessageBox.Show(
                            $"Загрузить превью чертежей для {detailsWithoutPreview} деталей?\n\nЭто может занять некоторое время.",
                            "Сохранение изделия",
                            MessageBoxButton.YesNoCancel,
                            MessageBoxImage.Question);

                        if (result == MessageBoxResult.Cancel)
                            return;

                        if (result == MessageBoxResult.Yes)
                        {
                            IsLoading = true;
                            await LoadAllDrawingPreviewsAsync();
                        }
                    }
                }

                var filePath = _storageService.Save(CurrentProduct);
                var fileName = Path.GetFileName(filePath);
                StatusMessage = $"Сохранено: {fileName}";
                ShowSnackbar($"Изделие \"{CurrentProduct.Name}\" успешно сохранено");
                RefreshSavedProducts();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка сохранения: {ex.Message}");
                StatusMessage = $"Ошибка сохранения: {ex.Message}";
                MessageBox.Show($"Не удалось сохранить изделие:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// Загружает превью чертежей для всех деталей, у которых их ещё нет
        /// </summary>
        private async Task LoadAllDrawingPreviewsAsync()
        {
            if (Details == null || !IsLinkedToKompas || CurrentProduct?.Context == null)
                return;

            // Получаем папку для изображений продукта
            string imagesFolder = _storageService.GetProductImagesFolder(CurrentProduct);

            // Группируем по уникальным деталям (по FilePath), чтобы не загружать одно и то же несколько раз
            var uniqueDetailsWithoutPreview = Details
                .Where(d => string.IsNullOrEmpty(d.CdfFilePath) && !d.IsBodyBased && !string.IsNullOrEmpty(d.FilePath))
                .GroupBy(d => d.FilePath)
                .Select(g => g.First())
                .ToList();

            if (uniqueDetailsWithoutPreview.Count == 0)
                return;

            int loaded = 0;
            int total = uniqueDetailsWithoutPreview.Count;

            foreach (var detail in uniqueDetailsWithoutPreview)
            {
                try
                {
                    StatusMessage = $"Загрузка чертежей: {loaded + 1}/{total} - {detail.Name}";
                    
                    await Task.Run(() => _kompasService.LoadDrawingPreview(detail, CurrentProduct, imagesFolder));
                    
                    // Копируем путь к превью для всех одинаковых деталей
                    if (!string.IsNullOrEmpty(detail.CdfFilePath))
                    {
                        foreach (var samePart in Details.Where(d => d.FilePath == detail.FilePath && d != detail))
                        {
                        samePart.CdfFilePath = detail.CdfFilePath;
                            samePart.SourceCdwPath = detail.SourceCdwPath;
                        }
                    }
                    
                    loaded++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка загрузки превью для {detail.Name}: {ex.Message}");
                    loaded++;
                }
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
                // Если удаляемый продукт - текущий, очищаем его
                if (CurrentProduct != null && 
                    CurrentProduct.Name == SelectedSavedProduct.ProductName &&
                    CurrentProduct.Marking == SelectedSavedProduct.Marking)
                {
                    // Очищаем превью у всех деталей
                    if (Details != null)
                    {
                        foreach (var detail in Details)
                        {
                            detail.FilePreview = null;
                        }
                    }
                    
                    CurrentProduct = new Product();
                    IsLinkedToKompas = false;
                }
                
                // Удаляем из кэша
                InvalidateProductCache(CurrentProduct?.FilePath);
                
                if (_storageService.Delete(SelectedSavedProduct.FileName))
                {
                    StatusMessage = $"Удалено: {SelectedSavedProduct.ProductName}";
                    RefreshSavedProducts();
                    
                    // Принудительная сборка мусора для освобождения файлов
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }
        }

        public void RefreshSavedProducts()
        {
            SavedProducts.Clear();
            foreach (var product in _storageService.GetSavedProducts())
                SavedProducts.Add(product);
        }

        public void ClearProductCache() => _linkedProductsCache.Clear();

        public void InvalidateProductCache(string filePath)
        {
            if (!string.IsNullOrEmpty(filePath))
                _linkedProductsCache.Remove(filePath);
        }

        #endregion

        #region KOMPAS Integration

        private void ShowDetailInKompas()
        {
            if (CurrentlySelectedPart == null) return;

            if (!IsLinkedToKompas || CurrentProduct?.Context == null)
            {
                StatusMessage = "Нет связи с КОМПАС. Дождитесь загрузки документа.";
                return;
            }

            try
            {
                _kompasService.ShowDetailInKompas(CurrentlySelectedPart, CurrentProduct);
                StatusMessage = $"Показана деталь: {CurrentlySelectedPart.Name}";
            }
            catch (Exception ex)
            {
                ShowError("Не удалось показать деталь в КОМПАС", ex);
            }
        }

        private async Task LoadDrawingPreviewForSelectedPartAsync()
        {
            var part = CurrentlySelectedPart;
            if (part == null)
                return;

            // Если есть связь с КОМПАС И у детали ещё нет превью - загружаем
            if (IsLinkedToKompas && CurrentProduct?.Context != null && string.IsNullOrEmpty(part.CdfFilePath))
            {
                try
                {
                    string imagesFolder = _storageService.GetProductImagesFolder(CurrentProduct);
                    await Task.Run(() => _kompasService.LoadDrawingPreview(part, CurrentProduct, imagesFolder));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка загрузки превью чертежа: {ex.Message}");
                }
            }
            
            // Всегда уведомляем UI для отображения превью (из кеша или только что загруженного)
            part.OnPropertyChanged(nameof(part.DrawingPreview));
        }

        #endregion

        #region Collection Views & Filtering

        private void InitializeCollectionViews()
        {
            DetailsView = CreatePartView(Details);
            StandardPartsView = CreatePartView(StandardParts);
            SheetMaterialsView = CreateMaterialView(SheetMaterials, SheetMaterialsSortType);
            TubularProductsView = CreateMaterialView(TubularProducts, TubularProductsSortType);
            OtherMaterialsView = CreateMaterialView(OtherMaterials, OtherMaterialsSortType);

            OnPropertyChanged(nameof(DetailsView));
            OnPropertyChanged(nameof(StandardPartsView));
            OnPropertyChanged(nameof(SheetMaterialsView));
            OnPropertyChanged(nameof(TubularProductsView));
            OnPropertyChanged(nameof(OtherMaterialsView));
        }

        private ICollectionView CreatePartView(ObservableCollection<PartModel> parts)
        {
            if (parts == null) return null;

            var view = CollectionViewSource.GetDefaultView(parts);
            view.Filter = FilterDetails;
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(new PartNameAndMarkingGroupDescription());
            return view;
        }

        private ICollectionView CreateMaterialView(ObservableCollection<MaterialInfo> materials, MaterialSortType sortType)
        {
            if (materials == null) return null;

            var view = CollectionViewSource.GetDefaultView(materials);
            ApplyMaterialSort(view, sortType);
            return view;
        }

        private bool FilterDetails(object obj)
        {
            if (!(obj is PartModel part)) return false;

            var materialFilter = SelectedMaterialFilter;
            if (materialFilter != null && part.Material != materialFilter.Name)
                return false;

            if (string.IsNullOrWhiteSpace(_searchText))
                return true;

            var searchLower = _searchText.ToLower();
            return (part.Name?.ToLower().Contains(searchLower) ?? false) ||
                   (part.Marking?.ToLower().Contains(searchLower) ?? false);
        }

        private void ApplyMaterialSort(ICollectionView view, MaterialSortType sortType)
        {
            if (view == null) return;

            view.SortDescriptions.Clear();
            switch (sortType)
            {
                case MaterialSortType.ByName:
                    view.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));
                    break;
                case MaterialSortType.ByMass:
                    view.SortDescriptions.Add(new SortDescription("TotalMass", ListSortDirection.Descending));
                    break;
                case MaterialSortType.ByLength:
                    view.SortDescriptions.Add(new SortDescription("TotalLength", ListSortDirection.Descending));
                    break;
            }
        }

        private void RefreshViews()
        {
            DetailsView?.Refresh();
            StandardPartsView?.Refresh();
        }

        public void ClearMaterialFilter()
        {
            SelectedSheetMaterial = null;
            SelectedTubularProduct = null;
        }

        #endregion

        #region Calculations

        private void UpdateCalculations()
        {
            if (_isUpdatingCalculations || DetailsView == null) return;

            _isUpdatingCalculations = true;
            try
            {
                var visibleParts = DetailsView.Cast<PartModel>().ToList();
                var groupedParts = visibleParts
                    .GroupBy(p => new { p.Name, p.Marking, p.Material })
                    .Where(g => g.Count() > 1)
                    .ToList();

                TotalMassMultipleParts = groupedParts.Sum(g => g.Sum(p => p.Mass));
                UniquePartsCount = groupedParts.Count;
            }
            finally
            {
                _isUpdatingCalculations = false;
            }
        }

        #endregion

        #region Server Storage & Sync

        private void SelectServerStorageFolder()
        {
            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Title = "Выберите серверную папку для хранения изделий";
                
                if (!string.IsNullOrEmpty(ServerStorageFolder) && Directory.Exists(ServerStorageFolder))
                {
                    dialog.InitialDirectory = ServerStorageFolder;
                }

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    ServerStorageFolder = dialog.FileName;
                    StatusMessage = $"Серверная папка: {ServerStorageFolderDisplay}";
                    
                    // После выбора папки автоматически синхронизируем
                    _ = SyncFromServerAsync();
                }
            }
        }

        private void ClearServerStorageFolder()
        {
            var result = MessageBox.Show(
                "Очистить серверную папку для хранения изделий?\n\nСинхронизация будет отключена.",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                ServerStorageFolder = null;
                StatusMessage = "Серверная папка очищена";
            }
        }

        private async Task SyncFromServerAsync()
        {
            if (!IsServerAvailable)
            {
                StatusMessage = "Серверная папка недоступна";
                return;
            }

            try
            {
                IsLoading = true;
                StatusMessage = "Двусторонняя синхронизация...";

                var syncResult = await Task.Run(() => _storageService.SyncFromServer());

                if (syncResult.Success)
                {
                    if (syncResult.NewProducts > 0 || syncResult.UpdatedProducts > 0)
                    {
                        StatusMessage = $"Синхронизация завершена: новых {syncResult.NewProducts}, обновлено {syncResult.UpdatedProducts}";
                    }
                    else
                    {
                        StatusMessage = "Синхронизация: данные актуальны";
                    }
                }
                else
                {
                    StatusMessage = $"Синхронизация с ошибками: {string.Join(", ", syncResult.Errors.Take(2))}";
                }

                // Обновляем список продуктов
                RefreshSavedProducts();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка синхронизации: {ex.Message}");
                StatusMessage = $"Ошибка синхронизации: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        #endregion

        #region Clipboard

        private void CopyToClipboard<T>(Action<IEnumerable<T>> copyAction, IEnumerable<T> items)
        {
            if (items == null || !items.Any())
            {
                StatusMessage = "Список пуст";
                return;
            }

            copyAction(items);
            StatusMessage = $"Скопировано элементов: {items.Count()}";
        }

        private void CopyAllDataToClipboard()
        {
            _excelService.CopyAllDataToClipboard(StandardParts, SheetMaterials, TubularProducts, OtherMaterials);
            
            int count = (StandardParts?.Count ?? 0) + (SheetMaterials?.Count ?? 0) + (TubularProducts?.Count ?? 0) + (OtherMaterials?.Count ?? 0);
            StatusMessage = $"Скопировано все данные: {count} элементов";
        }

        #endregion

        #region Helper Methods

        private void ShowSnackbar(string message, int durationMs = 3000)
        {
            // Останавливаем предыдущий таймер, если есть
            _snackbarTimer?.Dispose();

            SnackbarMessage = message;
            IsSnackbarVisible = true;

            // Автоматически скрываем через заданное время
            _snackbarTimer = new System.Threading.Timer(_ =>
            {
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    IsSnackbarVisible = false;
                });
            }, null, durationMs, System.Threading.Timeout.Infinite);
        }

        private void NotifyProductChanged()
        {
            OnPropertyChanged(nameof(CurrentProduct));
            OnPropertyChanged(nameof(Details));
            OnPropertyChanged(nameof(SheetMaterials));
            OnPropertyChanged(nameof(TubularProducts));
            OnPropertyChanged(nameof(StandardParts));
            OnPropertyChanged(nameof(OtherMaterials));
        }

        private void NotifyCopyCommandsCanExecuteChanged()
        {
            ((RelayCommand)ShowInKompasCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)CopyAllToClipboardCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)CopySheetToClipboardCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)CopyTubularProductsToClipboardCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)CopyStandartPartsToClipboardCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)CopyOtherMaterialsToClipboardCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)CopyAllDataToClipboardCommand)?.NotifyCanExecuteChanged();
        }

        private void NotifyLinkCommandCanExecuteChanged()
        {
            ((RelayCommand)LinkToKompasCommand)?.NotifyCanExecuteChanged();
        }

        private void NotifySaveCommandCanExecuteChanged()
        {
            ((RelayCommand)SaveProductCommand)?.NotifyCanExecuteChanged();
        }

        private void NotifyRefreshCommandCanExecuteChanged()
        {
            ((RelayCommand)RefreshFromKompasCommand)?.NotifyCanExecuteChanged();
        }

        private void ResetSelections()
        {
            SelectedDetail = null;
            SelectedStandardPart = null;
            SelectedSheetMaterial = null;
            SelectedTubularProduct = null;
            CurrentlySelectedPart = null;
        }

        private void OnMaterialFilterChanged()
        {
            OnPropertyChanged(nameof(SelectedSheetMaterial));
            OnPropertyChanged(nameof(SelectedTubularProduct));
            OnPropertyChanged(nameof(SelectedMaterialFilter));
            DetailsView?.Refresh();
            UpdateCalculations();
        }

        private void ShowError(string message, Exception ex)
        {
            StatusMessage = $"Ошибка: {ex.Message}";
            MessageBox.Show($"{message}: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private bool SetProperty<T>(ref T field, T value, params String[] propertyNames)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            
            field = value;
            foreach (var name in propertyNames.Length > 0 ? propertyNames : new[] { "" })
                if (!string.IsNullOrEmpty(name)) OnPropertyChanged(name);
            return true;
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            // Останавливаем таймер SnackBar
            _snackbarTimer?.Dispose();
            
            // Очищаем превью у всех деталей перед закрытием
            if (CurrentProduct?.Details != null)
            {
                foreach (var detail in CurrentProduct.Details)
                {
                    detail.FilePreview = null;
                    detail.InvalidateDrawingPreviewCache();
                }
            }
            
            if (CurrentProduct?.StandardParts != null)
            {
                foreach (var part in CurrentProduct.StandardParts)
                {
                    part.FilePreview = null;
                    part.InvalidateDrawingPreviewCache();
                }
            }
            
            _linkedProductsCache.Clear();
            CurrentProduct?.Clear();
            _kompasService?.Dispose();
            
            // Принудительная сборка мусора для освобождения файлов
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        #endregion

        #region Nested Types

        private class PartNameAndMarkingGroupDescription : GroupDescription
        {
            public override object GroupNameFromItem(object item, int level, CultureInfo culture)
            {
                return item is PartModel part ? $"{part.Name}|{part.Marking}|{part.Material}" : string.Empty;
            }

            public override bool NamesMatch(object groupName, object itemName)
            {
                return Equals(groupName, itemName);
            }
        }

        #endregion
    }
}