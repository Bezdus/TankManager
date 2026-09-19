using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
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
        private readonly ILogger _logger = new FileLogger();
        private readonly Dictionary<string, Product> _linkedProductsCache = new Dictionary<string, Product>(StringComparer.OrdinalIgnoreCase);
        private PricingSettings _pricingSettings;
        
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
        private CancellationTokenSource _backgroundPreviewCts;
        private bool _isProductSelected;

        #endregion

        #region Properties - Product

        public Product CurrentProduct
        {
            get => _currentProduct;
            private set
            {
                if (_currentProduct == value) return;
                
                var previous = _currentProduct;
                _currentProduct = value;
                NotifyProductChanged();
                ResetSelections();
                InitializeCollectionViews();
                NotifySaveCommandCanExecuteChanged();
                ReleaseProductIfUnused(previous);
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
                {
                    ((RelayCommand)DeleteProductCommand)?.NotifyCanExecuteChanged();
                    ((RelayCommand)DeleteProductLocalCommand)?.NotifyCanExecuteChanged();
                    ((RelayCommand)DeleteProductEverywhereCommand)?.NotifyCanExecuteChanged();
                }
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
                    RunSafe(LoadDocumentAsync(value));
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
                if (SetProperty(ref _selectedDetail, value, nameof(SelectedDetail)))
                {
                    if (value != null)
                    {
                        IsProductSelected = false;
                        SelectedStandardPart = null;
                        CurrentlySelectedPart = value;
                        RunSafe(LoadDrawingPreviewForSelectedPartAsync());
                    }
                    else
                    {
                        // При сбросе выбора очищаем CurrentlySelectedPart
                        CurrentlySelectedPart = null;
                    }
                }
            }
        }

        public PartModel SelectedStandardPart
        {
            get => _selectedStandardPart;
            set
            {
                if (SetProperty(ref _selectedStandardPart, value, nameof(SelectedStandardPart)))
                {
                    if (value != null)
                    {
                        IsProductSelected = false;
                        SelectedDetail = null;
                        CurrentlySelectedPart = value;
                        RunSafe(LoadDrawingPreviewForSelectedPartAsync());
                    }
                    else
                    {
                        // При сбросе выбора очищаем CurrentlySelectedPart
                        CurrentlySelectedPart = null;
                    }
                }
            }
        }

        /// <summary>
        /// Выбрано ли изделие (для отображения сводной карточки)
        /// </summary>
        public bool IsProductSelected
        {
            get => _isProductSelected;
            set
            {
                if (SetProperty(ref _isProductSelected, value, nameof(IsProductSelected)))
                {
                    if (value)
                    {
                        SelectedDetail = null;
                        SelectedStandardPart = null;
                        CurrentlySelectedPart = null;
                    }
                }
            }
        }

        #endregion

        #region Properties - Status

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                if (SetProperty(ref _isLoading, value, nameof(IsLoading)))
                    NotifyLoadingDependentCommands();
            }
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

        /// <summary>
        /// Настройки расценок для расчёта стоимости
        /// </summary>
        public PricingSettings PricingSettings
        {
            get => _pricingSettings;
            private set => SetProperty(ref _pricingSettings, value, nameof(PricingSettings));
        }

        #endregion

        #region Commands

        public ICommand ShowInKompasCommand { get; private set; }
        public ICommand LoadFromActiveDocumentCommand { get; private set; }
        public ICommand ClearSearchCommand { get; private set; }
        public ICommand LoadProductCommand { get; private set; }
        public ICommand DeleteProductCommand { get; private set; }
        public ICommand DeleteProductLocalCommand { get; private set; }
        public ICommand DeleteProductEverywhereCommand { get; private set; }
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
        public ICommand ExportToExcelCommand { get; private set; }
        public ICommand OpenPricingSettingsCommand { get; private set; }

        #endregion

        #region Constructors

        public MainViewModel() : this(new KompasService()) { }

        public MainViewModel(IKompasService kompasService)
        {
            _kompasService = kompasService ?? throw new ArgumentNullException(nameof(kompasService));
            _pricingSettings = PricingSettings.Load();
            
            SavedProducts = new ObservableCollection<ProductFileInfo>();
            CurrentProduct = new Product();

            InitializeCommands();
        }

        private void InitializeCommands()
        {
            ShowInKompasCommand = new RelayCommand(ShowDetailInKompas, () => CurrentlySelectedPart != null && IsLinkedToKompas && !IsLoading);
            LoadFromActiveDocumentCommand = new RelayCommand(async () => await LoadFromActiveDocumentAsync(), () => !IsLoading);
            ClearSearchCommand = new RelayCommand(() => SearchText = string.Empty);
            LoadProductCommand = new RelayCommand<string>(LoadProduct);
            DeleteProductCommand = new RelayCommand(async () => await DeleteSelectedProductAsync(everywhere: true), () => SelectedSavedProduct != null && !IsLoading);
            DeleteProductLocalCommand = new RelayCommand(async () => await DeleteSelectedProductAsync(everywhere: false), () => SelectedSavedProduct != null && !IsLoading);
            DeleteProductEverywhereCommand = new RelayCommand(async () => await DeleteSelectedProductAsync(everywhere: true), () => SelectedSavedProduct != null && !IsLoading);
            ToggleProductsPanelCommand = new RelayCommand(() => IsProductsPanelOpen = !IsProductsPanelOpen);
            SwitchToProductCommand = new RelayCommand<ProductFileInfo>(SwitchToProduct);
            CopyAllToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyPartsToClipboard, Details), () => Details?.Any() == true);
            CopySheetToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyMaterialsToClipboard, SheetMaterials), () => SheetMaterials?.Any() == true);
            CopyTubularProductsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyTubularProductsToClipboard, TubularProducts), () => TubularProducts?.Any() == true);
            CopyStandartPartsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyPartsToClipboard, StandardParts), () => StandardParts?.Any() == true);
            CopyOtherMaterialsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyMaterialsToClipboard, OtherMaterials), () => OtherMaterials?.Any() == true);
            CopyAllDataToClipboardCommand = new RelayCommand(CopyAllDataToClipboard, () => StandardParts?.Any() == true || SheetMaterials?.Any() == true || TubularProducts?.Any() == true || OtherMaterials?.Any() == true);
            CheckForUpdatesCommand = new RelayCommand(() => UpdateService.CheckForUpdates(showNoUpdateMessage: true));
            LinkToKompasCommand = new RelayCommand(async () => await LinkToKompasAsync(), () => !IsLinkedToKompas && !string.IsNullOrEmpty(CurrentProduct?.FilePath) && !IsLoading);
            SaveProductCommand = new RelayCommand(async () => await SaveProductAsync(), () => CurrentProduct != null && !string.IsNullOrEmpty(CurrentProduct.Name) && IsLinkedToKompas && !IsLoading);
            RefreshFromKompasCommand = new RelayCommand(async () => await RefreshFromKompasAsync(), () => IsLinkedToKompas && !string.IsNullOrEmpty(CurrentProduct?.FilePath) && !IsLoading);
            SelectServerStorageFolderCommand = new RelayCommand(SelectServerStorageFolder);
            ClearServerStorageFolderCommand = new RelayCommand(ClearServerStorageFolder, () => HasServerStorageFolder);
            SyncFromServerCommand = new RelayCommand(async () => await SyncFromServerAsync(), () => IsServerAvailable && !IsLoading);
            ExportToExcelCommand = new RelayCommand(ExportToExcel, () => Details?.Any() == true || StandardParts?.Any() == true || SheetMaterials?.Any() == true || TubularProducts?.Any() == true || OtherMaterials?.Any() == true);
            OpenPricingSettingsCommand = new RelayCommand(OpenPricingSettings);
        }

        #endregion

        #region Pricing

        private void OpenPricingSettings()
        {
            var dialog = new TankManager.Views.PricingSettingsDialog(_pricingSettings);
            dialog.Owner = Application.Current.MainWindow;
            if (dialog.ShowDialog() == true)
            {
                var newSettings = dialog.PricingSettings;

                try
                {
                    newSettings.Save();
                }
                catch (Exception ex)
                {
                    _logger.LogError("Не удалось сохранить расценки", ex);
                    MessageBox.Show(
                        $"Расценки применены, но не сохранены на диск:\n{ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                PricingSettings = newSettings;
                RecalculateAllCosts();
            }
        }

        /// <summary>
        /// Пересчитать стоимость всех деталей на основе текущих расценок
        /// </summary>
        public void RecalculateAllCosts()
        {
            if (_pricingSettings == null) return;

            int unreliableOperations = 0;

            var allParts = (Details ?? Enumerable.Empty<PartModel>())
                .Concat(StandardParts ?? Enumerable.Empty<PartModel>());

            foreach (var part in allParts)
            {
                // Считаем стоимость металла
                part.MetalCost = CalculateMetalCost(part);

                // Считаем стоимость операций
                foreach (var op in part.Operations)
                {
                    var rolling = op as RollingOperation;
                    if (rolling != null)
                        rolling.PartMass = part.Mass;

                    op.CalculateCost(_pricingSettings);
                    if (!op.IsCostReliable)
                        unreliableOperations++;
                }

                part.RecalculateOperationsCost();
            }

            CurrentProduct?.NotifyAggregatesChanged();

            if (unreliableOperations > 0)
            {
                _logger.LogWarning($"Стоимость не рассчитана для операций: {unreliableOperations}");
                ShowSnackbar($"Стоимость не рассчитана для операций: {unreliableOperations} (нет исходных данных)", 6000);
            }
        }

        private double CalculateMetalCost(PartModel part)
        {
            if (part.ProductType == ProductType.PurchasedPart)
                return 0;

            if (part.ProductType == ProductType.TubularProduct)
            {
                // Длина неизвестна (например, старый файл без длины) — оставляем сохранённую стоимость
                if (part.Length <= 0)
                    return part.MetalCost;

                double pricePerMeter = _pricingSettings.GetTubularPricePerMeter(part.Material);
                return (part.Length / 1000.0) * pricePerMeter;
            }

            if (part.ProductType == ProductType.SheetMaterial)
                return part.Mass * _pricingSettings.SheetMetalPricePerKg;

            return part.Mass * _pricingSettings.OtherMetalPricePerKg;
        }

        #endregion

        #region Product Loading

        private void LoadAndLinkProduct(Product savedProduct, string successMessage)
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
            if (IsLoading) return;

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
                    CacheProduct(filePath, linkedProduct);
                    RestoreImagePathsFromSaved(linkedProduct);
                    SetCurrentProduct(linkedProduct, isLinked: true);
                    RecalculateAllCosts();
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
                _logger.LogWarning($"Ошибка связывания с КОМПАС: {ex.Message}");
                
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
            if (IsLoading) return;

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
                    CacheProduct(filePath, refreshedProduct);
                    RestoreImagePathsFromSaved(refreshedProduct);
                    SetCurrentProduct(refreshedProduct, isLinked: true);
                    RecalculateAllCosts();
                    StatusMessage = $"Данные обновлены: {CurrentProduct.Name}, деталей: {Details?.Count ?? 0}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка обновления из КОМПАС: {ex.Message}");
                StatusMessage = $"Ошибка обновления: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
                NotifyCopyCommandsCanExecuteChanged();
            }
        }

        /// <summary>
        /// Восстанавливает пути к изображениям из ранее сохранённого продукта.
        /// Позволяет избежать повторных COM-вызовов при связывании/обновлении из КОМПАС,
        /// когда PNG-файлы уже существуют на диске и актуальны.
        /// </summary>
        private void RestoreImagePathsFromSaved(Product kompasProduct)
        {
            if (kompasProduct?.Details == null)
                return;

            try
            {
                var savedProduct = _storageService.TryLoadSavedProduct(kompasProduct);
                if (savedProduct?.Details == null)
                    return;

                // Строим словарь сохранённых деталей по FilePath для быстрого поиска
                var savedByFilePath = new Dictionary<string, PartModel>();
                foreach (var saved in savedProduct.Details)
                {
                    if (!string.IsNullOrEmpty(saved.FilePath) && !savedByFilePath.ContainsKey(saved.FilePath))
                        savedByFilePath[saved.FilePath] = saved;
                }

                foreach (var detail in kompasProduct.Details)
                {
                    if (string.IsNullOrEmpty(detail.FilePath))
                        continue;

                    PartModel savedDetail;
                    if (!savedByFilePath.TryGetValue(detail.FilePath, out savedDetail))
                        continue;

                    // Восстанавливаем путь к PNG чертежа, если файл существует и актуален
                    if (!string.IsNullOrEmpty(savedDetail.CdfFilePath) && File.Exists(savedDetail.CdfFilePath))
                    {
                        // Проверяем актуальность: если есть исходный CDW, PNG должен быть не старше
                        if (string.IsNullOrEmpty(savedDetail.SourceCdwPath) || !File.Exists(savedDetail.SourceCdwPath) ||
                            File.GetLastWriteTimeUtc(savedDetail.CdfFilePath) >= File.GetLastWriteTimeUtc(savedDetail.SourceCdwPath))
                        {
                            detail.CdfFilePath = savedDetail.CdfFilePath;
                            detail.SourceCdwPath = savedDetail.SourceCdwPath;
                        }
                    }

                    // Восстанавливаем путь к превью 3D-файла, если существует и актуален
                    if (!string.IsNullOrEmpty(savedDetail.FilePreviewPngPath) && File.Exists(savedDetail.FilePreviewPngPath))
                    {
                        if (string.IsNullOrEmpty(detail.FilePath) || !File.Exists(detail.FilePath) ||
                            File.GetLastWriteTimeUtc(savedDetail.FilePreviewPngPath) >= File.GetLastWriteTimeUtc(detail.FilePath))
                        {
                            detail.FilePreviewPngPath = savedDetail.FilePreviewPngPath;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка восстановления путей изображений: {ex.Message}");
            }
        }

        private async Task TryLinkToKompasAsync(string filePath)
        {
            if (IsLoading) return;

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
                    CacheProduct(filePath, linkedProduct);
                    RestoreImagePathsFromSaved(linkedProduct);
                    SetCurrentProduct(linkedProduct, isLinked: true);
                    RecalculateAllCosts();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка связывания с КОМПАС: {ex.Message}");
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
            if (IsLoading) return;

            if (string.IsNullOrEmpty(filePath)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа...";

                var product = await Task.Run(() => _kompasService.LoadDocument(filePath));
                CacheProduct(filePath, product);
                RestoreImagePathsFromSaved(product);
                CurrentProduct = product;
                IsLinkedToKompas = true;

                UpdateCalculations();
                RecalculateAllCosts();
                StatusMessage = $"Загружено изделие: {CurrentProduct.Name}, деталей: {Details.Count}";

                RunSafe(GenerateDrawingPreviewsInBackgroundAsync());
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
                NotifySaveCommandCanExecuteChanged();
            }
        }

        public async Task LoadFromActiveDocumentAsync()
        {
            if (IsLoading) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа из КОМПАС...";

                var product = await Task.Run(() => _kompasService.LoadActiveDocument());

                if (!string.IsNullOrEmpty(product.FilePath))
                    CacheProduct(product.FilePath, product);

                RestoreImagePathsFromSaved(product);
                CurrentProduct = product;
                IsLinkedToKompas = true;

                UpdateCalculations();
                RecalculateAllCosts();
                StatusMessage = $"Загружено изделие: {CurrentProduct.Name}, деталей: {Details.Count}";

                RunSafe(GenerateDrawingPreviewsInBackgroundAsync());
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
                NotifySaveCommandCanExecuteChanged();
            }
        }

        private void LoadProduct(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return;

            var product = _storageService.Load(fileName);
            if (product != null)
                LoadAndLinkProduct(product, $"Загружено: {product.Name}");
        }

        private void SwitchToProduct(ProductFileInfo productInfo)
        {
            if (productInfo == null) return;

            var product = _storageService.Load(productInfo.FileName);
            if (product != null)
            {
                IsProductsPanelOpen = false;
                LoadAndLinkProduct(product, $"Переключено на: {product.Name}");
            }
        }

        #endregion

        #region Product Management

        private void SetCurrentProduct(Product product, bool isLinked)
        {
            // Отменяем предыдущую фоновую генерацию превью
            _backgroundPreviewCts?.Cancel();

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
            
            var previous = _currentProduct;
            _currentProduct = product;
            _isLinkedToKompas = isLinked;

            ResetSelections();
            NotifyProductChanged();
            OnPropertyChanged(nameof(IsLinkedToKompas));
            OnPropertyChanged(nameof(KompasLinkStatus));
            InitializeCollectionViews();
            UpdateCalculations();
            product.NotifyAggregatesChanged();
            NotifyCopyCommandsCanExecuteChanged();
            NotifySaveCommandCanExecuteChanged();
            NotifyRefreshCommandCanExecuteChanged();
            NotifyLinkCommandCanExecuteChanged();

            ReleaseProductIfUnused(previous);

            // Фоновая синхронизация изображений с сервером
            RunSafe(SyncProductImagesInBackgroundAsync(product));

            // Фоновая генерация превью чертежей для нового изделия с КОМПАС
            if (isLinked)
            {
                RunSafe(GenerateDrawingPreviewsInBackgroundAsync());
            }

        }

        private async Task SaveProductAsync()
        {
            if (IsLoading) return;
            if (CurrentProduct == null || string.IsNullOrEmpty(CurrentProduct.Name)) return;

            var product = CurrentProduct;

            try
            {
                IsLoading = true;

                // Получаем папку для изображений продукта
                string imagesFolder = await Task.Run(() => _storageService.GetProductImagesFolder(product));

                // Сохраняем превью 3D-файлов для работы без исходных файлов КОМПАС
                await SaveAllFilePreviewsAsync(imagesFolder);

                // Запись на диск и в сетевую папку — вне UI-потока
                var filePath = await Task.Run(() => _storageService.Save(product));
                var fileName = Path.GetFileName(filePath);

                var serverError = _storageService.LastServerError;
                if (string.IsNullOrEmpty(serverError))
                {
                    StatusMessage = $"Сохранено: {fileName}";
                    ShowSnackbar($"Изделие \"{product.Name}\" успешно сохранено");
                }
                else
                {
                    StatusMessage = $"Сохранено локально: {fileName}. Сервер: {serverError}";
                    ShowSnackbar($"Изделие сохранено только локально: {serverError}", 6000);
                }

                RefreshSavedProducts();
            }
            catch (Exception ex)
            {
                _logger.LogError("Ошибка сохранения изделия", ex);
                StatusMessage = $"Ошибка сохранения: {ex.Message}";
                MessageBox.Show($"Не удалось сохранить изделие:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// Сохраняет превью 3D-файлов для всех уникальных деталей
        /// </summary>
        private async Task SaveAllFilePreviewsAsync(string imagesFolder)
        {
            if (string.IsNullOrEmpty(imagesFolder))
                return;

            // Собираем все детали из Details и StandardParts
            var allParts = (Details ?? Enumerable.Empty<PartModel>())
                .Concat(StandardParts ?? Enumerable.Empty<PartModel>())
                .Where(p => !string.IsNullOrEmpty(p.FilePath) && string.IsNullOrEmpty(p.FilePreviewPngPath))
                .GroupBy(p => p.FilePath)
                .Select(g => g.First())
                .ToList();

            if (allParts.Count == 0)
                return;

            int saved = 0;
            int total = allParts.Count;

            foreach (var part in allParts)
            {
                try
                {
                    StatusMessage = $"Сохранение превью: {saved + 1}/{total}";
                    
                    // Сохраняем превью
                    var savedPath = await Task.Run(() => part.SaveFilePreview(imagesFolder));
                    
                    // Копируем путь к превью для всех одинаковых деталей
                    if (!string.IsNullOrEmpty(savedPath))
                    {
                        var sameParts = (Details ?? Enumerable.Empty<PartModel>())
                            .Concat(StandardParts ?? Enumerable.Empty<PartModel>())
                            .Where(p => p.FilePath == part.FilePath && p != part);
                        
                        foreach (var samePart in sameParts)
                        {
                            samePart.FilePreviewPngPath = savedPath;
                        }
                    }
                    
                    saved++;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Ошибка сохранения превью для {part.Name}: {ex.Message}");
                    saved++;
                }
            }
        }

        /// <summary>
        /// Фоновая генерация превью чертежей для всех деталей без превью.
        /// Запускается автоматически после загрузки нового изделия из КОМПАС.
        /// </summary>
        private async Task GenerateDrawingPreviewsInBackgroundAsync()
        {
            if (Details == null || !IsLinkedToKompas || CurrentProduct?.Context == null)
                return;

            _backgroundPreviewCts?.Cancel();
            _backgroundPreviewCts = new CancellationTokenSource();
            var token = _backgroundPreviewCts.Token;
            var product = CurrentProduct;

            string imagesFolder = _storageService.GetProductImagesFolder(product);

            var detailsToProcess = Details
                .Where(d => string.IsNullOrEmpty(d.CdfFilePath) && !d.IsBodyBased && !string.IsNullOrEmpty(d.FilePath))
                .GroupBy(d => d.FilePath)
                .Select(g => g.First())
                .ToList();

            if (detailsToProcess.Count == 0)
                return;

            int processed = 0;
            int total = detailsToProcess.Count;

            foreach (var detail in detailsToProcess)
            {
                if (token.IsCancellationRequested || CurrentProduct != product)
                    return;

                try
                {
                    StatusMessage = $"Генерация превью чертежей: {processed + 1}/{total}";

                    await Task.Run(() => _kompasService.LoadDrawingPreview(detail, product, imagesFolder), token);

                    if (!string.IsNullOrEmpty(detail.CdfFilePath))
                    {
                        foreach (var samePart in Details.Where(d => d.FilePath == detail.FilePath && d != detail))
                        {
                            samePart.CdfFilePath = detail.CdfFilePath;
                            samePart.SourceCdwPath = detail.SourceCdwPath;
                        }
                    }

                    processed++;
                }
                catch (OperationCanceledException)
                {
                    return;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Фоновая генерация превью {detail.Name}: {ex.Message}");
                    processed++;
                }
            }

            if (CurrentProduct == product && !token.IsCancellationRequested)
            {
                StatusMessage = $"Превью чертежей готовы: {processed}/{total}";
            }
        }

        private async Task DeleteSelectedProductAsync(bool everywhere)
        {
            var info = SelectedSavedProduct;
            if (info == null || IsLoading) return;

            var confirmation = everywhere
                ? MessageBox.Show(
                    $"Удалить \"{info.ProductName}\" локально и с сервера?\n\nЭто действие нельзя отменить.",
                    "Подтверждение удаления",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning)
                : MessageBox.Show(
                    $"Удалить \"{info.ProductName}\" локально?\n\nИзделие останется на сервере.",
                    "Подтверждение удаления",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

            if (confirmation != MessageBoxResult.Yes) return;

            var currentFilePath = CurrentProduct?.FilePath;
            ClearCurrentProductIfMatches(info);
            if (CurrentProduct == null || CurrentProduct.FilePath != currentFilePath)
                InvalidateProductCache(currentFilePath);

            try
            {
                IsLoading = true;

                // Удаление с повторными попытками и сетевые операции — вне UI-потока
                bool deleted = await Task.Run(() => everywhere
                    ? _storageService.Delete(info.FileName)
                    : _storageService.DeleteLocal(info.FileName));

                if (deleted)
                {
                    StatusMessage = everywhere
                        ? $"Удалено отовсюду: {info.ProductName}"
                        : $"Удалено локально: {info.ProductName}";
                    RefreshSavedProducts();
                }

                if (everywhere && !string.IsNullOrEmpty(_storageService.LastServerError))
                {
                    StatusMessage += $". {_storageService.LastServerError}";
                    ShowSnackbar(_storageService.LastServerError, 6000);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Ошибка удаления изделия", ex);
                ShowError("Не удалось удалить изделие", ex);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ClearCurrentProductIfMatches(ProductFileInfo productInfo)
        {
            if (CurrentProduct != null &&
                CurrentProduct.Name == productInfo.ProductName &&
                CurrentProduct.Marking == productInfo.Marking)
            {
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
        }

        public void RefreshSavedProducts()
        {
            SavedProducts.Clear();
            foreach (var product in _storageService.GetSavedProducts())
                SavedProducts.Add(product);
        }

        public void InvalidateProductCache(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return;

            Product removed;
            if (_linkedProductsCache.TryGetValue(filePath, out removed))
            {
                _linkedProductsCache.Remove(filePath);
                ReleaseProductIfUnused(removed);
            }
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
                if (IsKompasLinkLost())
                    MarkLinkLost();
                else
                    ShowError("Не удалось показать деталь в КОМПАС", ex);
            }
        }

        private async Task LoadDrawingPreviewForSelectedPartAsync()
        {
            var part = CurrentlySelectedPart;
            if (part == null)
                return;

            if (IsLinkedToKompas && IsKompasLinkLost())
                MarkLinkLost();

            bool needsPreview = string.IsNullOrEmpty(part.CdfFilePath);
            bool isStale = !needsPreview && ImageSyncService.IsDrawingPreviewStale(part.CdfFilePath, part.SourceCdwPath);

            // Если есть связь с КОМПАС И у детали нет превью или оно устарело - загружаем/перегенерируем
            if (IsLinkedToKompas && CurrentProduct?.Context != null && (needsPreview || isStale))
            {
                if (isStale)
                {
                    part.InvalidateDrawingPreviewCache();
                }

                try
                {
                    string imagesFolder = _storageService.GetProductImagesFolder(CurrentProduct);
                    await Task.Run(() => _kompasService.LoadDrawingPreview(part, CurrentProduct, imagesFolder));
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Ошибка загрузки превью чертежа: {ex.Message}");
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
                    RunSafe(SyncFromServerAsync());
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
            if (IsLoading) return;

            if (!IsServerAvailable)
            {
                StatusMessage = "Серверная папка недоступна";
                return;
            }

            try
            {
                IsLoading = true;
                StatusMessage = "Двусторонняя синхронизация...";

                var syncResult = await Task.Run(() => _storageService.SyncFromServer(skipImages: true));

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
                _logger.LogWarning($"Ошибка синхронизации: {ex.Message}");
                StatusMessage = $"Ошибка синхронизации: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// Фоновая синхронизация изображений продукта с сервером
        /// </summary>
        private async Task SyncProductImagesInBackgroundAsync(Product product)
        {
            if (product == null || string.IsNullOrEmpty(product.Name) || !_storageService.IsServerAvailable)
                return;

            try
            {
                await Task.Run(() => _storageService.SyncProductImagesWithServer(product));
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка фоновой синхронизации изображений: {ex.Message}");
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
            int count = items.Count();
            StatusMessage = $"Скопировано элементов: {count}";
            ShowSnackbar($"Скопировано {count} элементов в буфер обмена");
        }

        private void CopyAllDataToClipboard()
        {
            _excelService.CopyAllDataToClipboard(StandardParts, SheetMaterials, TubularProducts, OtherMaterials);

            int count = (StandardParts?.Count ?? 0) + (SheetMaterials?.Count ?? 0) + (TubularProducts?.Count ?? 0) + (OtherMaterials?.Count ?? 0);
            StatusMessage = $"Скопировано все данные: {count} элементов";
            ShowSnackbar($"Все данные скопированы в Excel ({count} элементов)");
        }

        private void ExportToExcel()
        {
            try
            {
                var filePath = _excelService.ExportToExcelFile(
                    CurrentProduct?.Name,
                    Details,
                    StandardParts,
                    SheetMaterials,
                    TubularProducts,
                    OtherMaterials);

                if (filePath != null)
                {
                    StatusMessage = $"Файл сохранён: {filePath}";
                    ShowSnackbar("Ведомость материалов сохранена в Excel");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
            ((RelayCommand)ExportToExcelCommand)?.NotifyCanExecuteChanged();
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
            IsProductSelected = false;
        }

        private void OnMaterialFilterChanged()
        {
            OnPropertyChanged(nameof(SelectedSheetMaterial));
            OnPropertyChanged(nameof(SelectedTubularProduct));
            OnPropertyChanged(nameof(SelectedMaterialFilter));
            DetailsView?.Refresh();
            UpdateCalculations();
        }

        /// <summary>
        /// Запускает фоновую задачу и логирует необработанное исключение
        /// </summary>
        private void RunSafe(Task task)
        {
            task.ContinueWith(
                t => _logger.LogError("Необработанная ошибка фоновой операции", t.Exception?.GetBaseException()),
                TaskContinuationOptions.OnlyOnFaulted);
        }

        private void NotifyLoadingDependentCommands()
        {
            ((RelayCommand)ShowInKompasCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)LoadFromActiveDocumentCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)LinkToKompasCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)SaveProductCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)RefreshFromKompasCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)SyncFromServerCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)DeleteProductCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)DeleteProductLocalCommand)?.NotifyCanExecuteChanged();
            ((RelayCommand)DeleteProductEverywhereCommand)?.NotifyCanExecuteChanged();
        }

        /// <summary>
        /// Кладёт связанный продукт в кэш; заменённый продукт освобождается, если он больше не используется
        /// </summary>
        private void CacheProduct(string filePath, Product product)
        {
            Product old;
            bool hadOld = _linkedProductsCache.TryGetValue(filePath, out old);
            _linkedProductsCache[filePath] = product;

            if (hadOld && old != product)
                ReleaseProductIfUnused(old);
        }

        /// <summary>
        /// Освобождает COM-контекст продукта, если он не текущий и не лежит в кэше
        /// </summary>
        private void ReleaseProductIfUnused(Product product)
        {
            if (product == null || product == _currentProduct)
                return;

            if (_linkedProductsCache.ContainsValue(product))
                return;

            product.Dispose();
        }

        private bool IsKompasLinkLost()
        {
            var context = CurrentProduct?.Context;
            return context != null && !context.IsDocumentAlive();
        }

        private void MarkLinkLost()
        {
            IsLinkedToKompas = false;
            StatusMessage = "Связь с КОМПАС потеряна. Запустите КОМПАС и нажмите «Связать».";
            ShowSnackbar("Связь с КОМПАС потеряна", 5000);
            NotifyCopyCommandsCanExecuteChanged();
            NotifySaveCommandCanExecuteChanged();
            NotifyLinkCommandCanExecuteChanged();
            NotifyRefreshCommandCanExecuteChanged();
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
            // Отменяем фоновую генерацию превью
            _backgroundPreviewCts?.Cancel();
            _backgroundPreviewCts?.Dispose();

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
            
            foreach (var cached in _linkedProductsCache.Values.ToList())
            {
                if (cached != CurrentProduct)
                    cached.Dispose();
            }

            _linkedProductsCache.Clear();
            CurrentProduct?.Clear();
            _kompasService?.Dispose();
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