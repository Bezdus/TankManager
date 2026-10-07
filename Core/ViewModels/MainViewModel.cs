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
        private MaterialInfo _selectedOtherMaterial;
        private SnackbarKind _snackbarKind;
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
        private bool _isSyncing;
        private DateTime _lastSyncUtc = DateTime.MinValue;

        // Автосинхронизация при открытии списка изделий не чаще этого интервала
        private static readonly TimeSpan AutoSyncInterval = TimeSpan.FromMinutes(2);

        #endregion

        #region Properties - App Mode

        public const string ViewerModeLoadMessage = "Режим просмотра: загрузка сборок из КОМПАС доступна только конструктору";

        /// <summary>
        /// Режим просмотра (без КОМПАС): только чтение изделий с сервера
        /// </summary>
        public bool IsViewerMode => AppMode.IsViewer;

        /// <summary>
        /// Режим конструктора: загрузка из КОМПАС, сохранение, удаление
        /// </summary>
        public bool IsEngineerMode => !AppMode.IsViewer;

        /// <summary>
        /// Режим технолога (без КОМПАС): как просмотр, но можно править операции деталей
        /// </summary>
        public bool IsTechnologistMode => AppMode.IsTechnologist;

        /// <summary>
        /// Доступна ли правка операций изготовления (конструктор или технолог)
        /// </summary>
        public bool CanEditOperations => AppMode.CanEditOperations;

        public string ModeBannerText => IsTechnologistMode ? "Технолог"
            : IsViewerMode ? "Просмотр"
            : "Конструктор";

        public string ModeBannerToolTip => IsTechnologistMode
            ? "Режим технолога: изделия загружаются с сервера и доступны только для чтения, кроме операций изготовления: их можно править (кнопка «Изменить» в карточке детали). Сборки загружает и сохраняет конструктор."
            : IsViewerMode
            ? "Режим просмотра: изделия загружаются с сервера и доступны только для чтения. Сборки загружает и сохраняет конструктор."
            : "Режим конструктора: загрузка сборок из КОМПАС, сохранение и удаление изделий на сервере.";

        /// <summary>
        /// Значок текущего режима (Segoe MDL2 Assets)
        /// </summary>
        public string ModeIcon => IsTechnologistMode ? ""
            : IsViewerMode ? ""
            : "";

        /// <summary>
        /// Варианты настройки режима для выпадающего списка
        /// </summary>
        public IReadOnlyList<KeyValuePair<AppModeSetting, string>> ModeOptions { get; } = new[]
        {
            new KeyValuePair<AppModeSetting, string>(AppModeSetting.Auto, "Автоматически (по наличию КОМПАС)"),
            new KeyValuePair<AppModeSetting, string>(AppModeSetting.Engineer, "Конструктор"),
            new KeyValuePair<AppModeSetting, string>(AppModeSetting.Technologist, "Технолог (правка операций)"),
            new KeyValuePair<AppModeSetting, string>(AppModeSetting.Viewer, "Просмотр")
        };

        /// <summary>
        /// Настройка режима; применяется после перезапуска
        /// </summary>
        public AppModeSetting ModeSetting
        {
            get => _storageService.ModeSetting;
            set
            {
                if (_storageService.ModeSetting == value) return;

                _storageService.ModeSetting = value;
                OnPropertyChanged(nameof(ModeSetting));
                OnPropertyChanged(nameof(IsModeRestartRequired));
            }
        }

        /// <summary>
        /// Настройка режима изменена и вступит в силу после перезапуска
        /// </summary>
        public bool IsModeRestartRequired => ModeSetting != AppMode.Setting;

        #endregion

        #region Properties - Current User

        /// <summary>
        /// ФИО (или логин) текущего сотрудника
        /// </summary>
        public string CurrentUserName => CurrentUser.DisplayName;

        public string CurrentUserLogin => CurrentUser.Login;

        /// <summary>
        /// Роль задаётся в общем списке сотрудников (переключатель режима не действует)
        /// </summary>
        public bool AccountsEnabled => CurrentUser.AccountsEnabled && !CurrentUser.NeedsRegistration;

        /// <summary>
        /// Переключатель режима доступен, только пока роли не назначаются по списку сотрудников
        /// </summary>
        public bool IsModeSwitchAvailable => !AccountsEnabled;

        public bool IsAdmin => CurrentUser.IsAdmin;

        /// <summary>
        /// Окно «Сотрудники»: администратор при заданной серверной папке
        /// </summary>
        public bool CanManageUsers => IsAdmin && HasServerStorageFolder;

        private void NotifyCurrentUserChanged()
        {
            OnPropertyChanged(nameof(CurrentUserName));
            OnPropertyChanged(nameof(AccountsEnabled));
            OnPropertyChanged(nameof(IsModeSwitchAvailable));
            OnPropertyChanged(nameof(IsAdmin));
            OnPropertyChanged(nameof(CanManageUsers));
        }

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

        public string KompasLinkStatus => IsTechnologistMode ? "🛠 Режим технолога"
            : IsViewerMode ? "👁 Режим просмотра"
            : IsLinkedToKompas ? "🔗 Связан с КОМПАС" : "⚠️ Нет связи с КОМПАС";

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
                {
                    RefreshSavedProducts();

                    // Подтягиваем новые изделия с сервера, не дожидаясь ручной синхронизации
                    if (DateTime.UtcNow - _lastSyncUtc > AutoSyncInterval)
                        RunSafe(RunServerSyncAsync(interactive: false));
                }
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
                    if (value != null)
                    {
                        _selectedTubularProduct = null;
                        _selectedOtherMaterial = null;
                    }
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
                    if (value != null)
                    {
                        _selectedSheetMaterial = null;
                        _selectedOtherMaterial = null;
                    }
                    OnMaterialFilterChanged();
                }
            }
        }

        public MaterialInfo SelectedOtherMaterial
        {
            get => _selectedOtherMaterial;
            set
            {
                if (SetProperty(ref _selectedOtherMaterial, value, nameof(SelectedOtherMaterial)))
                {
                    if (value != null)
                    {
                        _selectedSheetMaterial = null;
                        _selectedTubularProduct = null;
                    }
                    OnMaterialFilterChanged();
                }
            }
        }

        public MaterialInfo SelectedMaterialFilter => SelectedSheetMaterial ?? SelectedTubularProduct ?? SelectedOtherMaterial;

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
                {
                    ((RelayCommand)ShowInKompasCommand)?.NotifyCanExecuteChanged();
                    ((RelayCommand)EditOperationsCommand)?.NotifyCanExecuteChanged();
                }
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

        public SnackbarKind SnackbarKind
        {
            get => _snackbarKind;
            set => SetProperty(ref _snackbarKind, value, nameof(SnackbarKind));
        }

        // Счётчики для заголовков списков
        public int DetailsVisibleCount => DetailsView?.Cast<object>().Count() ?? 0;
        public int SheetMaterialsCount => SheetMaterials?.Count ?? 0;
        public int TubularProductsCount => TubularProducts?.Count ?? 0;
        public int OtherMaterialsCount => OtherMaterials?.Count ?? 0;
        public int StandardPartsCount => StandardParts?.Count ?? 0;
        // Пустой Product() — заглушка до первой загрузки
        public bool HasProduct => CurrentProduct != null
            && (!string.IsNullOrEmpty(CurrentProduct.Name) || (Details?.Count ?? 0) > 0);

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
                    OnPropertyChanged(nameof(CanManageUsers));
                    OnPropertyChanged(nameof(CanDeleteFromServer));
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
        /// Показывать «удалить везде»: только конструктору и только при заданной серверной папке
        /// </summary>
        public bool CanDeleteFromServer => IsEngineerMode && HasServerStorageFolder;

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
        public ICommand EditOperationsCommand { get; private set; }
        public ICommand OpenUsersCommand { get; private set; }
        public ICommand OpenAuditLogCommand { get; private set; }
        public ICommand OpenProductAuditLogCommand { get; private set; }

        #endregion

        #region Constructors

        public MainViewModel() : this(new KompasService()) { }

        public MainViewModel(IKompasService kompasService)
        {
            _kompasService = kompasService ?? throw new ArgumentNullException(nameof(kompasService));
            // Роль — из общего списка сотрудников; без списка — по переключателю режима, как раньше
            var users = new UserDirectoryService(_storageService.ServerStorageFolder).LoadForStartup();
            AppMode.Initialize(CurrentUser.Initialize(users, _storageService.AdminLogins, _storageService.ModeSetting));
            _logger.LogInfo($"Сотрудник: {CurrentUser.Login} ({CurrentUser.DisplayName}), список сотрудников: {(CurrentUser.AccountsEnabled ? "есть" : "нет")}, администратор: {CurrentUser.IsAdmin}");
            _logger.LogInfo($"Режим работы: {(AppMode.IsTechnologist ? "технолог" : AppMode.IsViewer ? "просмотр" : "конструктор")} (настройка: {AppMode.Setting}, КОМПАС установлен: {AppMode.IsKompasInstalled})");

            // Локальная копия расценок; общие с сервера подтянутся при синхронизации (OnWindowLoaded)
            _pricingSettings = PricingSettings.Load();
            
            SavedProducts = new ObservableCollection<ProductFileInfo>();
            CurrentProduct = new Product();

            InitializeCommands();
        }

        private void InitializeCommands()
        {
            // Команды КОМПАС, сохранения и удаления с сервера в режиме просмотра недоступны
            ShowInKompasCommand = new RelayCommand(ShowDetailInKompas, () => IsEngineerMode && CurrentlySelectedPart != null && IsLinkedToKompas && !IsLoading);
            LoadFromActiveDocumentCommand = new RelayCommand(async () => await LoadFromActiveDocumentAsync(), () => IsEngineerMode && !IsLoading);
            ClearSearchCommand = new RelayCommand(() => SearchText = string.Empty);
            LoadProductCommand = new RelayCommand<string>(LoadProduct);
            DeleteProductCommand = new RelayCommand(async () => await DeleteSelectedProductAsync(everywhere: true), () => IsEngineerMode && SelectedSavedProduct != null && !IsLoading);
            DeleteProductLocalCommand = new RelayCommand(async () => await DeleteSelectedProductAsync(everywhere: false), () => SelectedSavedProduct != null && !IsLoading);
            DeleteProductEverywhereCommand = new RelayCommand(async () => await DeleteSelectedProductAsync(everywhere: true), () => IsEngineerMode && SelectedSavedProduct != null && !IsLoading);
            ToggleProductsPanelCommand = new RelayCommand(() => IsProductsPanelOpen = !IsProductsPanelOpen);
            SwitchToProductCommand = new RelayCommand<ProductFileInfo>(SwitchToProduct);
            CopyAllToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyPartsToClipboard, Details), () => Details?.Any() == true);
            CopySheetToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyMaterialsToClipboard, SheetMaterials), () => SheetMaterials?.Any() == true);
            CopyTubularProductsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyTubularProductsToClipboard, TubularProducts), () => TubularProducts?.Any() == true);
            CopyStandartPartsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyPartsToClipboard, StandardParts), () => StandardParts?.Any() == true);
            CopyOtherMaterialsToClipboardCommand = new RelayCommand(() => CopyToClipboard(_excelService.CopyMaterialsToClipboard, OtherMaterials), () => OtherMaterials?.Any() == true);
            CopyAllDataToClipboardCommand = new RelayCommand(CopyAllDataToClipboard, () => StandardParts?.Any() == true || SheetMaterials?.Any() == true || TubularProducts?.Any() == true || OtherMaterials?.Any() == true);
            CheckForUpdatesCommand = new RelayCommand(() => UpdateService.CheckForUpdates(showNoUpdateMessage: true));
            LinkToKompasCommand = new RelayCommand(async () => await LinkToKompasAsync(), () => IsEngineerMode && !IsLinkedToKompas && !string.IsNullOrEmpty(CurrentProduct?.FilePath) && !IsLoading);
            SaveProductCommand = new RelayCommand(async () => await SaveProductAsync(), () => IsEngineerMode && CurrentProduct != null && !string.IsNullOrEmpty(CurrentProduct.Name) && IsLinkedToKompas && !IsLoading);
            RefreshFromKompasCommand = new RelayCommand(async () => await RefreshFromKompasAsync(), () => IsEngineerMode && IsLinkedToKompas && !string.IsNullOrEmpty(CurrentProduct?.FilePath) && !IsLoading);
            SelectServerStorageFolderCommand = new RelayCommand(SelectServerStorageFolder);
            ClearServerStorageFolderCommand = new RelayCommand(ClearServerStorageFolder, () => HasServerStorageFolder);
            SyncFromServerCommand = new RelayCommand(async () => await SyncFromServerAsync(), () => IsServerAvailable && !IsLoading);
            ExportToExcelCommand = new RelayCommand(ExportToExcel, () => Details?.Any() == true || StandardParts?.Any() == true || SheetMaterials?.Any() == true || TubularProducts?.Any() == true || OtherMaterials?.Any() == true);
            OpenPricingSettingsCommand = new RelayCommand(OpenPricingSettings);
            // Правки сохраняются сразу в отдельный файл (operations.json), связь с КОМПАС не нужна
            EditOperationsCommand = new RelayCommand(async () => await EditOperationsAsync(), () => CanEditOperations && CurrentlySelectedPart != null && !IsLoading);
            OpenUsersCommand = new RelayCommand(OpenUsers, () => CanManageUsers);
            OpenAuditLogCommand = new RelayCommand(() => OpenAuditLog(null));
            OpenProductAuditLogCommand = new RelayCommand<ProductFileInfo>(p => OpenAuditLog(p?.ProductName));
        }

        #endregion

        #region Pricing

        private void OpenPricingSettings()
        {
            // В режиме просмотра расценки общие и только для чтения
            var dialog = new TankManager.Views.PricingSettingsDialog(_pricingSettings, isReadOnly: IsViewerMode);
            dialog.Owner = Application.Current.MainWindow;
            if (dialog.ShowDialog() == true && IsEngineerMode)
            {
                var newSettings = dialog.PricingSettings;
                string changes = ChangeDescriber.DescribePricing(_pricingSettings, newSettings);

                // Без изменений — не перезаписываем общий файл и не подписываем
                if (changes == null)
                    return;

                newSettings.ModifiedBy = CurrentUser.Login;
                newSettings.ModifiedByName = CurrentUser.Name;
                newSettings.ModifiedUtcTicks = DateTime.UtcNow.Ticks;

                try
                {
                    // Общие расценки лежат в серверной папке, локальная копия — для работы без сети
                    string serverPath = _storageService.IsServerAvailable ? _storageService.ServerPricingFilePath : null;
                    string serverError = newSettings.Save(serverPath);

                    if (serverError != null)
                        ShowSnackbar($"Расценки сохранены только локально: {serverError}", SnackbarKind.Warning, 6000);
                    else if (serverPath == null && HasServerStorageFolder)
                        ShowSnackbar("Сервер недоступен: расценки сохранены локально и будут выложены при синхронизации", SnackbarKind.Warning, 6000);
                }
                catch (Exception ex)
                {
                    _logger.LogError("Не удалось сохранить расценки", ex);
                    MessageBox.Show(
                        $"Расценки применены, но не сохранены на диск:\n{ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                _storageService.Audit.Record(AuditAction.PricingChanged, details: changes);

                PricingSettings = newSettings;
                RecalculateAllCosts();
            }
        }

        /// <summary>
        /// Окно «Сотрудники» (администратор): роли берутся из _users.json в серверной папке
        /// </summary>
        private void OpenUsers()
        {
            if (!CanManageUsers) return;

            var directory = new UserDirectoryService(_storageService.ServerStorageFolder);
            List<UserAccount> users;
            try
            {
                users = directory.ReadServer();
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось прочитать список сотрудников: {ex.Message}");
                MessageBox.Show($"Не удалось прочитать список сотрудников с сервера:\n{ex.Message}",
                    "Сотрудники", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new TankManager.Views.UsersDialog(users, _storageService.AdminLogins.ToList());
            dialog.Owner = Application.Current.MainWindow;
            if (dialog.ShowDialog() != true || dialog.ChangedUsers.Count == 0)
                return;

            try
            {
                directory.SaveChanges(dialog.ChangedUsers);
                _storageService.Audit.Record(AuditAction.UsersChanged, details: dialog.ChangesDescription);
                ShowSnackbar("Список сотрудников сохранён. Роли применятся после перезапуска программы у сотрудников");

                // Своё ФИО могли поправить
                var me = UserDirectoryService.Find(dialog.ChangedUsers, CurrentUser.Login);
                if (me != null)
                {
                    CurrentUser.SetRegistered(me);
                    NotifyCurrentUserChanged();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Не удалось сохранить список сотрудников", ex);
                MessageBox.Show($"Не удалось сохранить список сотрудников:\n{ex.Message}",
                    "Сотрудники", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Журнал изменений (все сотрудники); productName — сразу отфильтровать по изделию
        /// </summary>
        private void OpenAuditLog(string productName)
        {
            var dialog = new TankManager.Views.AuditLogDialog(() => _storageService.Audit.ReadAll(), productName);
            dialog.Owner = Application.Current.MainWindow;
            dialog.ShowDialog();
        }

        /// <summary>
        /// Ручная правка операций выбранной детали; применяется ко всем её экземплярам в сборке
        /// </summary>
        private async Task EditOperationsAsync()
        {
            var part = CurrentlySelectedPart;
            if (part == null || CurrentProduct == null) return;

            string key = OperationEditsMerger.GetPartKey(part);
            var sameParts = (Details ?? Enumerable.Empty<PartModel>())
                .Concat(StandardParts ?? Enumerable.Empty<PartModel>())
                .Where(p => string.Equals(OperationEditsMerger.GetPartKey(p), key, StringComparison.OrdinalIgnoreCase))
                .Distinct()
                .ToList();
            if (!sameParts.Contains(part))
                sameParts.Add(part);

            var dialog = new TankManager.Views.OperationsEditorDialog(part, _pricingSettings, sameParts.Count);
            dialog.Owner = Application.Current.MainWindow;
            if (dialog.ShowDialog() != true)
                return;

            // Для журнала: операции до правки
            var operationsBefore = part.Operations.Select(op => op.Clone()).ToList();

            foreach (var target in sameParts)
            {
                target.Operations.Clear();
                foreach (var op in dialog.Operations)
                    target.Operations.Add(op.Clone());

                // У деталей из тел подписка на Operations не пересчитывает стоимость сама
                target.RecalculateOperationsCost();
            }

            RecalculateAllCosts();

            // Правки пишутся сразу, отдельно от product.json: их может делать технолог без права
            // сохранять изделие, и пересохранение изделия конструктором их не затирает
            var product = CurrentProduct;
            var operations = dialog.Operations.Select(op => op.Clone()).ToList();
            string changes = ChangeDescriber.DescribeOperations(operationsBefore, part.Operations);
            try
            {
                string serverError = await Task.Run(() => _storageService.SaveOperationEdits(product, part, operations));

                foreach (var target in sameParts)
                    target.SetOperationsModified(CurrentUser.DisplayName, DateTime.UtcNow);

                if (changes != null)
                    _storageService.Audit.Record(AuditAction.OperationsEdited, product.Name, product.Marking,
                        part.Name, part.Marking, changes);

                if (serverError == null)
                {
                    StatusMessage = $"Операции детали «{part.Name}» сохранены";
                    ShowSnackbar($"Операции детали «{part.Name}» сохранены");
                }
                else
                {
                    StatusMessage = $"Операции детали «{part.Name}» сохранены только локально: {serverError}";
                    ShowSnackbar($"Операции сохранены только локально: {serverError}", SnackbarKind.Warning, 6000);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Ошибка сохранения правок операций", ex);
                StatusMessage = $"Ошибка сохранения операций: {ex.Message}";
                MessageBox.Show($"Правки операций применены, но не сохранены:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Пересчитать стоимость всех деталей на основе текущих расценок
        /// </summary>
        /// <param name="showWarnings">Показывать предупреждение об операциях без исходных данных</param>
        public void RecalculateAllCosts(bool showWarnings = true)
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

                    if (op is LaserCuttingOperation laser)
                        laser.MaterialThickness = part.SheetThickness;

                    op.CalculateCost(_pricingSettings);
                    if (!op.IsCostReliable)
                        unreliableOperations++;
                }

                part.RecalculateOperationsCost();
            }

            CurrentProduct?.NotifyAggregatesChanged();

            if (unreliableOperations > 0 && showWarnings)
            {
                _logger.LogWarning($"Стоимость не рассчитана для операций: {unreliableOperations}");
                ShowSnackbar($"Стоимость не рассчитана для операций: {unreliableOperations} (нет исходных данных)", SnackbarKind.Warning, 6000);
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
                // Правки операций могли прийти с сервера (от технолога) после того, как изделие попало в кэш
                _storageService.ApplyOperationEdits(cachedProduct);
                SetCurrentProduct(cachedProduct, isLinked: true);
                RecalculateAllCosts(showWarnings: false);
                StatusMessage = $"{successMessage} (из кэша)";
                return;
            }

            SetCurrentProduct(savedProduct, isLinked: false);

            // Стоимость в файле посчитана по расценкам на момент сохранения — пересчитываем по текущим общим
            RecalculateAllCosts(showWarnings: false);

            StatusMessage = IsViewerMode ? successMessage : $"{successMessage} (без связи с КОМПАС)";
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
                    ShowSnackbar($"Файл не найден: {filePath}", SnackbarKind.Error);
                    return;
                }

                var linkedProduct = await Task.Run(() => _kompasService.LoadDocument(filePath));
                if (linkedProduct != null)
                {
                    CacheProduct(filePath, linkedProduct);
                    RestoreFromSaved(linkedProduct);
                    SetCurrentProduct(linkedProduct, isLinked: true);
                    RecalculateAllCosts();
                    StatusMessage = $"Связано с КОМПАС: {CurrentProduct.Name}";
                }
                else
                {
                    ShowSnackbar("Не удалось связаться с КОМПАС: убедитесь, что КОМПАС-3D запущен и документ открывается",
                        SnackbarKind.Warning, 6000);
                    StatusMessage = "Не удалось связаться с КОМПАС. Убедитесь, что КОМПАС запущен.";
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка связывания с КОМПАС: {ex.Message}");
                
                string errorMessage;

                if (ex.Message.Contains("Не удалось подключиться к KOMPAS-3D") ||
                    ex is InvalidOperationException)
                {
                    errorMessage = "Не удалось связаться с КОМПАС: запустите КОМПАС-3D и попробуйте снова";
                }
                else
                {
                    errorMessage = $"Не удалось связаться с КОМПАС: {ex.Message}";
                }

                ShowSnackbar(errorMessage, SnackbarKind.Warning, 6000);
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
                    RestoreFromSaved(refreshedProduct);
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
        /// Переносит на только что прочитанный из КОМПАС продукт данные из прежней версии:
        /// ручные правки операций и пути к изображениям
        /// </summary>
        private void RestoreFromSaved(Product kompasProduct)
        {
            if (kompasProduct?.Details == null)
                return;

            Product savedProduct = null;
            try
            {
                savedProduct = _storageService.TryLoadSavedProduct(kompasProduct);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка чтения сохранённого изделия: {ex.Message}");
            }

            RestoreOperationEdits(kompasProduct);
            RestoreImagePathsFromSaved(kompasProduct, savedProduct);

            // Кто сохранял — показываем и для изделия, открытого из КОМПАС
            if (savedProduct != null)
            {
                kompasProduct.SavedBy = savedProduct.SavedBy;
                kompasProduct.SavedByName = savedProduct.SavedByName;
                kompasProduct.SavedUtc = savedProduct.SavedUtc;
            }
        }

        /// <summary>
        /// Ручные правки операций хранятся в operations.json папки изделия (их пишут конструктор и технолог)
        /// </summary>
        private void RestoreOperationEdits(Product kompasProduct)
        {
            try
            {
                _storageService.ApplyOperationEdits(kompasProduct);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка переноса правок операций: {ex.Message}");
            }
        }

        /// <summary>
        /// Восстанавливает пути к изображениям из ранее сохранённого продукта.
        /// Позволяет избежать повторных COM-вызовов при связывании/обновлении из КОМПАС,
        /// когда PNG-файлы уже существуют на диске и актуальны.
        /// </summary>
        private void RestoreImagePathsFromSaved(Product kompasProduct, Product savedProduct)
        {
            if (savedProduct?.Details == null)
                return;

            try
            {
                // Превью сборки, если актуально
                if (!string.IsNullOrEmpty(savedProduct.FilePreviewPngPath) && File.Exists(savedProduct.FilePreviewPngPath) &&
                    (string.IsNullOrEmpty(kompasProduct.FilePath) || !File.Exists(kompasProduct.FilePath) ||
                     File.GetLastWriteTimeUtc(savedProduct.FilePreviewPngPath) >= File.GetLastWriteTimeUtc(kompasProduct.FilePath)))
                {
                    kompasProduct.FilePreviewPngPath = savedProduct.FilePreviewPngPath;
                }

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

                    // Восстанавливаем путь к превью 3D-файла, если существует и актуален.
                    // Изделиям из библиотек КОМПАС не восстанавливаем: типоразмеры делят FilePath,
                    // их превью находит по имени файла GenerateLibraryPartPreviewsAsync
                    if (!detail.IsLibraryPart &&
                        !string.IsNullOrEmpty(savedDetail.FilePreviewPngPath) && File.Exists(savedDetail.FilePreviewPngPath))
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
                    RestoreFromSaved(linkedProduct);
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

            if (IsViewerMode)
            {
                ShowSnackbar(ViewerModeLoadMessage, SnackbarKind.Warning);
                return;
            }

            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа...";

                var product = await Task.Run(() => _kompasService.LoadDocument(filePath));
                CacheProduct(filePath, product);
                RestoreFromSaved(product);
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

            if (IsViewerMode)
            {
                ShowSnackbar(ViewerModeLoadMessage, SnackbarKind.Warning);
                return;
            }

            try
            {
                IsLoading = true;
                StatusMessage = "Загрузка документа из КОМПАС...";

                var product = await Task.Run(() => _kompasService.LoadActiveDocument());

                if (!string.IsNullOrEmpty(product.FilePath))
                    CacheProduct(product.FilePath, product);

                RestoreFromSaved(product);
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
                await Task.Run(() => SaveProductPreview(product, imagesFolder));

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
                    ShowSnackbar($"Изделие сохранено только локально: {serverError}", SnackbarKind.Warning, 6000);
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
        /// Сохраняет миниатюру сборки в PNG: без КОМПАС на компьютере её не получить из файла .a3d
        /// </summary>
        private void SaveProductPreview(Product product, string imagesFolder)
        {
            var sourcePath = product?.FilePath;
            if (string.IsNullOrEmpty(imagesFolder) || string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                return;

            try
            {
                // Отдельный префикс: имя сборки может совпадать с именем детали
                string pngPath = Path.Combine(imagesFolder, ThumbnailService.GeneratePreviewFileName(sourcePath, "product"));

                bool isActual = File.Exists(pngPath) &&
                    File.GetLastWriteTimeUtc(pngPath) >= File.GetLastWriteTimeUtc(sourcePath);

                if (!isActual)
                {
                    var preview = ThumbnailService.GetFileThumbnail(sourcePath);
                    if (preview == null || !ThumbnailService.SavePreviewToFile(preview, pngPath))
                        return;
                }

                product.FilePreviewPngPath = pngPath;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось сохранить превью сборки: {ex.Message}");
            }
        }

        /// <summary>
        /// Сохраняет превью 3D-файлов для всех уникальных деталей
        /// </summary>
        private async Task SaveAllFilePreviewsAsync(string imagesFolder)
        {
            if (string.IsNullOrEmpty(imagesFolder))
                return;

            // Собираем все детали, кроме изделий из библиотек КОМПАС: их превью по вхождению
            // создаёт GenerateLibraryPartPreviewsAsync (см. PartModel.IsLibraryPart)
            var allParts = (Details ?? Enumerable.Empty<PartModel>())
                .Where(p => !p.IsLibraryPart)
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

            await GenerateDrawingPreviewsAsync(product, imagesFolder, token);
            await GenerateLibraryPartPreviewsAsync(product, imagesFolder, token);
        }

        private async Task GenerateDrawingPreviewsAsync(Product product, string imagesFolder, CancellationToken token)
        {
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

        /// <summary>
        /// Превью изделий из библиотек КОМПАС по геометрии вхождения. Типоразмеры одного
        /// шаблона (общий FilePath) различаются наименованием и обозначением.
        /// </summary>
        private async Task GenerateLibraryPartPreviewsAsync(Product product, string imagesFolder, CancellationToken token)
        {
            if (Details == null || string.IsNullOrEmpty(imagesFolder))
                return;

            var groups = Details
                .Where(p => p.IsLibraryPart && !p.IsBodyBased)
                .GroupBy(p => new { p.FilePath, p.Name, p.Marking })
                .ToList();

            var partsToProcess = new List<PartModel>();
            foreach (var group in groups)
            {
                // Превью уже есть на диске — КОМПАС не трогаем
                string cachedPath = Path.Combine(imagesFolder, StandardPartPreviewService.GetPreviewFileName(group.First()));
                if (File.Exists(cachedPath))
                {
                    foreach (var part in group)
                        part.FilePreviewPngPath = cachedPath;
                }
                else
                {
                    partsToProcess.Add(group.First());
                }
            }

            int processed = 0;
            int total = partsToProcess.Count;

            foreach (var part in partsToProcess)
            {
                if (token.IsCancellationRequested || CurrentProduct != product)
                    return;

                try
                {
                    StatusMessage = $"Генерация превью стд. изделий: {processed + 1}/{total}";

                    await Task.Run(() => _kompasService.LoadStandardPartPreview(part, product, imagesFolder), token);

                    if (!string.IsNullOrEmpty(part.FilePreviewPngPath))
                    {
                        foreach (var samePart in Details.Where(p => p != part &&
                            p.FilePath == part.FilePath && p.Name == part.Name && p.Marking == part.Marking))
                        {
                            samePart.FilePreviewPngPath = part.FilePreviewPngPath;
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
                    _logger.LogWarning($"Фоновая генерация превью стд. изделия {part.Name}: {ex.Message}");
                    processed++;
                }
            }

            if (total > 0 && CurrentProduct == product && !token.IsCancellationRequested)
            {
                StatusMessage = $"Превью стд. изделий готовы: {processed}/{total}";
            }
        }

        private async Task DeleteSelectedProductAsync(bool everywhere)
        {
            var info = SelectedSavedProduct;
            if (info == null || IsLoading) return;

            // Удалять с сервера может только конструктор
            if (everywhere && IsViewerMode) return;

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
                    ShowSnackbar(_storageService.LastServerError, SnackbarKind.Warning, 6000);
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
            NotifyCountsChanged();
        }

        private void NotifyCountsChanged()
        {
            OnPropertyChanged(nameof(DetailsVisibleCount));
            OnPropertyChanged(nameof(SheetMaterialsCount));
            OnPropertyChanged(nameof(TubularProductsCount));
            OnPropertyChanged(nameof(OtherMaterialsCount));
            OnPropertyChanged(nameof(StandardPartsCount));
            OnPropertyChanged(nameof(HasProduct));
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
            NotifyCountsChanged();
        }

        public void ClearMaterialFilter()
        {
            SelectedSheetMaterial = null;
            SelectedTubularProduct = null;
            SelectedOtherMaterial = null;
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

        /// <summary>
        /// Вызывается после показа окна: в режиме просмотра сразу открывает список изделий
        /// и в фоне подтягивает изделия и общие расценки с сервера
        /// </summary>
        public void OnWindowLoaded()
        {
            if (IsViewerMode && !HasProduct)
            {
                // Открытие панели само запускает синхронизацию
                IsProductsPanelOpen = true;

                if (!HasServerStorageFolder)
                    StatusMessage = "Укажите серверную папку с изделиями в панели «Изделия»";
            }

            RunSafe(RunServerSyncAsync(interactive: false));
            RunSafe(RegisterCurrentUserAsync());
        }

        /// <summary>
        /// Добавляет сотрудника в общий список при первом запуске (ФИО — из домена, в фоне).
        /// Первый администратор (из storage_settings) так же создаёт сам список
        /// </summary>
        private async Task RegisterCurrentUserAsync()
        {
            bool needed = CurrentUser.NeedsRegistration || (!CurrentUser.AccountsEnabled && CurrentUser.IsAdmin);
            if (!needed || !HasServerStorageFolder)
                return;

            string serverFolder = _storageService.ServerStorageFolder;
            var role = CurrentUser.RoleFromCurrentMode();
            bool isAdmin = CurrentUser.IsAdmin;

            try
            {
                bool created = false;
                var account = await Task.Run(() => new UserDirectoryService(serverFolder).Register(CurrentUser.Login, role, isAdmin, out created));
                if (account == null)
                    return;

                CurrentUser.SetRegistered(account);
                NotifyCurrentUserChanged();

                if (created)
                {
                    _storageService.Audit.Record(AuditAction.UserRegistered,
                        details: $"{account.Login} — {account.DisplayName}, роль: {UserAccount.RoleTitle(account.Role)}{(account.IsAdmin ? ", администратор" : "")}");
                    _logger.LogInfo($"Сотрудник {account.Login} добавлен в общий список");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось добавить сотрудника в общий список: {ex.Message}");
            }
        }

        private Task SyncFromServerAsync() => RunServerSyncAsync(interactive: true);

        /// <summary>
        /// Синхронизация изделий и общих расценок с сервером.
        /// Ручная (interactive) блокирует интерфейс и всегда пишет итог в строку состояния;
        /// фоновая сообщает только об изменениях и ошибках.
        /// </summary>
        private async Task RunServerSyncAsync(bool interactive)
        {
            if (!HasServerStorageFolder) return;
            if (interactive && IsLoading) return;

            if (_isSyncing)
            {
                if (interactive)
                    StatusMessage = "Синхронизация уже выполняется...";
                return;
            }

            _isSyncing = true;
            _lastSyncUtc = DateTime.UtcNow;

            try
            {
                if (interactive)
                {
                    IsLoading = true;
                    StatusMessage = IsViewerMode ? "Загрузка изделий с сервера..." : "Двусторонняя синхронизация...";
                }

                bool downloadOnly = IsViewerMode;
                string pricingPath = _storageService.ServerPricingFilePath;

                // Сетевые операции (в том числе проверка доступности сервера) — вне UI-потока
                var result = await Task.Run(() =>
                {
                    var sync = _storageService.SyncFromServer(skipImages: true);
                    var pricing = _storageService.IsServerAvailable
                        ? PricingSettings.SyncWithServer(pricingPath, downloadOnly)
                        : null;
                    return Tuple.Create(sync, pricing);
                });

                var syncResult = result.Item1;
                bool hasChanges = syncResult.NewProducts > 0 || syncResult.UpdatedProducts > 0;

                if (!syncResult.Success)
                {
                    StatusMessage = IsServerAvailable
                        ? $"Синхронизация с ошибками: {string.Join(", ", syncResult.Errors.Take(2))}"
                        : "Сервер недоступен: показаны локальные копии изделий";
                }
                else if (hasChanges)
                {
                    StatusMessage = $"Синхронизация завершена: новых {syncResult.NewProducts}, обновлено {syncResult.UpdatedProducts}";
                }
                else if (interactive)
                {
                    StatusMessage = "Синхронизация: данные актуальны";
                }

                if (result.Item2 != null)
                {
                    PricingSettings = result.Item2;
                    if (HasProduct)
                        RecalculateAllCosts(showWarnings: false);
                }

                if (interactive || hasChanges || IsProductsPanelOpen)
                    RefreshSavedProducts();
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка синхронизации: {ex.Message}");
                StatusMessage = $"Ошибка синхронизации: {ex.Message}";
            }
            finally
            {
                _isSyncing = false;
                if (interactive)
                    IsLoading = false;

                OnPropertyChanged(nameof(IsServerAvailable));
                ((RelayCommand)SyncFromServerCommand)?.NotifyCanExecuteChanged();
            }
        }

        /// <summary>
        /// Фоновая синхронизация изображений продукта с сервером
        /// </summary>
        private async Task SyncProductImagesInBackgroundAsync(Product product)
        {
            if (product == null || string.IsNullOrEmpty(product.Name) || !HasServerStorageFolder)
                return;

            try
            {
                int downloaded = await Task.Run(() => _storageService.SyncProductImagesWithServer(product));

                // Картинки пришли после того, как превью уже запрашивались — перечитываем
                if (downloaded > 0 && CurrentProduct == product)
                    InvalidatePreviews(product);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка фоновой синхронизации изображений: {ex.Message}");
            }
        }

        private static void InvalidatePreviews(Product product)
        {
            product.InvalidateFilePreviewCache();

            var parts = (product.Details ?? Enumerable.Empty<PartModel>())
                .Concat(product.StandardParts ?? Enumerable.Empty<PartModel>());

            foreach (var part in parts)
            {
                part.InvalidateFilePreviewCache();
                part.InvalidateDrawingPreviewCache();
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
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                _excelService.OpenInExcel(
                    CurrentProduct?.Name,
                    Details,
                    StandardParts,
                    SheetMaterials,
                    TubularProducts,
                    OtherMaterials);

                StatusMessage = "Ведомость открыта в Excel";
                ShowSnackbar("Ведомость материалов открыта в Excel");
            }
            catch (Exception ex)
            {
                _logger.LogError("Ошибка экспорта в Excel", ex);
                ShowSnackbar($"Ошибка при экспорте в Excel: {ex.Message}", SnackbarKind.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        #endregion

        #region Helper Methods

        public void ShowSnackbar(string message, SnackbarKind kind = SnackbarKind.Success, int durationMs = 0)
        {
            // Останавливаем предыдущий таймер, если есть
            _snackbarTimer?.Dispose();

            if (durationMs <= 0)
                durationMs = kind == SnackbarKind.Error ? 6000
                           : kind == SnackbarKind.Warning ? 5000
                           : 3000;

            SnackbarKind = kind;
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
            ((RelayCommand)EditOperationsCommand)?.NotifyCanExecuteChanged();
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
            SelectedOtherMaterial = null;
            CurrentlySelectedPart = null;
            IsProductSelected = false;
        }

        private void OnMaterialFilterChanged()
        {
            OnPropertyChanged(nameof(SelectedSheetMaterial));
            OnPropertyChanged(nameof(SelectedTubularProduct));
            OnPropertyChanged(nameof(SelectedOtherMaterial));
            OnPropertyChanged(nameof(SelectedMaterialFilter));
            DetailsView?.Refresh();
            NotifyCountsChanged();
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
            ((RelayCommand)EditOperationsCommand)?.NotifyCanExecuteChanged();
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
            ShowSnackbar("Связь с КОМПАС потеряна", SnackbarKind.Warning);
            NotifyCopyCommandsCanExecuteChanged();
            NotifySaveCommandCanExecuteChanged();
            NotifyLinkCommandCanExecuteChanged();
            NotifyRefreshCommandCanExecuteChanged();
        }

        private void ShowError(string message, Exception ex)
        {
            StatusMessage = $"Ошибка: {ex.Message}";
            ShowSnackbar($"{message}: {ex.Message}", SnackbarKind.Error);
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

    /// <summary>
    /// Вид уведомления SnackBar
    /// </summary>
    public enum SnackbarKind
    {
        Success,
        Info,
        Warning,
        Error
    }
}
