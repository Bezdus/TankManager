using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media.Imaging;
using KompasAPI7;
using TankManager.Core.Services;

namespace TankManager.Core.Models
{
    /// <summary>
    /// Представляет изделие (сборку) с его составными частями
    /// </summary>
    public class Product : PartModel
    {
        /// <summary>
        /// Контекст KOMPАС, связанный с этим продуктом
        /// </summary>
        public KompasContext Context { get; private set; }

        public Product() : base()
        {
            Details = new ObservableCollection<PartModel>();
            StandardParts = new ObservableCollection<PartModel>();
            SheetMaterials = new ObservableCollection<MaterialInfo>();
            TubularProducts = new ObservableCollection<MaterialInfo>();
            OtherMaterials = new ObservableCollection<MaterialInfo>();
        }

        public Product(IPart7 part, KompasContext context, int instanceIndex = 0) 
            : base(part, context, instanceIndex)
        {
            Context = context;
            Details = new ObservableCollection<PartModel>();
            StandardParts = new ObservableCollection<PartModel>();
            SheetMaterials = new ObservableCollection<MaterialInfo>();
            TubularProducts = new ObservableCollection<MaterialInfo>();
            OtherMaterials = new ObservableCollection<MaterialInfo>();
        }

        /// <summary>
        /// Все детали изделия
        /// </summary>
        public ObservableCollection<PartModel> Details { get; }

        /// <summary>
        /// Покупные детали (стандартные изделия)
        /// </summary>
        public ObservableCollection<PartModel> StandardParts { get; }

        /// <summary>
        /// Листовой прокат (материал -> суммарная масса)
        /// </summary>
        public ObservableCollection<MaterialInfo> SheetMaterials { get; }

        /// <summary>
        /// Трубный прокат (материал -> суммарная масса)
        /// </summary>
        public ObservableCollection<MaterialInfo> TubularProducts { get; }

        /// <summary>
        /// Трубный прокат (материал -> суммарная масса)
        /// </summary>
        public ObservableCollection<MaterialInfo> OtherMaterials { get; }

        private string _savedBy;
        private string _savedByName;
        private DateTime? _savedUtc;

        /// <summary>
        /// Логин сотрудника, сохранившего изделие (null — изделие не сохранялось или сохранено старой версией)
        /// </summary>
        public string SavedBy
        {
            get => _savedBy;
            set { _savedBy = value; NotifySavedInfoChanged(); }
        }

        public string SavedByName
        {
            get => _savedByName;
            set { _savedByName = value; NotifySavedInfoChanged(); }
        }

        public DateTime? SavedUtc
        {
            get => _savedUtc;
            set { _savedUtc = value; NotifySavedInfoChanged(); }
        }

        public string SavedByDisplay => string.IsNullOrWhiteSpace(SavedByName) ? SavedBy : SavedByName;

        /// <summary>
        /// «Сохранил: ФИО, дата» для шапки изделия (null — не известно)
        /// </summary>
        public string SavedInfo => string.IsNullOrEmpty(SavedByDisplay)
            ? null
            : SavedUtc.HasValue
                ? $"Сохранил: {SavedByDisplay}, {SavedUtc.Value.ToLocalTime():dd.MM.yyyy HH:mm}"
                : $"Сохранил: {SavedByDisplay}";

        private bool _hasUnsavedChanges;

        /// <summary>
        /// Загружено из КОМПАС и не сохранено: сохранения нет или сборка изменена после него
        /// </summary>
        public bool HasUnsavedChanges
        {
            get => _hasUnsavedChanges;
            set
            {
                if (_hasUnsavedChanges == value) return;
                _hasUnsavedChanges = value;
                OnPropertyChanged(nameof(HasUnsavedChanges));
            }
        }

        private void NotifySavedInfoChanged()
        {
            OnPropertyChanged(nameof(SavedBy));
            OnPropertyChanged(nameof(SavedByName));
            OnPropertyChanged(nameof(SavedUtc));
            OnPropertyChanged(nameof(SavedByDisplay));
            OnPropertyChanged(nameof(SavedInfo));
        }

        /// <summary>
        /// Общее количество деталей в изделии
        /// </summary>
        public int TotalPartsCount => AllParts.Count();

        /// <summary>
        /// Суммарная стоимость всех деталей изделия, руб
        /// </summary>
        public double TotalAssemblyCost => AllParts.Sum(p => p.TotalCost);

        /// <summary>
        /// Суммарная стоимость металла, руб
        /// </summary>
        public double TotalMetalCost => AllParts.Sum(p => p.MetalCost);

        /// <summary>
        /// Суммарная стоимость операций изготовления, руб
        /// </summary>
        public double TotalOperationsCost => AllParts.Sum(p => p.OperationsCost);

        /// <summary>
        /// Все детали без повторов. Покупные детали лежат и в Details, и в StandardParts
        /// (KompasService.ExtractAllParts); StandardParts добавляем, только если в Details
        /// покупных нет (сохранения, где покупные хранились отдельно)
        /// </summary>
        public IEnumerable<PartModel> AllParts =>
            Details.Any(p => p.ProductType == ProductType.PurchasedPart)
                ? Details
                : Details.Concat(StandardParts);

        /// <summary>
        /// Уведомляет UI об изменении агрегированных свойств (количество, стоимость)
        /// </summary>
        public void NotifyAggregatesChanged()
        {
            OnPropertyChanged(nameof(TotalPartsCount));
            OnPropertyChanged(nameof(TotalAssemblyCost));
            OnPropertyChanged(nameof(TotalMetalCost));
            OnPropertyChanged(nameof(TotalOperationsCost));
        }

        /// <summary>
        /// Проверяет, связан ли продукт с активным документом KOMPAS
        /// </summary>
        public bool IsLinkedToKompas => Context?.IsDocumentLoaded == true && Context.TopPart != null;

        // Превью изделия — FilePreview базового класса: миниатюра сборки или сохранённый
        // PNG (FilePreviewPngPath), который видят пользователи без КОМПАС

        protected override void Dispose(bool disposing)
        {
            if (disposing && Context != null)
            {
                Context.Dispose();
                Context = null;
            }

            base.Dispose(disposing);
        }

        public void Clear()
        {
            // Очищаем превью у всех деталей перед очисткой коллекций
            foreach (var detail in Details)
            {
                detail.FilePreview = null;
                detail.InvalidateDrawingPreviewCache();
                detail.Dispose();
            }
            
            foreach (var part in StandardParts)
            {
                part.FilePreview = null;
                part.InvalidateDrawingPreviewCache();
                part.Dispose();
            }
            
            Details.Clear();
            StandardParts.Clear();
            SheetMaterials.Clear();
            TubularProducts.Clear();
            OtherMaterials.Clear();
            Name = null;
            Marking = null;
            Mass = 0;
            SavedBy = null;
            SavedByName = null;
            SavedUtc = null;
            Context?.Dispose();
            Context = null;
            
            // Освобождаем собственные ресурсы
            FilePreviewPngPath = null;
            InvalidateFilePreviewCache();
            InvalidateDrawingPreviewCache();
            NotifyAggregatesChanged();
        }
    }
}