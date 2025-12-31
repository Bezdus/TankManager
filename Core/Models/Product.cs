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
        private BitmapSource _filePreview;
        private bool _previewLoaded;

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

        /// <summary>
        /// Проверяет, связан ли продукт с активным документом KOMPAS
        /// </summary>
        public bool IsLinkedToKompas => Context?.IsDocumentLoaded == true && Context.TopPart != null;

        /// <summary>
        /// Превью файла изделия с ленивой загрузкой
        /// </summary>
        public new BitmapSource FilePreview
        {
            get
            {
                if (!_previewLoaded)
                {
                    _previewLoaded = true;
                    _filePreview = TryLoadPreview(FilePath);
                }
                return _filePreview;
            }
            set
            {
                if (_filePreview != value)
                {
                    _filePreview = value;
                    _previewLoaded = true;
                    OnPropertyChanged(nameof(FilePreview));
                }
            }
        }

        private static BitmapSource TryLoadPreview(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                return null;

            try
            {
                return ThumbnailService.GetFileThumbnail(filePath);
            }
            catch
            {
                return null;
            }
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
            Context = null;
            
            // Освобождаем собственные ресурсы
            _filePreview = null;
            _previewLoaded = false;
            InvalidateDrawingPreviewCache();
        }
    }
}