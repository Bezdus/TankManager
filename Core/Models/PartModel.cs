using Kompas6API5;
using KompasAPI7;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;
using TankManager.Core.Constants;
using TankManager.Core.Services;

namespace TankManager.Core.Models
{
    public class PartModel : INotifyPropertyChanged, IDisposable
    {
        //private const string DefaultSteelGrade = "AISI 304";

        private string _name;
        private string _marking;
        private string _detailType;
        private string _material;
        private double _mass;
        private string _filePath;
        private BitmapSource _filePreview;
        private bool _disposed;
        private bool _previewLoaded;
        private ProductType _productType;
        private double _length;
        private string _pngFilePath;
        private string _cdwFilePath; // Путь к исходному файлу чертежа
        private BitmapSource _drawingPreview;
        private bool _drawingPreviewLoaded;
        private string _filePreviewPngPath; // Путь к сохранённому превью 3D-файла
        private string _dxfFilePath; // Путь к DXF-файлу для лазерной резки

        private static readonly DrawingPreviewService _previewService = new DrawingPreviewService();

        /// <summary>
        /// Операции изготовления детали
        /// </summary>
        public ObservableCollection<ManufacturingOperationBase> Operations { get; }
            = new ObservableCollection<ManufacturingOperationBase>();

        /// <summary>
        /// Есть ли операции изготовления
        /// </summary>
        public bool HasOperations => Operations.Count > 0;

        // Уникальные идентификаторы для поиска в KOMPAS
        public string PartId { get; private set; }
        public bool IsBodyBased { get; private set; }
        public int InstanceIndex { get; private set; } // Индекс экземпляра в сборке

        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        public string Marking
        {
            get { return _marking; }
            set
            {
                if (_marking != value)
                {
                    _marking = value;
                    OnPropertyChanged(nameof(Marking));
                }
            }
        }

        public string DetailType
        {
            get { return _detailType; }
            set
            {
                if (_detailType != value)
                {
                    _detailType = value;
                    OnPropertyChanged(nameof(DetailType));
                }
            }
        }

        public string Material
        {
            get { return _material; }
            set
            {
                if (_material != value)
                {
                    _material = value;
                    OnPropertyChanged(nameof(Material));
                }
            }
        }

        public double Mass
        {
            get { return _mass; }
            set
            {
                if (Math.Abs(_mass - value) > 0.0001)
                {
                    _mass = value;
                    OnPropertyChanged(nameof(Mass));
                }
            }
        }

        public double Length
        {
            get { return _length; }
            set
            {
                if (Math.Abs(_length - value) > 0.0001)
                {
                    _length = value;
                    OnPropertyChanged(nameof(Mass));
                }
            }
        }

        public string FilePath
        {
            get { return _filePath; }
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    OnPropertyChanged(nameof(FilePath));
                }
            }
        }

        public string CdfFilePath
        {
            get { return _pngFilePath; }
            set
            {
                if (_pngFilePath != value)
                {
                    _pngFilePath = value;
                    _drawingPreviewLoaded = false; // Сбрасываем флаг загрузки при изменении пути
                    OnPropertyChanged(nameof(CdfFilePath));
                    OnPropertyChanged(nameof(DrawingPreview));
                }
            }
        }

        /// <summary>
        /// Путь к исходному файлу чертежа (.cdw) для проверки актуальности кэша
        /// </summary>
        public string SourceCdwPath
        {
            get { return _cdwFilePath; }
            set
            {
                if (_cdwFilePath != value)
                {
                    _cdwFilePath = value;
                    OnPropertyChanged(nameof(SourceCdwPath));
                }
            }
        }

        /// <summary>
        /// Путь к DXF-файлу для лазерной резки
        /// </summary>
        public string DxfFilePath
        {
            get { return _dxfFilePath; }
            set
            {
                if (_dxfFilePath != value)
                {
                    _dxfFilePath = value;
                    OnPropertyChanged(nameof(DxfFilePath));
                }
            }
        }

        /// <summary>
        /// Путь к сохранённому PNG-превью 3D-файла (для работы без исходных файлов КОМПАС)
        /// </summary>
        public string FilePreviewPngPath
        {
            get { return _filePreviewPngPath; }
            set
            {
                if (_filePreviewPngPath != value)
                {
                    _filePreviewPngPath = value;
                    _previewLoaded = false; // Сбрасываем флаг загрузки при изменении пути
                    OnPropertyChanged(nameof(FilePreviewPngPath));
                    OnPropertyChanged(nameof(FilePreview));
                }
            }
        }

        /// <summary>
        /// Превью чертежа для отображения в UI
        /// </summary>
        public BitmapSource DrawingPreview
        {
            get
            {
                if (!_drawingPreviewLoaded)
                {
                    _drawingPreviewLoaded = true;
                    _drawingPreview = _previewService.LoadPreviewImage(_pngFilePath, _cdwFilePath);
                }
                return _drawingPreview;
            }
        }

        /// <summary>
        /// Сбрасывает кэш превью чертежа для принудительной перепроверки актуальности
        /// </summary>
        public void InvalidateDrawingPreviewCache()
        {
            _drawingPreviewLoaded = false;
            _drawingPreview = null;
            OnPropertyChanged(nameof(DrawingPreview));
        }

        public ProductType ProductType
        {
            get { return _productType; }
            set
            {
                if (_productType != value)
                {
                    _productType = value;
                    OnPropertyChanged(nameof(ProductType));
                }
            }
        }

        public BitmapSource FilePreview
        {
            get
            {
                if (!_previewLoaded)
                {
                    _previewLoaded = true;
                    _filePreview = TryLoadPreview(FilePath, _filePreviewPngPath);
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

        /// <summary>
        /// Защищённый конструктор для наследников
        /// </summary>
        protected PartModel()
        {
            _name = string.Empty;
            _marking = string.Empty;
            _material = string.Empty;
            _filePath = string.Empty;
            _productType = ProductType.Part;
            Operations.CollectionChanged += (s, e) => OnPropertyChanged(nameof(HasOperations));
        }

        public PartModel(IPart7 part, KompasContext context, int instanceIndex = 0)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));
            if (context == null) throw new ArgumentNullException(nameof(context));

            IsBodyBased = false;
            InstanceIndex = instanceIndex;
            PartId = $"{part.Name}|{part.Marking}|{part.FileName}|{instanceIndex}";
            Name = part.Name ?? string.Empty;
            Marking = part.Marking ?? string.Empty;
            DetailType = DetermineDetailType(context.GetDetailType(part));
            Material = FormatMaterial(part.Material);
            Mass = part.Mass / KompasConstants.MassConversionFactor;
            FilePath = part.FileName ?? string.Empty;
            ProductType = DetermineProductType(DetailType, Material);

            if (ProductType == ProductType.TubularProduct)
                Length = GetLength(part, context);
            else
                Length = -1;

            Operations.CollectionChanged += (s, e) => OnPropertyChanged(nameof(HasOperations));

        }

        public PartModel(IBody7 body, KompasContext context, int instanceIndex = 0)
        {
            if (body == null) throw new ArgumentNullException(nameof(body));
            if (context == null) throw new ArgumentNullException(nameof(context));

            IsBodyBased = true;
            InstanceIndex = instanceIndex;
            IPart7 parentPart = null;
            

            try
            {
                Name = body.Name ?? string.Empty;
                Marking = body.Marking ?? string.Empty;
                DetailType = KompasConstants.PartType;
                Material = FormatMaterial(
                    context.GetBodyPropertyValue(body, KompasConstants.MaterialPropertyName));

                Mass = ParseMass(
                    context.GetBodyPropertyValue(body, KompasConstants.MassPropertyName));
                
                parentPart = body.Parent as IPart7;
                string parentFileName = parentPart?.FileName ?? string.Empty;
                string parentName = parentPart?.Name ?? string.Empty;
                PartId = $"{parentName}|{Name}|{Marking}|{instanceIndex}";
                FilePath = parentFileName;
                ProductType = DetermineProductType(DetailType, Material);

                if (ProductType == ProductType.TubularProduct)
                    Length = GetLength(body, context);
                else
                    Length = -1;

                Operations.CollectionChanged += (s, e) => OnPropertyChanged(nameof(HasOperations));

            }
            finally
            {
                // Не освобождаем parentPart, т.к. это приведение типа
            }
        }


        /// <summary>
        /// Загружает превью чертежа для детали
        /// </summary>
        /// <param name="targetDirectory">Целевая папка для сохранения превью</param>
        public void LoadDrawingPreview(IPart7 part, KompasContext context, string targetDirectory)
        {
            if (string.IsNullOrEmpty(targetDirectory))
                return;

            string sourceCdwPath;
            string pngPath = _previewService.GetOrCreatePreview(part, context, out sourceCdwPath, targetDirectory);
            
            // Сначала устанавливаем путь к исходному файлу
            _cdwFilePath = sourceCdwPath;
            
            // Принудительно сбрасываем кэш изображения перед установкой пути
            _drawingPreviewLoaded = false;
            _drawingPreview = null;
            
            // Устанавливаем путь к PNG (напрямую, чтобы избежать проверки на равенство)
            _pngFilePath = pngPath;
            
            // Уведомляем UI об изменениях
            OnPropertyChanged(nameof(CdfFilePath));
            OnPropertyChanged(nameof(DrawingPreview));
        }

        /// <summary>
        /// Сохраняет текущее превью 3D-файла в PNG для работы без исходных файлов КОМПАС
        /// </summary>
        /// <param name="targetDirectory">Целевая папка для сохранения (обычно images)</param>
        /// <returns>Путь к сохранённому файлу или null если не удалось сохранить</returns>
        public string SaveFilePreview(string targetDirectory)
        {
            // Если превью уже сохранено в памяти, файл существует и актуален — возвращаем путь
            if (!string.IsNullOrEmpty(_filePreviewPngPath) && File.Exists(_filePreviewPngPath))
            {
                if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath) ||
                    File.GetLastWriteTimeUtc(_filePreviewPngPath) >= File.GetLastWriteTimeUtc(FilePath))
                    return _filePreviewPngPath;
            }

            if (string.IsNullOrEmpty(targetDirectory))
                return null;

            try
            {
                string fileName = ThumbnailService.GeneratePreviewFileName(FilePath, "file");
                string filePath = Path.Combine(targetDirectory, fileName);

                // Если файл уже существует на диске и актуален — просто запоминаем путь, не перегенерируем
                if (File.Exists(filePath))
                {
                    if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath) ||
                        File.GetLastWriteTimeUtc(filePath) >= File.GetLastWriteTimeUtc(FilePath))
                    {
                        _filePreviewPngPath = filePath;
                        OnPropertyChanged(nameof(FilePreviewPngPath));
                        return filePath;
                    }
                }

                // Только если файла нет на диске — загружаем превью (дорогая операция)
                var preview = FilePreview;
                if (preview == null)
                    return null;

                if (ThumbnailService.SavePreviewToFile(preview, filePath))
                {
                    _filePreviewPngPath = filePath;
                    OnPropertyChanged(nameof(FilePreviewPngPath));
                    return filePath;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Статический доступ к сервису превью (для очистки кэша и т.д.)
        /// </summary>
        public static DrawingPreviewService PreviewService => _previewService;

        private  double GetLength(object detail, KompasContext context)
        {
            if (detail == null || context == null)
                return 0;

            try
            {
                if (detail is IBody7 body)
                {
                    string lengthValue = context.GetBodyPropertyValue(body, "Длина профиля");
                    return ParseMass(lengthValue);
                }

                if (detail is IPart7 iPart)
                {
                    return context.GetDetailLengthByExtrusion(iPart);
                }
            }
            catch
            {
                return 0;
            }

            return 0;
        }

        private static string DetermineDetailType(string specificationSection)
        {
            if (specificationSection == KompasConstants.StandardPartsType ||
                specificationSection == KompasConstants.OtherPartsType)
            {
                return KompasConstants.PurchasedPartType;
            }
            return KompasConstants.PartType;
        }

        private static ProductType DetermineProductType(string detailType, string material)
        {
            // Покупная деталь
            if (detailType == KompasConstants.PurchasedPartType)
            {
                return ProductType.PurchasedPart;
            }

            // Определяем по материалу
            if (!string.IsNullOrWhiteSpace(material))
            {
                string materialLower = material.ToLowerInvariant();

                // Трубный прокат
                if (materialLower.Contains("труба") || 
                    materialLower.Contains("труб") ||
                    materialLower.Contains("круг") ||
                    materialLower.Contains("уголок") ||
                    materialLower.Contains("стержень"))
                {
                    return ProductType.TubularProduct;
                }

                // Листовой прокат
                if (materialLower.Contains("лист") || 
                    materialLower.Contains("полоса") ||
                    materialLower.Contains("рулон"))
                {
                    return ProductType.SheetMaterial;
                }
            }

            return ProductType.Part;
        }

        private static double ParseMass(string massString)
        {
            if (string.IsNullOrWhiteSpace(massString))
                return 0;

            string normalized = massString.Replace(',', '.');
            return double.TryParse(normalized, NumberStyles.Float,
                CultureInfo.InvariantCulture, out double mass) ? mass : 0;
        }

        private static BitmapSource TryLoadPreview(string filePath, string savedPreviewPath = null)
        {
            // Если исходный файл КОМПАС доступен — загружаем из него
            if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
            {
                try
                {
                    var preview = ThumbnailService.GetFileThumbnail(filePath);
                    if (preview != null)
                        return preview;
                }
                catch { }
            }

            // Fallback: загружаем из сохранённого PNG если исходный файл недоступен
            if (!string.IsNullOrEmpty(savedPreviewPath))
            {
                try
                {
                    return ThumbnailService.LoadPreviewFromFile(savedPreviewPath);
                }
                catch { }
            }

            return null;
        }

        private static string FormatMaterial(string material)
        {
            if (string.IsNullOrWhiteSpace(material))
                return material ?? string.Empty;

            string result = material;

            // 1. Заменяем $d на пробел
            result = result.Replace("$d", " ");

            // 2. Убираем оставшиеся $
            result = result.Replace("$", "");

            return result;
        }

        //private static string FormatMaterial(string material)
        //{
        //    if (string.IsNullOrWhiteSpace(material))
        //        return material ?? string.Empty;

        //    // Удаляем состояние поверхности
        //    string result = Regex.Replace(material,
        //        @"\s+х/к\s*\([^)]+\)", "", RegexOptions.IgnoreCase);

        //    // Извлекаем толщину
        //    var thicknessMatch = Regex.Match(result, @"\$d(\d+\.?\d*)");
        //    string thickness = thicknessMatch.Success ? thicknessMatch.Groups[1].Value : null;

        //    // Извлекаем марку стали
        //    var steelGradeMatch = Regex.Match(result, @";([A-Z]+\s*\d*)");
        //    string steelGrade = steelGradeMatch.Success
        //        ? steelGradeMatch.Groups[1].Value.Trim()
        //        : null;

        //    if (thicknessMatch.Success || steelGradeMatch.Success)
        //    {
        //        var baseMatch = Regex.Match(result, @"^([А-Яа-яA-Za-z]+)");
        //        string basePart = baseMatch.Success ? baseMatch.Groups[1].Value : "Лист";

        //        var parts = new[]
        //        {
        //            basePart,
        //            thickness != null ? $"{thickness} мм" : null,
        //            steelGrade ?? DefaultSteelGrade
        //        };

        //        return string.Join(" ", Array.FindAll(parts, p => p != null));
        //    }

        //    return result;
        //}

        public event PropertyChangedEventHandler PropertyChanged;

        public virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                // Очищаем все превью
                _filePreview = null;
                _drawingPreview = null;
                _previewLoaded = false;
                _drawingPreviewLoaded = false;
            }

            _disposed = true;
        }
    }
}