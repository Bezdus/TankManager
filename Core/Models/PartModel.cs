using Kompas6API5;
using KompasAPI7;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;
using TankManager.Core.Constants;
using TankManager.Core.Services;
using static Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System;

namespace TankManager.Core.Models
{
    public class PartModel : INotifyPropertyChanged, IDisposable
    {
        private const string DefaultSteelGrade = "AISI 304";

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
            }
            finally
            {
                // Не освобождаем parentPart, т.к. это приведение типа
            }
        }

        private static double GetLength(object detail, KompasContext context)
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
                    materialLower.Contains("труб"))
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

        private static string FormatMaterial(string material)
        {
            if (string.IsNullOrWhiteSpace(material))
                return material ?? string.Empty;

            // Удаляем состояние поверхности
            string result = Regex.Replace(material,
                @"\s+х/к\s*\([^)]+\)", "", RegexOptions.IgnoreCase);

            // Извлекаем толщину
            var thicknessMatch = Regex.Match(result, @"\$d(\d+\.?\d*)");
            string thickness = thicknessMatch.Success ? thicknessMatch.Groups[1].Value : null;

            // Извлекаем марку стали
            var steelGradeMatch = Regex.Match(result, @";([A-Z]+\s*\d*)");
            string steelGrade = steelGradeMatch.Success
                ? steelGradeMatch.Groups[1].Value.Trim()
                : null;

            if (thicknessMatch.Success || steelGradeMatch.Success)
            {
                var baseMatch = Regex.Match(result, @"^([А-Яа-яA-Za-z]+)");
                string basePart = baseMatch.Success ? baseMatch.Groups[1].Value : "Лист";

                var parts = new[]
                {
                    basePart,
                    thickness != null ? $"{thickness} мм" : null,
                    steelGrade ?? DefaultSteelGrade
                };

                return string.Join(" ", Array.FindAll(parts, p => p != null));
            }

            return result;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
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
                _filePreview = null;
            }

            _disposed = true;
        }
    }
}