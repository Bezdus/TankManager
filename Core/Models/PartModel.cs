using KompasAPI7;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;
using TankManager.Core.Constants;
using TankManager.Core.Services;

namespace TankManager.Core.Models
{
    public class PartModel : INotifyPropertyChanged
    {
        private const string DefaultSteelGrade = "AISI 304";

        private string _name;
        private string _marking;
        private string _detailType;
        private string _material;
        private double _mass;
        private string _filePath;
        private BitmapSource _filePreview;

        public IBody7 Body { get; private set; }
        public IPart7 Part { get; private set; }

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

        public BitmapSource FilePreview
        {
            get { return _filePreview; }
            set
            {
                if (_filePreview != value)
                {
                    _filePreview = value;
                    OnPropertyChanged(nameof(FilePreview));
                }
            }
        }

        public PartModel(IPart7 part, KompasContext context)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));
            if (context == null) throw new ArgumentNullException(nameof(context));

            Part = part;
            Name = part.Name ?? string.Empty;
            Marking = part.Marking ?? string.Empty;
            DetailType = DetermineDetailType(context.GetDetailType(part));
            Material = FormatMaterial(part.Material);
            Mass = part.Mass / KompasConstants.MassConversionFactor;
            FilePath = part.FileName ?? string.Empty;
            FilePreview = TryLoadPreview(FilePath);
        }

        public PartModel(IBody7 body, KompasContext context)
        {
            if (body == null) throw new ArgumentNullException(nameof(body));
            if (context == null) throw new ArgumentNullException(nameof(context));

            Body = body;
            Name = body.Name ?? string.Empty;
            Marking = body.Marking ?? string.Empty;
            DetailType = KompasConstants.PartType;
            Material = FormatMaterial(
                context.GetBodyPropertyValue(body, KompasConstants.MaterialPropertyName));
            Mass = ParseMass(
                context.GetBodyPropertyValue(body, KompasConstants.MassPropertyName));
            FilePath = ((IPart7)body.Parent)?.FileName ?? string.Empty;
            FilePreview = TryLoadPreview(FilePath);
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
    }
}