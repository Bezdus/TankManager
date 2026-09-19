using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using TankManager.Core.Services;

namespace TankManager.Core.Models
{
    /// <summary>
    /// Запись сортамента трубного проката
    /// </summary>
    [DataContract]
    public class TubularPricingEntry : INotifyPropertyChanged
    {
        private string _size;
        private double _pricePerMeter;

        [DataMember]
        public string Size
        {
            get => _size;
            set
            {
                if (_size != value)
                {
                    _size = value;
                    OnPropertyChanged(nameof(Size));
                }
            }
        }

        [DataMember]
        public double PricePerMeter
        {
            get => _pricePerMeter;
            set
            {
                if (Math.Abs(_pricePerMeter - value) > 0.0001)
                {
                    _pricePerMeter = value;
                    OnPropertyChanged(nameof(PricePerMeter));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Настройки расценок для расчёта стоимости деталей
    /// </summary>
    [DataContract]
    public class PricingSettings : INotifyPropertyChanged
    {
        private static readonly string SettingsPath =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pricing_settings.json");

        private static readonly ILogger Logger = new FileLogger();

        private double _sheetMetalPricePerKg;
        private double _otherMetalPricePerKg;
        private double _laserCuttingPricePerMm;
        private double _engravingPricePerMm;
        private double _bendingPricePerOperation;
        private double _rollingPricePerMm;
        private double _flangingPricePerOperation;

        /// <summary>
        /// Цена листового проката, руб/кг
        /// </summary>
        [DataMember]
        public double SheetMetalPricePerKg
        {
            get => _sheetMetalPricePerKg;
            set
            {
                if (Math.Abs(_sheetMetalPricePerKg - value) > 0.0001)
                {
                    _sheetMetalPricePerKg = value;
                    OnPropertyChanged(nameof(SheetMetalPricePerKg));
                }
            }
        }

        /// <summary>
        /// Цена прочего металла, руб/кг
        /// </summary>
        [DataMember]
        public double OtherMetalPricePerKg
        {
            get => _otherMetalPricePerKg;
            set
            {
                if (Math.Abs(_otherMetalPricePerKg - value) > 0.0001)
                {
                    _otherMetalPricePerKg = value;
                    OnPropertyChanged(nameof(OtherMetalPricePerKg));
                }
            }
        }

        /// <summary>
        /// Сортамент трубного проката с ценами за метр
        /// </summary>
        [DataMember]
        public ObservableCollection<TubularPricingEntry> TubularPricing { get; set; }
            = new ObservableCollection<TubularPricingEntry>();

        /// <summary>
        /// Цена лазерной резки, руб/мм
        /// </summary>
        [DataMember]
        public double LaserCuttingPricePerMm
        {
            get => _laserCuttingPricePerMm;
            set
            {
                if (Math.Abs(_laserCuttingPricePerMm - value) > 0.0001)
                {
                    _laserCuttingPricePerMm = value;
                    OnPropertyChanged(nameof(LaserCuttingPricePerMm));
                }
            }
        }

        /// <summary>
        /// Цена гравировки, руб/мм
        /// </summary>
        [DataMember]
        public double EngravingPricePerMm
        {
            get => _engravingPricePerMm;
            set
            {
                if (Math.Abs(_engravingPricePerMm - value) > 0.0001)
                {
                    _engravingPricePerMm = value;
                    OnPropertyChanged(nameof(EngravingPricePerMm));
                }
            }
        }

        /// <summary>
        /// Цена гибки, руб/операция
        /// </summary>
        [DataMember]
        public double BendingPricePerOperation
        {
            get => _bendingPricePerOperation;
            set
            {
                if (Math.Abs(_bendingPricePerOperation - value) > 0.0001)
                {
                    _bendingPricePerOperation = value;
                    OnPropertyChanged(nameof(BendingPricePerOperation));
                }
            }
        }

        /// <summary>
        /// Цена вальцовки, руб/мм
        /// </summary>
        [DataMember]
        public double RollingPricePerMm
        {
            get => _rollingPricePerMm;
            set
            {
                if (Math.Abs(_rollingPricePerMm - value) > 0.0001)
                {
                    _rollingPricePerMm = value;
                    OnPropertyChanged(nameof(RollingPricePerMm));
                }
            }
        }

        /// <summary>
        /// Цена отбортовки, руб/операция
        /// </summary>
        [DataMember]
        public double FlangingPricePerOperation
        {
            get => _flangingPricePerOperation;
            set
            {
                if (Math.Abs(_flangingPricePerOperation - value) > 0.0001)
                {
                    _flangingPricePerOperation = value;
                    OnPropertyChanged(nameof(FlangingPricePerOperation));
                }
            }
        }

        /// <summary>
        /// Найти цену за метр трубы по подстроке сортамента в материале
        /// </summary>
        public double GetTubularPricePerMeter(string material)
        {
            if (string.IsNullOrEmpty(material) || TubularPricing == null || TubularPricing.Count == 0)
                return 0;

            // Самая длинная подходящая запись = самая точная («40х40х3» приоритетнее «40х40»)
            TubularPricingEntry best = null;

            foreach (var entry in TubularPricing)
            {
                if (string.IsNullOrEmpty(entry.Size) ||
                    material.IndexOf(entry.Size, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                if (best == null || entry.Size.Length > best.Size.Length)
                    best = entry;
            }

            return best != null ? best.PricePerMeter : 0;
        }

        /// <summary>
        /// Загрузить настройки из файла
        /// </summary>
        public static PricingSettings Load()
        {
            if (!File.Exists(SettingsPath))
                return new PricingSettings();

            try
            {
                using (var stream = new FileStream(SettingsPath, FileMode.Open, FileAccess.Read))
                {
                    var serializer = new DataContractJsonSerializer(typeof(PricingSettings));
                    var loaded = (PricingSettings)serializer.ReadObject(stream);
                    if (loaded != null)
                        return loaded;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Не удалось прочитать {SettingsPath}, используются расценки по умолчанию", ex);

                // Сохраняем повреждённый файл, чтобы следующее сохранение не стёрло расценки безвозвратно
                try { File.Copy(SettingsPath, SettingsPath + ".bad", true); }
                catch (Exception copyEx) { Logger.LogWarning($"Не удалось сохранить копию повреждённого файла: {copyEx.Message}"); }
            }

            return new PricingSettings();
        }

        /// <summary>
        /// Сохранить настройки в файл (атомарно). При ошибке бросает исключение.
        /// </summary>
        public void Save()
        {
            var serializer = new DataContractJsonSerializer(typeof(PricingSettings));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, this);
                AtomicFile.WriteAllBytes(SettingsPath, stream.ToArray());
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}