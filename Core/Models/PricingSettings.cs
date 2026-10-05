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
        private double _rollingPricePerKg;
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
        /// Цена вальцовки, руб/кг (по массе детали)
        /// </summary>
        [DataMember]
        public double RollingPricePerKg
        {
            get => _rollingPricePerKg;
            set
            {
                if (Math.Abs(_rollingPricePerKg - value) > 0.0001)
                {
                    _rollingPricePerKg = value;
                    OnPropertyChanged(nameof(RollingPricePerKg));
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
        /// Имя файла общих расценок в серверной папке изделий
        /// </summary>
        public const string ServerFileName = "_pricing_settings.json";

        /// <summary>
        /// Загрузить расценки: общие с сервера (с обновлением локальной копии),
        /// а если сервер недоступен — локальную копию
        /// </summary>
        /// <param name="serverFilePath">Путь к общим расценкам на сервере (может быть null)</param>
        public static PricingSettings Load(string serverFilePath = null)
        {
            if (!string.IsNullOrEmpty(serverFilePath))
            {
                try
                {
                    if (File.Exists(serverFilePath))
                    {
                        byte[] data = File.ReadAllBytes(serverFilePath);
                        var fromServer = Deserialize(data);
                        if (fromServer != null)
                        {
                            // Локальная копия нужна для работы без сети. Дата как у серверной:
                            // по ней SyncWithServer понимает, что локальных изменений нет
                            try
                            {
                                AtomicFile.WriteAllBytes(SettingsPath, data);
                                File.SetLastWriteTimeUtc(SettingsPath, File.GetLastWriteTimeUtc(serverFilePath));
                            }
                            catch (Exception ex) { Logger.LogWarning($"Не удалось обновить локальную копию расценок: {ex.Message}"); }

                            return fromServer;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"Не удалось прочитать общие расценки {serverFilePath}: {ex.Message}");
                }
            }

            return LoadLocal();
        }

        /// <summary>
        /// Согласовать расценки с сервером и загрузить актуальные.
        /// Конструктор выкладывает локальную копию, если она новее серверной (или на сервере файла нет);
        /// в режиме просмотра расценки только читаются с сервера.
        /// </summary>
        /// <param name="serverFilePath">Путь к общим расценкам на сервере (null — только локальные)</param>
        /// <param name="downloadOnly">Не записывать на сервер</param>
        public static PricingSettings SyncWithServer(string serverFilePath, bool downloadOnly)
        {
            if (string.IsNullOrEmpty(serverFilePath))
                return LoadLocal();

            if (!downloadOnly && File.Exists(SettingsPath))
            {
                try
                {
                    bool serverExists = File.Exists(serverFilePath);
                    if (!serverExists || File.GetLastWriteTimeUtc(SettingsPath) > File.GetLastWriteTimeUtc(serverFilePath))
                    {
                        // Повреждённую локальную копию на сервер не выкладываем
                        var local = Deserialize(File.ReadAllBytes(SettingsPath));
                        if (local != null)
                        {
                            local.Save(serverFilePath);
                            return local;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"Не удалось выложить расценки на сервер: {ex.Message}");
                }
            }

            return Load(serverFilePath);
        }

        private static PricingSettings LoadLocal()
        {
            if (!File.Exists(SettingsPath))
                return new PricingSettings();

            try
            {
                var loaded = Deserialize(File.ReadAllBytes(SettingsPath));
                if (loaded != null)
                    return loaded;
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

        private static PricingSettings Deserialize(byte[] data)
        {
            using (var stream = new MemoryStream(data))
            {
                var serializer = new DataContractJsonSerializer(typeof(PricingSettings));
                return (PricingSettings)serializer.ReadObject(stream);
            }
        }

        /// <summary>
        /// Сохранить расценки (атомарно) локально и, если указан путь, на сервер.
        /// Ошибка локальной записи бросает исключение; ошибка сервера возвращается строкой.
        /// </summary>
        /// <returns>Текст ошибки записи на сервер или null</returns>
        public string Save(string serverFilePath = null)
        {
            byte[] data;
            var serializer = new DataContractJsonSerializer(typeof(PricingSettings));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, this);
                data = stream.ToArray();
            }

            string serverError = null;
            if (!string.IsNullOrEmpty(serverFilePath))
            {
                try
                {
                    AtomicFile.WriteAllBytes(serverFilePath, data);
                }
                catch (Exception ex)
                {
                    serverError = ex.Message;
                    Logger.LogError("Не удалось сохранить общие расценки на сервер", ex);
                }
            }

            AtomicFile.WriteAllBytes(SettingsPath, data);

            // Одинаковая дата — признак того, что локальная копия совпадает с серверной (см. SyncWithServer)
            if (!string.IsNullOrEmpty(serverFilePath) && serverError == null)
            {
                try { File.SetLastWriteTimeUtc(SettingsPath, File.GetLastWriteTimeUtc(serverFilePath)); }
                catch (Exception ex) { Logger.LogWarning($"Не удалось выставить дату локальной копии расценок: {ex.Message}"); }
            }

            return serverError;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}