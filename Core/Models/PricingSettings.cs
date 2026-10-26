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
    /// Цена лазерной резки для листа определённой толщины
    /// </summary>
    [DataContract]
    public class LaserCuttingPricingEntry : INotifyPropertyChanged
    {
        private double _thickness;
        private double _pricePerMeter;

        /// <summary>
        /// Толщина листа, мм
        /// </summary>
        [DataMember]
        public double Thickness
        {
            get => _thickness;
            set
            {
                if (Math.Abs(_thickness - value) > 0.0001)
                {
                    _thickness = value;
                    OnPropertyChanged(nameof(Thickness));
                }
            }
        }

        /// <summary>
        /// Цена резки, руб/м
        /// </summary>
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
        private double _laserCuttingPricePerMeter;
        private double _engravingPricePerMm;

        // Цена резки из файлов старых версий (руб/мм); переводится в LaserCuttingPricePerMeter при чтении
        [DataMember(Name = "LaserCuttingPricePerMm", EmitDefaultValue = false)]
        private double _legacyLaserCuttingPricePerMm;
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
        /// Цена лазерной резки по умолчанию (для толщин, которых нет в LaserCuttingPricing), руб/м
        /// </summary>
        [DataMember]
        public double LaserCuttingPricePerMeter
        {
            get => _laserCuttingPricePerMeter;
            set
            {
                if (Math.Abs(_laserCuttingPricePerMeter - value) > 0.0001)
                {
                    _laserCuttingPricePerMeter = value;
                    OnPropertyChanged(nameof(LaserCuttingPricePerMeter));
                }
            }
        }

        /// <summary>
        /// Цены лазерной резки по толщине листа
        /// </summary>
        [DataMember]
        public ObservableCollection<LaserCuttingPricingEntry> LaserCuttingPricing { get; set; }
            = new ObservableCollection<LaserCuttingPricingEntry>();

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
        /// Кто и когда последним изменил расценки (в старых файлах отсутствует)
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string ModifiedBy { get; set; }

        [DataMember(EmitDefaultValue = false)]
        public string ModifiedByName { get; set; }

        [DataMember(EmitDefaultValue = false)]
        public long ModifiedUtcTicks { get; set; }

        /// <summary>
        /// ModifiedUtcTicks версии, от которой сделана эта правка. Если на сервере к моменту выкладки
        /// лежит более поздняя версия, её изменил кто-то другой — перезаписывать её нельзя
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public long BaseModifiedUtcTicks { get; set; }

        /// <summary>
        /// «Изменил: ФИО, дата» для окна расценок (null — не известно)
        /// </summary>
        public string ModifiedInfo
        {
            get
            {
                string by = string.IsNullOrWhiteSpace(ModifiedByName) ? ModifiedBy : ModifiedByName;
                if (string.IsNullOrEmpty(by))
                    return null;

                return ModifiedUtcTicks > 0
                    ? $"Последнее изменение: {by}, {new DateTime(ModifiedUtcTicks, DateTimeKind.Utc).ToLocalTime():dd.MM.yyyy HH:mm}"
                    : $"Последнее изменение: {by}";
            }
        }

        /// <summary>
        /// Найти цену за метр трубы по подстроке сортамента в материале
        /// </summary>
        public double GetTubularPricePerMeter(string material)
        {
            if (string.IsNullOrEmpty(material) || TubularPricing == null || TubularPricing.Count == 0)
                return 0;

            string normalizedMaterial = NormalizeSize(material);

            // Самая длинная подходящая запись = самая точная («40х40х3» приоритетнее «40х40»)
            TubularPricingEntry best = null;
            int bestLength = 0;

            foreach (var entry in TubularPricing)
            {
                string size = NormalizeSize(entry.Size);
                if (size.Length == 0 || !ContainsSize(normalizedMaterial, size))
                    continue;

                if (best == null || size.Length > bestLength)
                {
                    best = entry;
                    bestLength = size.Length;
                }
            }

            return best != null ? best.PricePerMeter : 0;
        }

        /// <summary>
        /// Сортамент для сравнения: «x», «X», «×» и «Х» → кириллическая «х», «,» → «.», без пробелов вокруг «х»
        /// («40 x 40» и «40х40» — один размер), остальные пробелы — по одному
        /// </summary>
        private static string NormalizeSize(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            var chars = new StringBuilder(text.Length);
            foreach (char c in text.Trim())
            {
                switch (c)
                {
                    case 'x':
                    case 'X':
                    case '×':
                    case 'Х':
                        chars.Append('х');
                        break;
                    case ',':
                        chars.Append('.');
                        break;
                    default:
                        chars.Append(char.IsWhiteSpace(c) ? ' ' : char.ToLowerInvariant(c));
                        break;
                }
            }

            var sb = new StringBuilder(chars.Length);
            for (int i = 0; i < chars.Length; i++)
            {
                char c = chars[i];
                if (c != ' ')
                {
                    sb.Append(c);
                    continue;
                }

                int next = i;
                while (next < chars.Length && chars[next] == ' ')
                    next++;

                bool nearX = (sb.Length > 0 && sb[sb.Length - 1] == 'х') || (next < chars.Length && chars[next] == 'х');
                if (!nearX)
                    sb.Append(' ');
                i = next - 1;
            }

            return sb.ToString();
        }

        /// <summary>
        /// Размер входит в материал отдельным числом: «40х40» не совпадает с «140х40х3» и «40х400»
        /// </summary>
        private static bool ContainsSize(string material, string size)
        {
            for (int index = material.IndexOf(size, StringComparison.Ordinal);
                 index >= 0;
                 index = material.IndexOf(size, index + 1, StringComparison.Ordinal))
            {
                int end = index + size.Length;
                bool startOk = !char.IsDigit(size[0]) || !ContinuesNumberBefore(material, index);
                bool endOk = !char.IsDigit(size[size.Length - 1]) || !ContinuesNumberAfter(material, end);
                if (startOk && endOk)
                    return true;
            }

            return false;
        }

        // «1» перед «40х40» или «2.» перед «5х…» — число продолжается
        private static bool ContinuesNumberBefore(string text, int index)
        {
            if (index == 0)
                return false;

            char prev = text[index - 1];
            return char.IsDigit(prev) || (prev == '.' && index >= 2 && char.IsDigit(text[index - 2]));
        }

        // «0» после «40х40» или «.5» — число продолжается; точка в конце предложения — нет
        private static bool ContinuesNumberAfter(string text, int end)
        {
            if (end >= text.Length)
                return false;

            char next = text[end];
            return char.IsDigit(next) || (next == '.' && end + 1 < text.Length && char.IsDigit(text[end + 1]));
        }

        /// <summary>
        /// Цена резки за метр для листа заданной толщины: из таблицы по точной толщине,
        /// иначе цена по умолчанию
        /// </summary>
        public double GetLaserCuttingPricePerMeter(double thickness)
        {
            var entry = FindLaserCuttingEntry(thickness);
            return entry != null ? entry.PricePerMeter : LaserCuttingPricePerMeter;
        }

        /// <summary>
        /// Цены по толщинам заданы, но для этой толщины (или нераспознанной) цены нет —
        /// резка посчитается по цене по умолчанию
        /// </summary>
        public bool UsesDefaultLaserCuttingPrice(double thickness)
        {
            return LaserCuttingPricing != null && LaserCuttingPricing.Count > 0 && FindLaserCuttingEntry(thickness) == null;
        }

        private LaserCuttingPricingEntry FindLaserCuttingEntry(double thickness)
        {
            if (thickness <= 0 || LaserCuttingPricing == null)
                return null;

            foreach (var entry in LaserCuttingPricing)
            {
                if (Math.Abs(entry.Thickness - thickness) < 0.001)
                    return entry;
            }

            return null;
        }

        /// <summary>
        /// Приведение файлов старых версий: DataContractJsonSerializer не выполняет инициализаторы
        /// полей (коллекции остались бы null), цена резки раньше хранилась в руб/мм
        /// </summary>
        [OnDeserialized]
        private void OnDeserialized(StreamingContext context)
        {
            if (TubularPricing == null)
                TubularPricing = new ObservableCollection<TubularPricingEntry>();
            if (LaserCuttingPricing == null)
                LaserCuttingPricing = new ObservableCollection<LaserCuttingPricingEntry>();

            if (_legacyLaserCuttingPricePerMm > 0 && _laserCuttingPricePerMeter <= 0)
                _laserCuttingPricePerMeter = _legacyLaserCuttingPricePerMm * 1000;
            _legacyLaserCuttingPricePerMm = 0;
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
        /// <param name="conflict">Не null — локальные правки не выложены: расценки на сервере за это время
        /// изменил другой сотрудник. Локальная версия сохранена в отдельный файл, загружены серверные</param>
        public static PricingSettings SyncWithServer(string serverFilePath, bool downloadOnly, out string conflict)
        {
            conflict = null;

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
                            string changedBy = DescribeServerChange(serverFilePath, local.BaseModifiedUtcTicks);
                            if (changedBy == null)
                            {
                                // Не удалось выложить — локальная копия остаётся новее и уйдёт при следующей синхронизации
                                local.Save(serverFilePath);
                                return local;
                            }
                            else
                            {
                                string conflictPath = Path.Combine(Path.GetDirectoryName(SettingsPath),
                                    $"pricing_settings.conflict_{DateTime.Now:yyyyMMdd_HHmmss}.json");
                                File.Copy(SettingsPath, conflictPath, true);
                                conflict = $"расценки на сервере изменил {changedBy}, ваши правки без сети не выложены " +
                                           $"(сохранены в {Path.GetFileName(conflictPath)})";
                                Logger.LogWarning($"Конфликт расценок: {conflict}");
                            }
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

        /// <summary>
        /// Если расценки на сервере изменены позже версии baseModifiedUtcTicks — «ФИО, дата» этого изменения, иначе null
        /// </summary>
        public static string DescribeServerChange(string serverFilePath, long baseModifiedUtcTicks)
        {
            if (string.IsNullOrEmpty(serverFilePath))
                return null;

            PricingSettings server;
            try
            {
                if (!File.Exists(serverFilePath))
                    return null;
                server = Deserialize(File.ReadAllBytes(serverFilePath));
            }
            catch (Exception ex)
            {
                // Нечитаемый файл на сервере сохранением как раз чинится
                Logger.LogWarning($"Не удалось прочитать общие расценки {serverFilePath}: {ex.Message}");
                return null;
            }

            if (server == null || server.ModifiedUtcTicks <= baseModifiedUtcTicks)
                return null;

            string by = string.IsNullOrWhiteSpace(server.ModifiedByName) ? server.ModifiedBy : server.ModifiedByName;
            string when = new DateTime(server.ModifiedUtcTicks, DateTimeKind.Utc).ToLocalTime().ToString("dd.MM.yyyy HH:mm");
            return string.IsNullOrEmpty(by) ? when : $"{by}, {when}";
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
                    using (SharedFileLock.Acquire(serverFilePath))
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