using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using TankManager.Core.Services;

namespace TankManager.Core.Models
{
    /// <summary>
    /// Единица цены материала
    /// </summary>
    public enum MaterialPriceUnit
    {
        /// <summary>
        /// Руб/кг (по массе детали)
        /// </summary>
        Kilogram,

        /// <summary>
        /// Руб/м (по длине детали — есть только у трубного проката)
        /// </summary>
        Meter
    }

    /// <summary>
    /// Правило цены материала: вид проката, марка и размер (пустое поле — любое значение).
    /// Детали подходит самое точное правило (см. PricingSettings.FindMaterialPrice)
    /// </summary>
    [DataContract]
    public class MaterialPriceEntry : INotifyPropertyChanged
    {
        private string _kind;
        private string _grade;
        private string _size;
        private MaterialPriceUnit _unit;
        private double? _price;

        public static readonly KeyValuePair<MaterialPriceUnit, string>[] UnitOptions =
        {
            new KeyValuePair<MaterialPriceUnit, string>(MaterialPriceUnit.Kilogram, "руб/кг"),
            new KeyValuePair<MaterialPriceUnit, string>(MaterialPriceUnit.Meter, "руб/м")
        };

        /// <summary>
        /// Вид проката («Лист», «Труба»; пусто — любой)
        /// </summary>
        [DataMember]
        public string Kind
        {
            get => _kind;
            set
            {
                if (_kind != value)
                {
                    _kind = value;
                    OnPropertyChanged(nameof(Kind));
                }
            }
        }

        /// <summary>
        /// Марка («AISI 304»; пусто — любая)
        /// </summary>
        [DataMember]
        public string Grade
        {
            get => _grade;
            set
            {
                if (_grade != value)
                {
                    _grade = value;
                    OnPropertyChanged(nameof(Grade));
                }
            }
        }

        /// <summary>
        /// Размер: толщина листа, сортамент трубы, диаметр круга (пусто — любой)
        /// </summary>
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
        public MaterialPriceUnit Unit
        {
            get => _unit;
            set
            {
                if (_unit != value)
                {
                    _unit = value;
                    OnPropertyChanged(nameof(Unit));
                    OnPropertyChanged(nameof(UnitTitle));
                }
            }
        }

        /// <summary>
        /// Цена, руб за единицу Unit (null — материал добавлен из изделия, цену ещё не ввели)
        /// </summary>
        [DataMember]
        public double? Price
        {
            get => _price;
            set
            {
                if (_price.HasValue != value.HasValue || (value.HasValue && Math.Abs(_price.Value - value.Value) > 0.0001))
                {
                    _price = value;
                    OnPropertyChanged(nameof(Price));
                    OnPropertyChanged(nameof(HasPrice));
                }
            }
        }

        public bool HasPrice => Price.HasValue;

        /// <summary>
        /// Материал правила (для правила без цены — материала, добавленного из изделия)
        /// </summary>
        public MaterialKey Key => new MaterialKey(Kind?.Trim(), MaterialKey.FormatSize(Size), Grade?.Trim());

        public string UnitTitle => Unit == MaterialPriceUnit.Meter ? "руб/м" : "руб/кг";

        /// <summary>
        /// «Лист · AISI 304» (пустые поля не показываются)
        /// </summary>
        public string Title => MaterialKey.FormatTitle(Kind?.Trim(), MaterialKey.FormatSize(Size), Grade?.Trim(), "Любой материал");

        /// <summary>
        /// Все заполненные поля правила совпадают с материалом
        /// </summary>
        public bool Matches(MaterialKey key)
        {
            if (key == null || key.IsUnknown)
                return false;

            return (string.IsNullOrWhiteSpace(Kind) || MaterialKey.TextKey(Kind) == MaterialKey.TextKey(key.Kind))
                && (string.IsNullOrWhiteSpace(Grade) || MaterialKey.TextKey(Grade) == MaterialKey.TextKey(key.Grade))
                && (string.IsNullOrWhiteSpace(Size) || MaterialKey.SizeKey(Size) == MaterialKey.SizeKey(key.Size));
        }

        /// <summary>
        /// Точность правила: марка важнее размера, размер важнее вида проката
        /// </summary>
        public int Specificity =>
            (string.IsNullOrWhiteSpace(Grade) ? 0 : 4) +
            (string.IsNullOrWhiteSpace(Size) ? 0 : 2) +
            (string.IsNullOrWhiteSpace(Kind) ? 0 : 1);

        /// <summary>
        /// Правило задано ровно для этого материала (все три поля совпадают)
        /// </summary>
        public bool IsExactFor(MaterialKey key) =>
            key != null &&
            MaterialKey.TextKey(Kind) == MaterialKey.TextKey(key.Kind) &&
            MaterialKey.TextKey(Grade) == MaterialKey.TextKey(key.Grade) &&
            MaterialKey.SizeKey(Size) == MaterialKey.SizeKey(key.Size);

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

        private double _laserCuttingPricePerMeter;
        private double _engravingPricePerMm;

        // Цена резки из файлов старых версий (руб/мм); переводится в LaserCuttingPricePerMeter при чтении
        [DataMember(Name = "LaserCuttingPricePerMm", EmitDefaultValue = false)]
        private double _legacyLaserCuttingPricePerMm;
        private double _bendingPricePerOperation;
        private double _rollingPricePerKg;
        private double _flangingPricePerOperation;

        /// <summary>
        /// Цены материалов (правила по виду проката, марке и размеру)
        /// </summary>
        [DataMember]
        public ObservableCollection<MaterialPriceEntry> MaterialPricing { get; set; }
            = new ObservableCollection<MaterialPriceEntry>();

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
        /// Самое точное правило цены для материала (null — цены нет). Правила без цены не учитываются.
        /// При равной точности выигрывает правило выше в таблице
        /// </summary>
        public MaterialPriceEntry FindMaterialPrice(MaterialKey key)
        {
            if (key == null || key.IsUnknown || MaterialPricing == null)
                return null;

            MaterialPriceEntry best = null;
            foreach (var entry in MaterialPricing)
            {
                if (entry != null && entry.HasPrice && entry.Matches(key) && (best == null || entry.Specificity > best.Specificity))
                    best = entry;
            }
            return best;
        }

        /// <summary>
        /// Задать цену ровно для этого материала (null — стереть цену: будет действовать более общее правило,
        /// а если его нет — материал останется в списке без цены)
        /// </summary>
        public void SetMaterialPrice(MaterialKey key, double? price, MaterialPriceUnit unit)
        {
            if (key == null || key.IsUnknown)
                return;

            var exact = MaterialPricing.FirstOrDefault(e => e != null && e.IsExactFor(key));
            if (exact == null)
            {
                if (price == null)
                    return;
                exact = new MaterialPriceEntry { Kind = key.Kind, Grade = key.Grade, Size = key.Size };
                MaterialPricing.Add(exact);
            }
            exact.Unit = unit;
            exact.Price = price;
        }

        /// <summary>
        /// Добавить в список материалы изделия, для которых нет ни одного правила с ценой
        /// (без цены — чтобы её ввели). Возвращает число добавленных
        /// </summary>
        public int AddMissingMaterials(IEnumerable<KeyValuePair<MaterialKey, MaterialPriceUnit>> materials)
        {
            int added = 0;
            foreach (var material in materials)
            {
                var key = material.Key;
                if (key == null || key.IsUnknown || FindMaterialPrice(key) != null ||
                    MaterialPricing.Any(e => e != null && e.IsExactFor(key)))
                    continue;

                MaterialPricing.Add(new MaterialPriceEntry { Kind = key.Kind, Grade = key.Grade, Size = key.Size, Unit = material.Value });
                added++;
            }
            return added;
        }

        /// <summary>
        /// Убрать материалы без цены, которым цену уже даёт общее правило (например, задали цену на марку),
        /// и пустые строки таблицы
        /// </summary>
        public void RemoveCoveredMaterials()
        {
            foreach (var entry in MaterialPricing.Where(e => e == null ||
                (!e.HasPrice && (e.Key.IsUnknown || FindMaterialPrice(e.Key) != null))).ToList())
                MaterialPricing.Remove(entry);
        }

        /// <summary>
        /// Список материалов вместе с правилами без цены (их нет в журнале, но сохранять их нужно)
        /// </summary>
        public string MaterialListSignature => string.Join("\n", MaterialPricing.Where(e => e != null)
            .Select(e => $"{e.Kind}|{e.Size}|{e.Grade}|{e.Unit}|{e.Price}"));

        /// <summary>
        /// Независимая копия (для редактирования)
        /// </summary>
        public PricingSettings Clone()
        {
            var serializer = new DataContractJsonSerializer(typeof(PricingSettings));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, this);
                return Deserialize(stream.ToArray());
            }
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
            if (MaterialPricing == null)
                MaterialPricing = new ObservableCollection<MaterialPriceEntry>();
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