using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using TankManager.Core.Models;


namespace TankManager.Core.Services
{
    /// <summary>
    /// Результат синхронизации
    /// </summary>
    public class SyncResult
    {
        public int NewProducts { get; set; }
        public int UpdatedProducts { get; set; }
        public int FailedProducts { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public bool Success => FailedProducts == 0 && Errors.Count == 0;
    }

    /// <summary>
    /// Сервис для сохранения и загрузки Product в локальную базу с синхронизацией с сервером
    /// </summary>
    public class ProductStorageService
    {
        private static readonly string ProductsDirectory =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "products");

        private static readonly string SettingsFilePath =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "storage_settings.json");

        private const string TombstonesFileName = "_deleted.json";
        private const string PendingTombstonesFileName = "_pending_deleted.json";
        private const int TombstoneRetentionDays = 90;
        private static readonly TimeSpan ServerCheckTtl = TimeSpan.FromSeconds(10);
        private const string ProductJsonFileName = "product.json";
        private const string ImagesSubfolder = "images";
        private const string FileExtension = ".json";

        private readonly ImageSyncService _imageSyncService = new ImageSyncService();
        private readonly ILogger _logger = new FileLogger();
        private readonly object _ioLock = new object();
        private readonly Dictionary<string, ProductMeta> _metaCache =
            new Dictionary<string, ProductMeta>(StringComparer.OrdinalIgnoreCase);
        private string _serverStorageFolder;
        private DateTime _serverCheckedAtUtc = DateTime.MinValue;
        private bool _serverAvailableCached;

        /// <summary>
        /// Ошибка последней операции с серверной папкой (null, если всё прошло успешно)
        /// </summary>
        public string LastServerError { get; private set; }

        /// <summary>
        /// Серверная (сетевая) папка для хранения изделий
        /// </summary>
        public string ServerStorageFolder
        {
            get => _serverStorageFolder;
            set
            {
                _serverStorageFolder = value;
                _serverCheckedAtUtc = DateTime.MinValue;
                SaveSettings();
            }
        }

        /// <summary>
        /// Проверяет, установлена ли серверная папка
        /// </summary>
        public bool HasServerFolder => !string.IsNullOrEmpty(_serverStorageFolder);

        /// <summary>
        /// Проверяет, доступна ли серверная папка
        /// </summary>
        public bool IsServerAvailable
        {
            get
            {
                if (!HasServerFolder)
                    return false;

                // Проверка сетевой папки дорогая, а свойство вызывается из CanExecute — кэшируем результат
                var now = DateTime.UtcNow;
                if (now - _serverCheckedAtUtc > ServerCheckTtl)
                {
                    _serverAvailableCached = Directory.Exists(_serverStorageFolder);
                    _serverCheckedAtUtc = now;
                }

                return _serverAvailableCached;
            }
        }

        public ProductStorageService()
        {
            Directory.CreateDirectory(ProductsDirectory);
            LoadSettings();
        }

        #region Settings

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(StorageSettings));
                    using (var fileStream = new FileStream(SettingsFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    using (var memoryStream = new MemoryStream())
                    {
                        fileStream.CopyTo(memoryStream);
                        memoryStream.Position = 0;
                        var settings = (StorageSettings)serializer.ReadObject(memoryStream);
                        _serverStorageFolder = settings?.ServerStorageFolder;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка загрузки настроек хранения: {ex.Message}");
            }
        }

        private void SaveSettings()
        {
            try
            {
                var settings = new StorageSettings { ServerStorageFolder = _serverStorageFolder };
                var serializer = new DataContractJsonSerializer(typeof(StorageSettings));

                using (var memoryStream = new MemoryStream())
                {
                    serializer.WriteObject(memoryStream, settings);
                    AtomicFile.WriteAllBytes(SettingsFilePath, memoryStream.ToArray());
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка сохранения настроек хранения: {ex.Message}");
            }
        }

        #endregion

        #region Synchronization

        /// <summary>
        /// Синхронизирует данные между локальной папкой и сервером (двусторонняя синхронизация).
        /// Копирует новые и обновлённые изделия в обе стороны.
        /// </summary>
        /// <param name="skipImages">Если true, изображения не копируются при синхронизации</param>
        public SyncResult SyncFromServer(bool skipImages = false)
        {
            lock (_ioLock)
            {
                return SyncFromServerCore(skipImages);
            }
        }

        private SyncResult SyncFromServerCore(bool skipImages)
        {
            var result = new SyncResult();

            // Явная проверка без кэша: синхронизация выполняется вне UI-потока
            _serverCheckedAtUtc = DateTime.MinValue;

            if (!IsServerAvailable)
            {
                if (HasServerFolder)
                    result.Errors.Add("Серверная папка недоступна");
                return result;
            }

            try
            {
                // Фаза 0: применение удалений (tombstone), чтобы удалённые изделия не воскресали
                var deleted = ProcessTombstones(result);

                // Фаза 1: Синхронизация С СЕРВЕРА В ЛОКАЛЬНУЮ ПАПКУ
                var serverFolders = Directory.GetDirectories(_serverStorageFolder)
                    .Where(f => !Path.GetFileName(f).StartsWith("_") && !deleted.Contains(Path.GetFileName(f)))
                    .ToList();

                foreach (var serverFolder in serverFolders)
                {
                    try
                    {
                        string folderName = Path.GetFileName(serverFolder);
                        string localFolder = Path.Combine(ProductsDirectory, folderName);
                        string serverJsonPath = Path.Combine(serverFolder, ProductJsonFileName);

                        if (!File.Exists(serverJsonPath))
                            continue;

                        var serverFileInfo = new FileInfo(serverJsonPath);
                        string localJsonPath = Path.Combine(localFolder, ProductJsonFileName);

                        bool needsCopy = false;
                        bool isNew = false;

                        if (!Directory.Exists(localFolder) || !File.Exists(localJsonPath))
                        {
                            // Новое изделие - нужно скопировать
                            needsCopy = true;
                            isNew = true;
                        }
                        else
                        {
                            // Проверяем дату модификации
                            var localFileInfo = new FileInfo(localJsonPath);
                            if (serverFileInfo.LastWriteTimeUtc > localFileInfo.LastWriteTimeUtc)
                            {
                                // Серверная версия новее - нужно обновить
                                needsCopy = true;
                                isNew = false;
                            }
                        }

                        if (needsCopy)
                        {
                            CopyProductFolder(serverFolder, localFolder, skipImages);

                            if (isNew)
                                result.NewProducts++;
                            else
                                result.UpdatedProducts++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedProducts++;
                        result.Errors.Add($"Ошибка синхронизации с сервера {Path.GetFileName(serverFolder)}: {ex.Message}");
                    }
                }

                // Фаза 2: Синхронизация ИЗ ЛОКАЛЬНОИ ПАПКИ НА СЕРВЕР
                var localFolders = Directory.GetDirectories(ProductsDirectory)
                    .Where(f => !Path.GetFileName(f).StartsWith("_") && !deleted.Contains(Path.GetFileName(f)))
                    .ToList();

                foreach (var localFolder in localFolders)
                {
                    try
                    {
                        string folderName = Path.GetFileName(localFolder);
                        string serverFolder = Path.Combine(_serverStorageFolder, folderName);
                        string localJsonPath = Path.Combine(localFolder, ProductJsonFileName);

                        // Пропускаем папки без product.json (это папки только с превью)
                        if (!File.Exists(localJsonPath))
                            continue;

                        var localFileInfo = new FileInfo(localJsonPath);
                        string serverJsonPath = Path.Combine(serverFolder, ProductJsonFileName);

                        bool needsCopy = false;

                        if (!Directory.Exists(serverFolder) || !File.Exists(serverJsonPath))
                        {
                            // Новое локальное изделие - загружаем на сервер
                            needsCopy = true;
                        }
                        else
                        {
                            // Проверяем дату модификации
                            var serverFileInfo = new FileInfo(serverJsonPath);
                            if (localFileInfo.LastWriteTimeUtc > serverFileInfo.LastWriteTimeUtc)
                            {
                                // Локальная версия новее - загружаем на сервер
                                needsCopy = true;
                            }
                        }

                        if (needsCopy)
                        {
                            CopyProductFolder(localFolder, serverFolder, skipImages);
                            // Не увеличиваем счётчики, так как уже посчитали в первой фазе
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add($"Ошибка загрузки на сервер {Path.GetFileName(localFolder)}: {ex.Message}");
                    }
                }

                // Фаза 3: Синхронизация изображений на уровне отдельных файлов
                if (!skipImages)
                    SyncAllProductImages();
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Ошибка доступа к папкам: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Копирует папку продукта с сервера в локальную директорию
        /// </summary>
        /// <param name="skipImages">Если true, подпапка images не копируется</param>
        private void CopyProductFolder(string sourceFolder, string destFolder, bool skipImages = false)
        {
            // Создаём целевую папку
            Directory.CreateDirectory(destFolder);

            // Копируем все файлы
            foreach (var file in Directory.GetFiles(sourceFolder))
            {
                string fileName = Path.GetFileName(file);
                string extension = Path.GetExtension(fileName);
                if (string.Equals(extension, ".bak", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(extension, ".tmp", StringComparison.OrdinalIgnoreCase))
                    continue;

                string destFile = Path.Combine(destFolder, fileName);

                // Резервная копия перед перезаписью product.json
                if (string.Equals(fileName, ProductJsonFileName, StringComparison.OrdinalIgnoreCase) && File.Exists(destFile))
                    File.Copy(destFile, destFile + ".bak", true);

                File.Copy(file, destFile, true);
            }

            // Копируем подпапки (пропускаем images если skipImages)
            foreach (var dir in Directory.GetDirectories(sourceFolder))
            {
                if (skipImages && string.Equals(Path.GetFileName(dir), ImagesSubfolder, StringComparison.OrdinalIgnoreCase))
                    continue;

                string destSubDir = Path.Combine(destFolder, Path.GetFileName(dir));
                CopyProductFolder(dir, destSubDir, skipImages);
            }
        }

        #endregion

        /// <summary>
        /// Синхронизирует изображения для всех продуктов между локальной папкой и сервером
        /// </summary>
        private void SyncAllProductImages()
        {
            if (!IsServerAvailable)
                return;

            try
            {
                var localFolders = Directory.GetDirectories(ProductsDirectory)
                    .Where(f => !Path.GetFileName(f).StartsWith("_"))
                    .ToList();

                foreach (var localFolder in localFolders)
                {
                    try
                    {
                        string folderName = Path.GetFileName(localFolder);
                        string serverFolder = Path.Combine(_serverStorageFolder, folderName);

                        if (!Directory.Exists(serverFolder))
                            continue;

                        string localImages = Path.Combine(localFolder, ImagesSubfolder);
                        string serverImages = Path.Combine(serverFolder, ImagesSubfolder);

                        _imageSyncService.SyncImageDirectories(localImages, serverImages);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Ошибка синхронизации изображений {Path.GetFileName(localFolder)}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка синхронизации изображений: {ex.Message}");
            }
        }

        /// <summary>
        /// Синхронизирует изображения конкретного продукта между локальной папкой и сервером
        /// </summary>
        public void SyncProductImagesWithServer(Product product)
        {
            lock (_ioLock)
            {
                SyncProductImagesCore(product);
            }
        }

        private void SyncProductImagesCore(Product product)
        {
            if (!IsServerAvailable || product == null || string.IsNullOrEmpty(product.Name))
                return;

            try
            {
                string localFolder = FindExistingProductFolder(product, ProductsDirectory);
                string serverFolder = FindExistingProductFolder(product, _serverStorageFolder);

                if (localFolder == null || serverFolder == null)
                    return;

                string localImages = Path.Combine(localFolder, ImagesSubfolder);
                string serverImages = Path.Combine(serverFolder, ImagesSubfolder);

                _imageSyncService.SyncImageDirectories(localImages, serverImages);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка синхронизации изображений продукта: {ex.Message}");
            }
        }

        /// <summary>
        /// Возвращает путь к папке с изображениями для продукта
        /// </summary>
        public string GetProductImagesFolder(Product product)
        {
            if (product == null || string.IsNullOrEmpty(product.Name))
                return null;

            string productFolderName = GetProductFolderName(product);
            string productFolder = Path.Combine(ProductsDirectory, productFolderName);
            string imagesFolder = Path.Combine(productFolder, ImagesSubfolder);

            Directory.CreateDirectory(imagesFolder);
            return imagesFolder;
        }

        /// <summary>
        /// Сохраняет изделие в локальную папку и на сервер (если доступен).
        /// Ошибка локального сохранения пробрасывается, ошибка сервера попадает в <see cref="LastServerError"/>.
        /// </summary>
        public string Save(Product product, string customName = null)
        {
            if (product == null) return null;

            lock (_ioLock)
            {
                LastServerError = null;

                string localFilePath = SaveToDirectory(product, ProductsDirectory, customName);

                if (IsServerAvailable)
                {
                    try
                    {
                        SaveToDirectory(product, _serverStorageFolder, customName);
                    }
                    catch (Exception ex)
                    {
                        LastServerError = ex.Message;
                        _logger.LogError("Ошибка сохранения на сервер", ex);
                    }
                }
                else if (HasServerFolder)
                {
                    LastServerError = "Серверная папка недоступна";
                }

                return localFilePath;
            }
        }

        /// <summary>
        /// Сохраняет изделие в указанную директорию
        /// </summary>
        private string SaveToDirectory(Product product, string baseDirectory, string customName)
        {
            // Проверяем, есть ли уже сохранённое изделие
            string existingFolderPath = FindExistingProductFolder(product, baseDirectory);

            if (existingFolderPath != null)
            {
                // Перезаписываем существующее изделие
                string existingFilePath = Path.Combine(existingFolderPath, ProductJsonFileName);
                SaveToFile(product, existingFilePath);

                // Копируем изображения, если они в другой папке
                CopyImagesIfNeeded(product, existingFolderPath);

                return existingFilePath;
            }

            // Создаём новую папку для изделия
            string productFolderName = GenerateProductFolderName(product, baseDirectory, customName);
            string productFolderPath = Path.Combine(baseDirectory, productFolderName);
            Directory.CreateDirectory(productFolderPath);

            // Создаём папку для изображений
            string imagesFolder = Path.Combine(productFolderPath, ImagesSubfolder);
            Directory.CreateDirectory(imagesFolder);

            // Сохраняем JSON
            string filePath = Path.Combine(productFolderPath, ProductJsonFileName);
            SaveToFile(product, filePath);

            // Копируем изображения, если они в другой папке
            CopyImagesIfNeeded(product, productFolderPath);

            return filePath;
        }

        /// <summary>
        /// Копирует изображения из локальной папки продукта в целевую папку
        /// </summary>
        private void CopyImagesIfNeeded(Product product, string destProductFolder)
        {
            // Находим локальную папку продукта
            string localFolder = FindExistingProductFolder(product, ProductsDirectory);
            if (localFolder == null || localFolder == destProductFolder)
                return;

            string sourceImagesFolder = Path.Combine(localFolder, ImagesSubfolder);
            string destImagesFolder = Path.Combine(destProductFolder, ImagesSubfolder);

            if (!Directory.Exists(sourceImagesFolder))
                return;

            Directory.CreateDirectory(destImagesFolder);

            foreach (var file in Directory.GetFiles(sourceImagesFolder))
            {
                string destFile = Path.Combine(destImagesFolder, Path.GetFileName(file));
                try
                {
                    if (!File.Exists(destFile))
                    {
                        File.Copy(file, destFile, false);
                    }
                    else if (File.GetLastWriteTimeUtc(file) > File.GetLastWriteTimeUtc(destFile))
                    {
                        File.Copy(file, destFile, true);
                    }
                }
                catch (Exception ex) { _logger.LogWarning($"Ошибка копирования изображения {Path.GetFileName(file)}: {ex.Message}"); }
            }
        }

        /// <summary>
        /// Ищет существующую папку изделия по имени и обозначению (по кэшу метаданных)
        /// </summary>
        private string FindExistingProductFolder(Product product, string baseDirectory)
        {
            if (!Directory.Exists(baseDirectory))
                return null;

            var folders = Directory.GetDirectories(baseDirectory)
                .Where(f => !Path.GetFileName(f).StartsWith("_"));

            foreach (var folder in folders)
            {
                try
                {
                    var meta = GetMeta(folder);
                    if (meta != null && meta.Name == product.Name && meta.Marking == product.Marking)
                        return folder;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Не удалось прочитать папку {Path.GetFileName(folder)}: {ex.Message}");
                }
            }

            // Если не нашли по содержимому, ищем по имени папки (для старых версий)
            string expectedFolder = Path.Combine(baseDirectory, GetProductFolderName(product));
            return Directory.Exists(expectedFolder) ? expectedFolder : null;
        }

        /// <summary>
        /// Загрузить продукт по имени папки (только из локальной папки)
        /// </summary>
        public Product Load(string folderName)
        {
            string folderPath = Path.Combine(ProductsDirectory, folderName);
            string filePath = Path.Combine(folderPath, ProductJsonFileName);

            if (File.Exists(filePath))
            {
                return LoadFromFile(filePath, folderPath);
            }

            return null;
        }

        /// <summary>
        /// Получает список всех сохранённых изделий (только из локальной папки)
        /// </summary>
        public List<ProductFileInfo> GetSavedProducts()
        {
            var result = new List<ProductFileInfo>();

            if (!Directory.Exists(ProductsDirectory))
                return result;

            var folders = Directory.GetDirectories(ProductsDirectory)
                .Where(f => !Path.GetFileName(f).StartsWith("_"));

            foreach (var folder in folders)
            {
                try
                {
                    var meta = GetMeta(folder);
                    if (meta == null)
                        continue;

                    result.Add(new ProductFileInfo
                    {
                        FileName = Path.GetFileName(folder),
                        ProductName = meta.Name,
                        Marking = meta.Marking,
                        DetailsCount = meta.DetailsCount,
                        SavedDate = meta.SavedDate
                    });
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Не удалось прочитать изделие {Path.GetFileName(folder)}: {ex.Message}");
                }
            }

            return result.OrderByDescending(p => p.SavedDate).ToList();
        }

        /// <summary>
        /// Удаляет изделие только из локальной папки
        /// </summary>
        public bool DeleteLocal(string folderName)
        {
            lock (_ioLock)
            {
                return DeleteLocalCore(folderName);
            }
        }

        private bool DeleteLocalCore(string folderName)
        {
            if (!IsSafeFolderName(folderName))
                return false;

            try
            {
                string folderPath = Path.Combine(ProductsDirectory, folderName);
                if (Directory.Exists(folderPath))
                {
                    DeleteDirectoryRecursive(folderPath);
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Ошибка удаления из локальной папки", ex);
            }

            return false;
        }

        /// <summary>
        /// Удаляет изделие везде (из локальной папки и с сервера).
        /// Если сервер недоступен, удаление запоминается и применяется при следующей синхронизации.
        /// </summary>
        public bool Delete(string folderName)
        {
            if (!IsSafeFolderName(folderName))
                return false;

            lock (_ioLock)
            {
                LastServerError = null;
                bool deletedAny = DeleteLocalCore(folderName);

                if (!HasServerFolder)
                    return deletedAny;

                bool serverDone = false;

                if (IsServerAvailable)
                {
                    try
                    {
                        string folderPath = Path.Combine(_serverStorageFolder, folderName);
                        if (Directory.Exists(folderPath))
                        {
                            DeleteDirectoryRecursive(folderPath);
                            deletedAny = true;
                        }

                        RecordTombstone(Path.Combine(_serverStorageFolder, TombstonesFileName), folderName);
                        serverDone = true;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Ошибка удаления с сервера", ex);
                    }
                }

                if (!serverDone)
                {
                    LastServerError = "Сервер недоступен: удаление будет применено при следующей синхронизации";

                    try
                    {
                        RecordTombstone(Path.Combine(ProductsDirectory, PendingTombstonesFileName), folderName);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Не удалось запомнить удаление для сервера", ex);
                    }
                }

                return deletedAny;
            }
        }

        /// <summary>
        /// Удаляет директорию с повторными попытками; при неудаче бросает IOException
        /// </summary>
        private void DeleteDirectoryRecursive(string path)
        {
            if (!Directory.Exists(path))
                return;

            if (!FileLockDiagnostics.ForceDeleteDirectory(path, maxAttempts: 5, delayMs: 200))
                throw new IOException($"Не удалось удалить папку: {path}");
        }

        /// <summary>
        /// Проверить существование продукта
        /// </summary>
        public bool Exists(string folderName)
        {
            string folderPath = Path.Combine(ProductsDirectory, folderName);
            string jsonPath = Path.Combine(folderPath, ProductJsonFileName);
            return File.Exists(jsonPath);
        }

        /// <summary>
        /// Пытается загрузить ранее сохранённый продукт по имени и обозначению.
        /// Используется для восстановления путей к изображениям при повторном связывании с КОМПАС.
        /// </summary>
        public Product TryLoadSavedProduct(Product product)
        {
            if (product == null || string.IsNullOrEmpty(product.Name))
                return null;

            try
            {
                string folderName = GetProductFolderName(product);
                return Load(folderName);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось загрузить сохранённое изделие: {ex.Message}");
                return null;
            }
        }

        #region Private Methods

        private class ProductMeta
        {
            public string Stamp;
            public string Name;
            public string Marking;
            public int DetailsCount;
            public DateTime SavedDate;
        }

        /// <summary>
        /// Возвращает метаданные изделия; product.json перечитывается только если файл изменился
        /// </summary>
        private ProductMeta GetMeta(string folder)
        {
            string jsonPath = Path.Combine(folder, ProductJsonFileName);
            var info = new FileInfo(jsonPath);
            if (!info.Exists)
                return null;

            string stamp = info.LastWriteTimeUtc.Ticks + ":" + info.Length;

            lock (_metaCache)
            {
                ProductMeta cached;
                if (_metaCache.TryGetValue(folder, out cached) && cached.Stamp == stamp)
                    return cached;
            }

            // Без папки изделия: пути к изображениям для метаданных не нужны
            var product = LoadFromFile(jsonPath, null);
            if (product == null)
                return null;

            var meta = new ProductMeta
            {
                Stamp = stamp,
                Name = product.Name,
                Marking = product.Marking,
                DetailsCount = product.Details.Count,
                SavedDate = info.LastWriteTime
            };

            lock (_metaCache)
            {
                _metaCache[folder] = meta;
            }

            return meta;
        }

        /// <summary>
        /// Имя папки изделия должно быть простым именем без разделителей и служебных сегментов
        /// (значения приходят в том числе из файлов на общей папке)
        /// </summary>
        private static bool IsSafeFolderName(string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName) || folderName == "." || folderName == "..")
                return false;

            return folderName.IndexOfAny(Path.GetInvalidFileNameChars()) < 0;
        }

        private List<TombstoneEntry> ReadTombstones(string path)
        {
            try
            {
                if (!File.Exists(path))
                    return new List<TombstoneEntry>();

                var serializer = new DataContractJsonSerializer(typeof(TombstoneFile));
                using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var file = (TombstoneFile)serializer.ReadObject(fileStream);
                    return file?.Items ?? new List<TombstoneEntry>();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось прочитать список удалений {path}: {ex.Message}");
                return new List<TombstoneEntry>();
            }
        }

        private void WriteTombstones(string path, List<TombstoneEntry> items)
        {
            var serializer = new DataContractJsonSerializer(typeof(TombstoneFile));
            using (var memoryStream = new MemoryStream())
            {
                serializer.WriteObject(memoryStream, new TombstoneFile { Items = items });
                AtomicFile.WriteAllBytes(path, memoryStream.ToArray());
            }
        }

        private void RecordTombstone(string path, string folderName)
        {
            var items = ReadTombstones(path);
            items.RemoveAll(i => string.Equals(i.FolderName, folderName, StringComparison.OrdinalIgnoreCase));
            items.Add(new TombstoneEntry { FolderName = folderName, DeletedUtcTicks = DateTime.UtcNow.Ticks });
            WriteTombstones(path, items);
        }

        private static bool IsNewerThan(string filePath, DateTime utc)
        {
            return File.Exists(filePath) && File.GetLastWriteTimeUtc(filePath) > utc;
        }

        /// <summary>
        /// Переносит отложенные удаления на сервер и применяет список удалений к локальной и серверной папкам.
        /// Возвращает имена папок, которые нельзя синхронизировать в этот раз.
        /// </summary>
        private HashSet<string> ProcessTombstones(SyncResult result)
        {
            var skip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string serverPath = Path.Combine(_serverStorageFolder, TombstonesFileName);
            string pendingPath = Path.Combine(ProductsDirectory, PendingTombstonesFileName);

            var items = ReadTombstones(serverPath);
            var pending = ReadTombstones(pendingPath);
            bool changed = false;

            foreach (var entry in pending)
            {
                var existing = items.FirstOrDefault(i => string.Equals(i.FolderName, entry.FolderName, StringComparison.OrdinalIgnoreCase));
                if (existing == null)
                {
                    items.Add(entry);
                    changed = true;
                }
                else if (entry.DeletedUtcTicks > existing.DeletedUtcTicks)
                {
                    existing.DeletedUtcTicks = entry.DeletedUtcTicks;
                    changed = true;
                }
            }

            long cutoff = DateTime.UtcNow.AddDays(-TombstoneRetentionDays).Ticks;
            var keep = new List<TombstoneEntry>();

            foreach (var entry in items)
            {
                if (!IsSafeFolderName(entry.FolderName) || entry.DeletedUtcTicks < cutoff)
                {
                    changed = true;
                    continue;
                }

                var deletedUtc = new DateTime(entry.DeletedUtcTicks, DateTimeKind.Utc);
                string serverFolder = Path.Combine(_serverStorageFolder, entry.FolderName);
                string localFolder = Path.Combine(ProductsDirectory, entry.FolderName);

                // Изделие сохранили заново после удаления — запись об удалении устарела
                if (IsNewerThan(Path.Combine(serverFolder, ProductJsonFileName), deletedUtc) ||
                    IsNewerThan(Path.Combine(localFolder, ProductJsonFileName), deletedUtc))
                {
                    changed = true;
                    continue;
                }

                try
                {
                    DeleteDirectoryRecursive(localFolder);
                    DeleteDirectoryRecursive(serverFolder);
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"Не удалось применить удаление {entry.FolderName}: {ex.Message}");
                }

                skip.Add(entry.FolderName);
                keep.Add(entry);
            }

            if (changed)
                WriteTombstones(serverPath, keep);

            if (pending.Count > 0)
            {
                try { File.Delete(pendingPath); }
                catch (Exception ex) { _logger.LogWarning($"Не удалось удалить {pendingPath}: {ex.Message}"); }
            }

            return skip;
        }

        /// <summary>
        /// Путь из JSON (в том числе с общей папки) должен указывать внутрь локальной папки изделий
        /// </summary>
        private static string EnsureInsideProducts(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            try
            {
                string root = Path.GetFullPath(ProductsDirectory).TrimEnd('\\', '/') + Path.DirectorySeparatorChar;
                return Path.GetFullPath(path).StartsWith(root, StringComparison.OrdinalIgnoreCase) ? path : null;
            }
            catch
            {
                return null;
            }
        }

        private string GetProductFolderName(Product product)
        {
            string baseName = $"{product.Name}_{product.Marking}";
            return string.Join("_", baseName.Split(Path.GetInvalidFileNameChars()));
        }

        private string GenerateProductFolderName(Product product, string baseDirectory, string customName)
        {
            string baseName = customName ?? $"{product.Name}_{product.Marking}";
            // Убираем недопустимые символы
            baseName = string.Join("_", baseName.Split(Path.GetInvalidFileNameChars()));

            string folderPath = Path.Combine(baseDirectory, baseName);

            // Если папка существует, добавляем номер
            int counter = 1;
            while (Directory.Exists(folderPath))
            {
                string folderName = $"{baseName}_{counter}";
                folderPath = Path.Combine(baseDirectory, folderName);
                counter++;
            }

            return Path.GetFileName(folderPath);
        }

        private void SaveToFile(Product product, string filePath)
        {
            // Папка изделия нужна для сохранения относительных путей к изображениям
            string productFolder = Path.GetDirectoryName(filePath);

            var dto = ToDto(product, productFolder);
            var serializer = new DataContractJsonSerializer(typeof(ProductDto));

            using (var memoryStream = new MemoryStream())
            {
                serializer.WriteObject(memoryStream, dto);
                AtomicFile.WriteAllBytes(filePath, memoryStream.ToArray());
            }
        }

        private Product LoadFromFile(string filePath, string productFolder)
        {
            try
            {
                if (!File.Exists(filePath))
                    return null;

                var serializer = new DataContractJsonSerializer(typeof(ProductDto));

                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var memoryStream = new MemoryStream())
                {
                    fileStream.CopyTo(memoryStream);
                    memoryStream.Position = 0;
                    var dto = (ProductDto)serializer.ReadObject(memoryStream);
                    return FromDto(dto, productFolder);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка загрузки Product: {ex.Message}");
                return null;
            }
        }

        private static ProductDto ToDto(Product product, string productFolder)
        {
            return new ProductDto
            {
                Name = product.Name,
                Marking = product.Marking,
                Mass = product.Mass,
                FilePath = product.FilePath,
                Details = product.Details.Select(d => ToPartDto(d, productFolder)).ToList(),
                StandardParts = product.StandardParts.Select(d => ToPartDto(d, productFolder)).ToList(),
                SheetMaterials = product.SheetMaterials.Select(m => ToMaterialDto(m)).ToList(),
                TubularProducts = product.TubularProducts.Select(m => ToMaterialDto(m)).ToList(),
                OtherMaterials = product.OtherMaterials.Select(m => ToMaterialDto(m)).ToList()
            };
        }

        private static PartModelDto ToPartDto(PartModel part, string productFolder)
        {
            string relativeCdfPath = null;
            string relativeFilePreviewPath = null;

            // Преобразуем абсолютный путь в относительный для сохранения
            if (!string.IsNullOrEmpty(part.CdfFilePath) && !string.IsNullOrEmpty(productFolder))
            {
                relativeCdfPath = MakeRelativePath(part.CdfFilePath, productFolder);
            }

            // Преобразуем путь к превью 3D-файла
            if (!string.IsNullOrEmpty(part.FilePreviewPngPath) && !string.IsNullOrEmpty(productFolder))
            {
                relativeFilePreviewPath = MakeRelativePath(part.FilePreviewPngPath, productFolder);
            }

            return new PartModelDto
            {
                Name = part.Name,
                Marking = part.Marking,
                DetailType = part.DetailType,
                Material = part.Material,
                Mass = part.Mass,
                Length = part.Length,
                FilePath = part.FilePath,
                PartId = part.PartId,
                IsBodyBased = part.IsBodyBased,
                InstanceIndex = part.InstanceIndex,
                ProductType = (int)part.ProductType,
                CdfFilePath = relativeCdfPath,
                SourceCdwPath = part.SourceCdwPath,
                FilePreviewPngPath = relativeFilePreviewPath,
                DxfFilePath = part.DxfFilePath,
                MetalCost = part.MetalCost,
                OperationsCost = part.OperationsCost,
                TotalCost = part.TotalCost,
                Operations = part.Operations.Select(ToOperationDto).ToList()
            };
        }

        private static MaterialInfoDto ToMaterialDto(MaterialInfo material)
        {
            return new MaterialInfoDto
            {
                Name = material.Name,
                TotalMass = material.TotalMass,
                TotalLength = material.TotalLength
            };
        }

        private static OperationDto ToOperationDto(ManufacturingOperationBase op)
        {
            var dto = new OperationDto
            {
                Type = (int)op.Type,
                Cost = op.Cost
            };

            if (op is LaserCuttingOperation laser)
            {
                dto.CutLength = laser.CutLength;
                dto.EngravingLength = laser.EngravingLength;
            }
            else if (op is BendingOperation bend)
            {
                dto.BendAngle = bend.BendAngle;
                dto.BendLength = bend.BendLength;
            }
            else if (op is RollingOperation roll)
            {
                dto.RollDiameter = roll.RollDiameter;
                dto.Radius = roll.Radius;
                dto.Length = roll.Length;
            }
            else if (op is FlangingOperation flange)
            {
                dto.Diameter = flange.Diameter;
                dto.Radius = flange.Radius;
            }

            return dto;
        }

        private static MaterialInfo FromMaterialDto(MaterialInfoDto dto)
        {
            return new MaterialInfo
            {
                Name = dto.Name,
                TotalMass = dto.TotalMass,
                TotalLength = dto.TotalLength
            };
        }

        private static ManufacturingOperationBase FromOperationDto(OperationDto dto)
        {
            ManufacturingOperationBase result = null;
            var type = (ManufacturingOperationType)dto.Type;

            switch (type)
            {
                case ManufacturingOperationType.LaserCutting:
                    result = new LaserCuttingOperation
                    {
                        CutLength = dto.CutLength,
                        EngravingLength = dto.EngravingLength
                    };
                    break;
                case ManufacturingOperationType.Bending:
                    result = new BendingOperation
                    {
                        BendAngle = dto.BendAngle,
                        BendLength = dto.BendLength
                    };
                    break;
                case ManufacturingOperationType.Rolling:
                    result = new RollingOperation
                    {
                        RollDiameter = dto.RollDiameter,
                        Radius = dto.Radius,
                        Length = dto.Length
                    };
                    break;
                case ManufacturingOperationType.Flanging:
                    result = new FlangingOperation
                    {
                        Diameter = dto.Diameter,
                        Radius = dto.Radius
                    };
                    break;
            }

            if (result != null)
                result.Cost = dto.Cost;

            return result;
        }

        /// <summary>
        /// Преобразует относительный или абсолютный путь из DTO в абсолютный путь.
        /// Если путь абсолютный и файл существует — используем как есть.
        /// Если путь абсолютный но файл не существует — пробуем найти в папке продукта.
        /// Если путь относительный — разрешаем относительно папки продукта.
        /// </summary>
        private static string ResolveAbsolutePath(string path, string productFolder)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(productFolder))
                return path;

            try
            {
                if (!Path.IsPathRooted(path))
                {
                    // Относительный путь — разрешаем относительно папки продукта
                    string resolved = Path.Combine(productFolder, path);
                    if (File.Exists(resolved))
                        return resolved;

                    // Если не нашли напрямую, пробуем в подпапке images
                    string fileName = Path.GetFileName(path);
                    string inImages = Path.Combine(productFolder, ImagesSubfolder, fileName);
                    if (File.Exists(inImages))
                        return inImages;

                    // Возвращаем первый вариант даже если файл не существует
                    return resolved;
                }
                else
                {
                    // Абсолютный путь
                    if (File.Exists(path))
                        return path;

                    // Файл не существует по абсолютному пути — пробуем найти в папке продукта
                    string fileName = Path.GetFileName(path);

                    // Проверяем в images подпапке
                    string inImages = Path.Combine(productFolder, ImagesSubfolder, fileName);
                    if (File.Exists(inImages))
                        return inImages;

                    // Проверяем напрямую в папке продукта
                    string inFolder = Path.Combine(productFolder, fileName);
                    if (File.Exists(inFolder))
                        return inFolder;

                    // Ничего не нашли — возвращаем исходный путь
                    return path;
                }
            }
            catch
            {
                return path;
            }
        }

        private static Product FromDto(ProductDto dto, string productFolder)
        {
            var product = new Product();
            product.Name = dto.Name;
            product.Marking = dto.Marking;
            product.Mass = dto.Mass;
            product.FilePath = dto.FilePath;

            foreach (var partDto in dto.Details ?? Enumerable.Empty<PartModelDto>())
            {
                product.Details.Add(FromPartDto(partDto, productFolder));
            }

            foreach (var partDto in dto.StandardParts ?? Enumerable.Empty<PartModelDto>())
            {
                product.StandardParts.Add(FromPartDto(partDto, productFolder));
            }

            foreach (var materialDto in dto.SheetMaterials ?? Enumerable.Empty<MaterialInfoDto>())
            {
                product.SheetMaterials.Add(FromMaterialDto(materialDto));
            }

            foreach (var materialDto in dto.TubularProducts ?? Enumerable.Empty<MaterialInfoDto>())
            {
                product.TubularProducts.Add(FromMaterialDto(materialDto));
            }

            foreach (var materialDto in dto.OtherMaterials ?? Enumerable.Empty<MaterialInfoDto>())
            {
                product.OtherMaterials.Add(FromMaterialDto(materialDto));
            }

            return product;
        }

        private static PartModel FromPartDto(PartModelDto dto, string productFolder)
        {
            string absoluteCdfPath = null;
            string absoluteFilePreviewPath = null;

            // Преобразуем относительный путь в абсолютный
            if (!string.IsNullOrEmpty(dto.CdfFilePath) && !string.IsNullOrEmpty(productFolder))
            {
                absoluteCdfPath = EnsureInsideProducts(ResolveAbsolutePath(dto.CdfFilePath, productFolder));
            }

            // Преобразуем путь к превью 3D-файла
            if (!string.IsNullOrEmpty(dto.FilePreviewPngPath) && !string.IsNullOrEmpty(productFolder))
            {
                absoluteFilePreviewPath = EnsureInsideProducts(ResolveAbsolutePath(dto.FilePreviewPngPath, productFolder));
            }

            var part = new PartModelFromStorage
            {
                Name = dto.Name,
                Marking = dto.Marking,
                DetailType = dto.DetailType,
                Material = dto.Material,
                Mass = dto.Mass,
                Length = dto.Length,
                FilePath = dto.FilePath,
                PartId = dto.PartId,
                IsBodyBased = dto.IsBodyBased,
                InstanceIndex = dto.InstanceIndex,
                ProductType = (ProductType)(dto.ProductType),
                CdfFilePath = absoluteCdfPath,
                SourceCdwPath = dto.SourceCdwPath,
                FilePreviewPngPath = absoluteFilePreviewPath,
                DxfFilePath = dto.DxfFilePath,
                MetalCost = dto.MetalCost,
                OperationsCost = dto.OperationsCost,
                TotalCost = dto.TotalCost
            };

            if (dto.Operations != null)
            {
                foreach (var opDto in dto.Operations)
                {
                    var op = FromOperationDto(opDto);
                    if (op != null)
                        part.Operations.Add(op);
                }
            }

            return part;
        }

        /// <summary>
        /// Преобразует относительный путь в абсолютный для сохранения.
        /// Если путь находится внутри папки продукта — делает его относительным.
        /// Если путь находится в другой папке products — сохраняет только относительную часть images\filename.
        /// </summary>
        private static string MakeRelativePath(string absolutePath, string productFolder)
        {
            if (string.IsNullOrEmpty(absolutePath))
                return null;

            try
            {
                // Если путь уже находится внутри папки продукта, делаем его относительным
                if (absolutePath.StartsWith(productFolder, StringComparison.OrdinalIgnoreCase))
                {
                    return absolutePath.Substring(productFolder.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                }

                // Если путь содержит images\filename — извлекаем относительную часть
                // Это нужно когда путь указывает на локальную папку products, а сохраняем в серверную
                string fileName = Path.GetFileName(absolutePath);
                string parentDir = Path.GetFileName(Path.GetDirectoryName(absolutePath) ?? "");
                if (string.Equals(parentDir, ImagesSubfolder, StringComparison.OrdinalIgnoreCase))
                {
                    return Path.Combine(ImagesSubfolder, fileName);
                }

                // Сохраняем только имя файла как fallback
                return fileName;
            }
            catch
            {
                return absolutePath;
            }
        }

        #endregion
    }

    #region DTO и вспомогательные классы

    /// <summary>
    /// Информация о сохранённом файле продукта
    /// </summary>
    public class ProductFileInfo
    {
        public string FileName { get; set; }
        public string ProductName { get; set; }
        public string Marking { get; set; }
        public int DetailsCount { get; set; }
        public DateTime SavedDate { get; set; }

        public string DisplayName => $"{ProductName} ({Marking}) - {DetailsCount} дет.";
    }

    /// <summary>
    /// Настройки хранения
    /// </summary>
    [System.Runtime.Serialization.DataContract]
    public class StorageSettings
    {
        [System.Runtime.Serialization.DataMember]
        public string ServerStorageFolder { get; set; }
    }

    [System.Runtime.Serialization.DataContract]
    public class TombstoneEntry
    {
        [System.Runtime.Serialization.DataMember]
        public string FolderName { get; set; }

        [System.Runtime.Serialization.DataMember]
        public long DeletedUtcTicks { get; set; }
    }

    [System.Runtime.Serialization.DataContract]
    public class TombstoneFile
    {
        [System.Runtime.Serialization.DataMember]
        public List<TombstoneEntry> Items { get; set; }
    }

    [System.Runtime.Serialization.DataContract]
    public class ProductDto
    {
        [System.Runtime.Serialization.DataMember]
        public string Name { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string Marking { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double Mass { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string FilePath { get; set; }

        [System.Runtime.Serialization.DataMember]
        public List<PartModelDto> Details { get; set; }

        [System.Runtime.Serialization.DataMember]
        public List<PartModelDto> StandardParts { get; set; }

        [System.Runtime.Serialization.DataMember]
        public List<MaterialInfoDto> SheetMaterials { get; set; }

        [System.Runtime.Serialization.DataMember]
        public List<MaterialInfoDto> TubularProducts { get; set; }

        [System.Runtime.Serialization.DataMember]
        public List<MaterialInfoDto> OtherMaterials { get; set; }
    }

    [System.Runtime.Serialization.DataContract]
    public class PartModelDto
    {
        [System.Runtime.Serialization.DataMember]
        public string Name { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string Marking { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string DetailType { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string Material { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double Mass { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string FilePath { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string PartId { get; set; }

        [System.Runtime.Serialization.DataMember]
        public bool IsBodyBased { get; set; }

        [System.Runtime.Serialization.DataMember]
        public int InstanceIndex { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double Length { get; set; }

        [System.Runtime.Serialization.DataMember]
        public int ProductType { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string CdfFilePath { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string SourceCdwPath { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string FilePreviewPngPath { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string DxfFilePath { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double MetalCost { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double OperationsCost { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double TotalCost { get; set; }

        [System.Runtime.Serialization.DataMember]
        public List<OperationDto> Operations { get; set; }
    }

    [System.Runtime.Serialization.DataContract]
    public class MaterialInfoDto
    {
        [System.Runtime.Serialization.DataMember]
        public string Name { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double TotalMass { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double TotalLength { get; set; }
    }

    [System.Runtime.Serialization.DataContract]
    public class OperationDto
    {
        [System.Runtime.Serialization.DataMember]
        public int Type { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double Cost { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double CutLength { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double EngravingLength { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double BendAngle { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double BendLength { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double RollDiameter { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double Radius { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double Length { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double Diameter { get; set; }
    }

    #endregion
}
