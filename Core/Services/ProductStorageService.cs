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

        private const string LastProductFileName = "_last_product.json";
        private const string ProductJsonFileName = "product.json";
        private const string ImagesSubfolder = "images";
        private const string FileExtension = ".json";

        private string _serverStorageFolder;

        /// <summary>
        /// Серверная (сетевая) папка для хранения изделий
        /// </summary>
        public string ServerStorageFolder
        {
            get => _serverStorageFolder;
            set
            {
                _serverStorageFolder = value;
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
        public bool IsServerAvailable => HasServerFolder && Directory.Exists(_serverStorageFolder);

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
                System.Diagnostics.Debug.WriteLine($"Ошибка загрузки настроек хранения: {ex.Message}");
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
                    memoryStream.Position = 0;
                    
                    using (var fileStream = new FileStream(SettingsFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        memoryStream.CopyTo(fileStream);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка сохранения настроек хранения: {ex.Message}");
            }
        }

        #endregion

        #region Synchronization

        /// <summary>
        /// Синхронизирует данные между локальной папкой и сервером (двусторонняя синхронизация).
        /// Копирует новые и обновлённые изделия в обе стороны.
        /// </summary>
        public SyncResult SyncFromServer()
        {
            var result = new SyncResult();

            if (!IsServerAvailable)
            {
                if (HasServerFolder)
                    result.Errors.Add("Серверная папка недоступна");
                return result;
            }

            try
            {
                // Фаза 1: Синхронизация С СЕРВЕРА В ЛОКАЛЬНУЮ ПАПКУ
                var serverFolders = Directory.GetDirectories(_serverStorageFolder)
                    .Where(f => !Path.GetFileName(f).StartsWith("_"))
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
                            CopyProductFolder(serverFolder, localFolder);
                            
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

                // Фаза 2: Синхронизация ИЗ ЛОКАЛЬНОЙ ПАПКИ НА СЕРВЕР
                var localFolders = Directory.GetDirectories(ProductsDirectory)
                    .Where(f => !Path.GetFileName(f).StartsWith("_"))
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
                            CopyProductFolder(localFolder, serverFolder);
                            // Не увеличиваем счётчики, так как уже посчитали в первой фазе
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add($"Ошибка загрузки на сервер {Path.GetFileName(localFolder)}: {ex.Message}");
                    }
                }
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
        private void CopyProductFolder(string sourceFolder, string destFolder)
        {
            // Создаём целевую папку
            Directory.CreateDirectory(destFolder);

            // Копируем все файлы
            foreach (var file in Directory.GetFiles(sourceFolder))
            {
                string destFile = Path.Combine(destFolder, Path.GetFileName(file));
                File.Copy(file, destFile, true);
            }

            // Копируем подпапки (например, images)
            foreach (var dir in Directory.GetDirectories(sourceFolder))
            {
                string destSubDir = Path.Combine(destFolder, Path.GetFileName(dir));
                CopyProductFolder(dir, destSubDir);
            }
        }

        #endregion

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
        /// Сохранить как последний открытый продукт (автосохранение)
        /// </summary>
        public void SaveAsLast(Product product)
        {
            if (product == null || string.IsNullOrEmpty(product.Name)) return;
            SaveToFile(product, Path.Combine(ProductsDirectory, LastProductFileName));
        }

        /// <summary>
        /// Загрузить последний открытый продукт
        /// </summary>
        public Product LoadLast()
        {
            string filePath = Path.Combine(ProductsDirectory, LastProductFileName);
            return LoadFromFile(filePath, null);
        }

        /// <summary>
        /// Сохранить продукт в локальную папку и на сервер (если доступен)
        /// </summary>
        public string Save(Product product, string customName = null)
        {
            if (product == null) return null;

            // Сохраняем в локальную папку
            string localFilePath = SaveToDirectory(product, ProductsDirectory, customName);

            // Если сервер доступен - сохраняем и туда
            if (IsServerAvailable)
            {
                try
                {
                    SaveToDirectory(product, _serverStorageFolder, customName);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Ошибка сохранения на сервер: {ex.Message}");
                }
            }

            return localFilePath;
        }

        /// <summary>
        /// Сохранить продукт в указанную директорию
        /// </summary>
        private string SaveToDirectory(Product product, string baseDirectory, string customName)
        {
            // Проверяем, есть ли уже сохранённый продукт
            string existingFolderPath = FindExistingProductFolder(product, baseDirectory);
            
            if (existingFolderPath != null)
            {
                // Перезаписываем существующий продукт
                string existingFilePath = Path.Combine(existingFolderPath, ProductJsonFileName);
                SaveToFile(product, existingFilePath);
                
                // Копируем изображения если они есть в локальной папке
                CopyImagesIfNeeded(product, existingFolderPath);
                
                return existingFilePath;
            }

            // Создаём новую папку для продукта
            string productFolderName = GenerateProductFolderName(product, baseDirectory, customName);
            string productFolderPath = Path.Combine(baseDirectory, productFolderName);
            Directory.CreateDirectory(productFolderPath);
            
            // Создаём папку для изображений
            string imagesFolder = Path.Combine(productFolderPath, ImagesSubfolder);
            Directory.CreateDirectory(imagesFolder);

            // Сохраняем JSON
            string filePath = Path.Combine(productFolderPath, ProductJsonFileName);
            SaveToFile(product, filePath);
            
            // Копируем изображения если они есть в локальной папке
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
                if (!File.Exists(destFile))
                {
                    try
                    {
                        File.Copy(file, destFile, false);
                    }
                    catch { /* Игнорируем ошибки копирования */ }
                }
            }
        }

        /// <summary>
        /// Найти существующую папку продукта по имени и обозначению в указанной директории
        /// </summary>
        private string FindExistingProductFolder(Product product, string baseDirectory)
        {
            if (!Directory.Exists(baseDirectory))
                return null;

            var folders = Directory.GetDirectories(baseDirectory)
                .Where(f => !Path.GetFileName(f).StartsWith("_"));

            // Сначала ищем папку с product.json (полностью сохранённый продукт)
            foreach (var folder in folders)
            {
                try
                {
                    string jsonPath = Path.Combine(folder, ProductJsonFileName);
                    if (!File.Exists(jsonPath))
                        continue;

                    var existingProduct = LoadFromFile(jsonPath, folder);
                    if (existingProduct != null &&
                        existingProduct.Name == product.Name &&
                        existingProduct.Marking == product.Marking)
                    {
                        return folder;
                    }
                }
                catch
                {
                    // Пропускаем повреждённые папки
                }
            }

            // Если не нашли сохранённый продукт, ищем папку по имени (для превью)
            string expectedFolderName = GetProductFolderName(product);
            string expectedFolder = Path.Combine(baseDirectory, expectedFolderName);
            
            if (Directory.Exists(expectedFolder))
            {
                return expectedFolder;
            }

            return null;
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
        /// Получить список всех сохранённых продуктов (только из локальной папки)
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
                    string jsonPath = Path.Combine(folder, ProductJsonFileName);
                    if (!File.Exists(jsonPath))
                        continue;

                    var fileInfo = new FileInfo(jsonPath);
                    var product = LoadFromFile(jsonPath, folder);

                    if (product != null)
                    {
                        result.Add(new ProductFileInfo
                        {
                            FileName = Path.GetFileName(folder),
                            ProductName = product.Name,
                            Marking = product.Marking,
                            DetailsCount = product.Details.Count,
                            SavedDate = fileInfo.LastWriteTime
                        });
                    }
                }
                catch
                {
                    // Пропускаем повреждённые папки
                }
            }

            return result.OrderByDescending(p => p.SavedDate).ToList();
        }

        /// <summary>
        /// Удалить продукт из базы (из локальной папки и с сервера)
        /// </summary>
        public bool Delete(string folderName)
        {
            bool deletedAny = false;

            // Принудительная сборка мусора перед удалением
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Удаляем из локальной папки
            try
            {
                string folderPath = Path.Combine(ProductsDirectory, folderName);
                if (Directory.Exists(folderPath))
                {
                    // Пытаемся удалить файлы по одному, чтобы определить заблокированные
                    DeleteDirectoryRecursive(folderPath);
                    deletedAny = true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка удаления из локальной папки: {ex.Message}");
            }

            // Удаляем с сервера если доступен
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
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Ошибка удаления с сервера: {ex.Message}");
                }
            }

            return deletedAny;
        }

        /// <summary>
        /// Рекурсивное удаление директории с повторными попытками
        /// </summary>
        private void DeleteDirectoryRecursive(string path)
        {
            if (!Directory.Exists(path))
                return;

            System.Diagnostics.Debug.WriteLine($"🗑️ Попытка удаления: {path}");

            // Используем диагностическую утилиту для детального анализа
            if (!FileLockDiagnostics.ForceDeleteDirectory(path, maxAttempts: 5, delayMs: 200))
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Не удалось удалить папку после всех попыток");
                
                // Пробуем ещё раз с более длительной задержкой
                System.Threading.Thread.Sleep(1000);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                try
                {
                    Directory.Delete(path, true);
                    System.Diagnostics.Debug.WriteLine($"✅ Папка удалена после дополнительной задержки");
                }
                catch (IOException ex)
                {
                    System.Diagnostics.Debug.WriteLine($"❌ Окончательная ошибка удаления: {ex.Message}");
                    throw; // Пробрасываем исключение выше
                }
            }
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

        #region Private Methods

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
            try
            {
                // Получаем папку продукта для сохранения относительных путей
                string productFolder = Path.GetDirectoryName(filePath);
                
                var dto = ToDto(product, productFolder);
                var serializer = new DataContractJsonSerializer(typeof(ProductDto));

                using (var memoryStream = new MemoryStream())
                {
                    serializer.WriteObject(memoryStream, dto);
                    memoryStream.Position = 0;
                    
                    using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        memoryStream.CopyTo(fileStream);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка сохранения Product: {ex.Message}");
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
                System.Diagnostics.Debug.WriteLine($"Ошибка загрузки Product: {ex.Message}");
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
            
            // Преобразуем абсолютный путь в относительный для сохранения в папке продукта
            if (!string.IsNullOrEmpty(part.CdfFilePath) && !string.IsNullOrEmpty(productFolder))
            {
                try
                {
                    // Если путь уже находится внутри папки продукта, делаем его относительным
                    if (part.CdfFilePath.StartsWith(productFolder, StringComparison.OrdinalIgnoreCase))
                    {
                        relativeCdfPath = part.CdfFilePath.Substring(productFolder.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    }
                    else
                    {
                        relativeCdfPath = part.CdfFilePath;
                    }
                }
                catch
                {
                    relativeCdfPath = part.CdfFilePath;
                }
            }
            
            return new PartModelDto
            {
                Name = part.Name,
                Marking = part.Marking,
                DetailType = part.DetailType,
                Material = part.Material,
                Mass = part.Mass,
                FilePath = part.FilePath,
                PartId = part.PartId,
                IsBodyBased = part.IsBodyBased,
                InstanceIndex = part.InstanceIndex,
                ProductType = (int)part.ProductType,
                CdfFilePath = relativeCdfPath,
                SourceCdwPath = part.SourceCdwPath
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
            
            // Преобразуем относительный путь в абсолютный
            if (!string.IsNullOrEmpty(dto.CdfFilePath) && !string.IsNullOrEmpty(productFolder))
            {
                try
                {
                    // Если путь относительный, делаем его абсолютным
                    if (!Path.IsPathRooted(dto.CdfFilePath))
                    {
                        absoluteCdfPath = Path.Combine(productFolder, dto.CdfFilePath);
                    }
                    else
                    {
                        absoluteCdfPath = dto.CdfFilePath;
                    }
                }
                catch
                {
                    absoluteCdfPath = dto.CdfFilePath;
                }
            }
            
            return new PartModelFromStorage
            {
                Name = dto.Name,
                Marking = dto.Marking,
                DetailType = dto.DetailType,
                Material = dto.Material,
                Mass = dto.Mass,
                FilePath = dto.FilePath,
                ProductType = (ProductType)(dto.ProductType),
                CdfFilePath = absoluteCdfPath,
                SourceCdwPath = dto.SourceCdwPath
            };
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
        public int ProductType { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string CdfFilePath { get; set; }

        [System.Runtime.Serialization.DataMember]
        public string SourceCdwPath { get; set; }
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

    #endregion
}