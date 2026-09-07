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
    /// ��������� �������������
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
    /// ������ ��� ���������� � �������� Product � ��������� ���� � �������������� � ��������
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

        private readonly ImageSyncService _imageSyncService = new ImageSyncService();
        private string _serverStorageFolder;

        /// <summary>
        /// ��������� (�������) ����� ��� �������� �������
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
        /// ���������, ����������� �� ��������� �����
        /// </summary>
        public bool HasServerFolder => !string.IsNullOrEmpty(_serverStorageFolder);

        /// <summary>
        /// ���������, �������� �� ��������� �����
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
                System.Diagnostics.Debug.WriteLine($"������ �������� �������� ��������: {ex.Message}");
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
                System.Diagnostics.Debug.WriteLine($"������ ���������� �������� ��������: {ex.Message}");
            }
        }

        #endregion

        #region Synchronization

        /// <summary>
        /// �������������� ������ ����� ��������� ������ � �������� (������������ �������������).
        /// �������� ����� � ���������� ������� � ��� �������.
        /// </summary>
        /// <param name="skipImages">���� true, ����������� �� ���������� ��� �������������</param>
        public SyncResult SyncFromServer(bool skipImages = false)
        {
            var result = new SyncResult();

            if (!IsServerAvailable)
            {
                if (HasServerFolder)
                    result.Errors.Add("��������� ����� ����������");
                return result;
            }

            try
            {
                // ���� 1: ������������� � ������� � ��������� �����
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
                            // ����� ������� - ����� �����������
                            needsCopy = true;
                            isNew = true;
                        }
                        else
                        {
                            // ��������� ���� �����������
                            var localFileInfo = new FileInfo(localJsonPath);
                            if (serverFileInfo.LastWriteTimeUtc > localFileInfo.LastWriteTimeUtc)
                            {
                                // ��������� ������ ����� - ����� ��������
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
                        result.Errors.Add($"������ ������������� � ������� {Path.GetFileName(serverFolder)}: {ex.Message}");
                    }
                }

                // ���� 2: ������������� �� ��������� ����� �� ������
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

                        // ���������� ����� ��� product.json (��� ����� ������ � ������)
                        if (!File.Exists(localJsonPath))
                            continue;

                        var localFileInfo = new FileInfo(localJsonPath);
                        string serverJsonPath = Path.Combine(serverFolder, ProductJsonFileName);

                        bool needsCopy = false;

                        if (!Directory.Exists(serverFolder) || !File.Exists(serverJsonPath))
                        {
                            // ����� ��������� ������� - ��������� �� ������
                            needsCopy = true;
                        }
                        else
                        {
                            // ��������� ���� �����������
                            var serverFileInfo = new FileInfo(serverJsonPath);
                            if (localFileInfo.LastWriteTimeUtc > serverFileInfo.LastWriteTimeUtc)
                            {
                                // ��������� ������ ����� - ��������� �� ������
                                needsCopy = true;
                            }
                        }

                        if (needsCopy)
                        {
                            CopyProductFolder(localFolder, serverFolder, skipImages);
                            // �� ����������� ��������, ��� ��� ��� ��������� � ������ ����
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add($"������ �������� �� ������ {Path.GetFileName(localFolder)}: {ex.Message}");
                    }
                }

                // ���� 3: ������������� ����������� �� ������ ��������� ������
                if (!skipImages)
                    SyncAllProductImages();
            }
            catch (Exception ex)
            {
                result.Errors.Add($"������ ������� � ������: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// �������� ����� �������� � ������� � ��������� ����������
        /// </summary>
        /// <param name="skipImages">���� true, �������� images �� ����������</param>
        private void CopyProductFolder(string sourceFolder, string destFolder, bool skipImages = false)
        {
            // ������ ������� �����
            Directory.CreateDirectory(destFolder);

            // �������� ��� �����
            foreach (var file in Directory.GetFiles(sourceFolder))
            {
                string destFile = Path.Combine(destFolder, Path.GetFileName(file));
                File.Copy(file, destFile, true);
            }

            // �������� �������� (���������� images ���� skipImages)
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
        /// �������������� ����������� ��� ���� ��������� ����� ��������� ������ � ��������
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
                        System.Diagnostics.Debug.WriteLine($"������ ������������� ����������� {Path.GetFileName(localFolder)}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"������ ������������� �����������: {ex.Message}");
            }
        }

        /// <summary>
        /// �������������� ����������� ����������� �������� ����� ��������� ������ � ��������
        /// </summary>
        public void SyncProductImagesWithServer(Product product)
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
                System.Diagnostics.Debug.WriteLine($"������ ������������� ����������� ��������: {ex.Message}");
            }
        }

        /// <summary>
        /// ���������� ���� � ����� � ������������� ��� ��������
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
        /// ��������� ��� ��������� �������� ������� (��������������)
        /// </summary>
        public void SaveAsLast(Product product)
        {
            if (product == null || string.IsNullOrEmpty(product.Name)) return;
            
            // ���������� ����� �������� ��� ����������� ���������� ������������� �����
            string productFolderName = GetProductFolderName(product);
            string productFolder = Path.Combine(ProductsDirectory, productFolderName);
            
            SaveToFile(product, Path.Combine(ProductsDirectory, LastProductFileName), productFolder);
        }

        /// <summary>
        /// ��������� ��������� �������� �������
        /// </summary>
        public Product LoadLast()
        {
            string filePath = Path.Combine(ProductsDirectory, LastProductFileName);
            if (!File.Exists(filePath))
                return null;

            // ������� ��������� ��� �����, ����� ������ ��� ��������
            var product = LoadFromFile(filePath, null);
            if (product == null)
                return null;

            // ������� ����� �������� ��� ����������� ���������� ����� � ������������
            string productFolderName = GetProductFolderName(product);
            string productFolder = Path.Combine(ProductsDirectory, productFolderName);
            
            if (Directory.Exists(productFolder))
            {
                // ������������� � ���������� ������ ��� ���������� �����
                return LoadFromFile(filePath, productFolder);
            }
            
            return product;
        }

        /// <summary>
        /// ��������� ������� � ��������� ����� � �� ������ (���� ��������)
        /// </summary>
        public string Save(Product product, string customName = null)
        {
            if (product == null) return null;

            // ��������� � ��������� �����
            string localFilePath = SaveToDirectory(product, ProductsDirectory, customName);

            // ���� ������ �������� - ��������� � ����
            if (IsServerAvailable)
            {
                try
                {
                    SaveToDirectory(product, _serverStorageFolder, customName);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"������ ���������� �� ������: {ex.Message}");
                }
            }

            return localFilePath;
        }

        /// <summary>
        /// ��������� ������� � ��������� ����������
        /// </summary>
        private string SaveToDirectory(Product product, string baseDirectory, string customName)
        {
            // ���������, ���� �� ��� ����������� �������
            string existingFolderPath = FindExistingProductFolder(product, baseDirectory);
            
            if (existingFolderPath != null)
            {
                // �������������� ������������ �������
                string existingFilePath = Path.Combine(existingFolderPath, ProductJsonFileName);
                SaveToFile(product, existingFilePath);
                
                // �������� ����������� ���� ��� ���� � ��������� �����
                CopyImagesIfNeeded(product, existingFolderPath);
                
                return existingFilePath;
            }

            // ������ ����� ����� ��� ��������
            string productFolderName = GenerateProductFolderName(product, baseDirectory, customName);
            string productFolderPath = Path.Combine(baseDirectory, productFolderName);
            Directory.CreateDirectory(productFolderPath);
            
            // ������ ����� ��� �����������
            string imagesFolder = Path.Combine(productFolderPath, ImagesSubfolder);
            Directory.CreateDirectory(imagesFolder);

            // ��������� JSON
            string filePath = Path.Combine(productFolderPath, ProductJsonFileName);
            SaveToFile(product, filePath);
            
            // �������� ����������� ���� ��� ���� � ��������� �����
            CopyImagesIfNeeded(product, productFolderPath);
            
            return filePath;
        }

        /// <summary>
        /// �������� ����������� �� ��������� ����� �������� � ������� �����
        /// </summary>
        private void CopyImagesIfNeeded(Product product, string destProductFolder)
        {
            // ������� ��������� ����� ��������
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
                catch { /* ���������� ������ ����������� */ }
            }
        }

        /// <summary>
        /// ����� ������������ ����� �������� �� ����� � ����������� � ��������� ����������
        /// </summary>
        private string FindExistingProductFolder(Product product, string baseDirectory)
        {
            if (!Directory.Exists(baseDirectory))
                return null;

            var folders = Directory.GetDirectories(baseDirectory)
                .Where(f => !Path.GetFileName(f).StartsWith("_"));

            // ������� ���� ����� � product.json (��������� ����������� �������)
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
                    // ���������� ����������� �����
                }
            }

            // ���� �� ����� ����������� �������, ���� ����� �� ����� (��� ������)
            string expectedFolderName = GetProductFolderName(product);
            string expectedFolder = Path.Combine(baseDirectory, expectedFolderName);
            
            if (Directory.Exists(expectedFolder))
            {
                return expectedFolder;
            }

            return null;
        }

        /// <summary>
        /// ��������� ������� �� ����� ����� (������ �� ��������� �����)
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
        /// �������� ������ ���� ����������� ��������� (������ �� ��������� �����)
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
                    // ���������� ����������� �����
                }
            }

            return result.OrderByDescending(p => p.SavedDate).ToList();
        }

        /// <summary>
        /// ������� ������� ������ �� ��������� �����
        /// </summary>
        public bool DeleteLocal(string folderName)
        {
            // �������������� ������ ������ ����� ���������
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

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
                System.Diagnostics.Debug.WriteLine($"������ �������� �� ��������� �����: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// ������� ������� �� ���� (�� ��������� ����� � � �������)
        /// </summary>
        public bool Delete(string folderName)
        {
            bool deletedAny = DeleteLocal(folderName);

            // ������� � ������� ���� ��������
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
                    System.Diagnostics.Debug.WriteLine($"������ �������� � �������: {ex.Message}");
                }
            }

            return deletedAny;
        }

        /// <summary>
        /// ����������� �������� ���������� � ���������� ���������
        /// </summary>
        private void DeleteDirectoryRecursive(string path)
        {
            if (!Directory.Exists(path))
                return;

            System.Diagnostics.Debug.WriteLine($"??? ������� ��������: {path}");

            // ���������� ��������������� ������� ��� ���������� �������
            if (!FileLockDiagnostics.ForceDeleteDirectory(path, maxAttempts: 5, delayMs: 200))
            {
                System.Diagnostics.Debug.WriteLine($"?? �� ������� ������� ����� ����� ���� �������");
                
                // ������� ��� ��� � ����� ���������� ���������
                System.Threading.Thread.Sleep(1000);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                try
                {
                    Directory.Delete(path, true);
                    System.Diagnostics.Debug.WriteLine($"? ����� ������� ����� �������������� ��������");
                }
                catch (IOException ex)
                {
                    System.Diagnostics.Debug.WriteLine($"? ������������� ������ ��������: {ex.Message}");
                    throw; // ������������ ���������� ����
                }
            }
        }

        /// <summary>
        /// ��������� ������������� ��������
        /// </summary>
        public bool Exists(string folderName)
        {
            string folderPath = Path.Combine(ProductsDirectory, folderName);
            string jsonPath = Path.Combine(folderPath, ProductJsonFileName);
            return File.Exists(jsonPath);
        }

        /// <summary>
        /// �������� ��������� ����� ����������� ������� �� ����� � �����������.
        /// ������������ ��� �������������� ����� � ������������ ��� ��������� ���������� � ������.
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
            catch
            {
                return null;
            }
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
            // ������� ������������ �������
            baseName = string.Join("_", baseName.Split(Path.GetInvalidFileNameChars()));

            string folderPath = Path.Combine(baseDirectory, baseName);

            // ���� ����� ����������, ��������� �����
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
            SaveToFile(product, filePath, null);
        }

        private void SaveToFile(Product product, string filePath, string productFolderOverride)
        {
            try
            {
                // �������� ����� �������� ��� ���������� ������������� �����
                string productFolder = productFolderOverride ?? Path.GetDirectoryName(filePath);
                
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
                System.Diagnostics.Debug.WriteLine($"������ ���������� Product: {ex.Message}");
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
                System.Diagnostics.Debug.WriteLine($"������ �������� Product: {ex.Message}");
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
            
            // ����������� ���������� ���� � ������������� ��� ����������
            if (!string.IsNullOrEmpty(part.CdfFilePath) && !string.IsNullOrEmpty(productFolder))
            {
                relativeCdfPath = MakeRelativePath(part.CdfFilePath, productFolder);
            }

            // ����������� ���� � ������ 3D-�����
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
        /// ����������� ������������� ��� ���������� ���� �� DTO � ���������� ����.
        /// ���� ���� ���������� � ���� ���������� � ���������� ��� ����.
        /// ���� ���� ���������� �� ���� �� ���������� � ������� ����� � ����� ��������.
        /// ���� ���� ������������� � ��������� ������������ ����� ��������.
        /// </summary>
        private static string ResolveAbsolutePath(string path, string productFolder)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(productFolder))
                return path;

            try
            {
                if (!Path.IsPathRooted(path))
                {
                    // ������������� ���� � ��������� ������������ ����� ��������
                    string resolved = Path.Combine(productFolder, path);
                    if (File.Exists(resolved))
                        return resolved;

                    // ���� �� ����� ��������, ������� � �������� images
                    string fileName = Path.GetFileName(path);
                    string inImages = Path.Combine(productFolder, ImagesSubfolder, fileName);
                    if (File.Exists(inImages))
                        return inImages;

                    // ���������� ������ ������� ���� ���� ���� �� ����������
                    return resolved;
                }
                else
                {
                    // ���������� ����
                    if (File.Exists(path))
                        return path;

                    // ���� �� ���������� �� ����������� ���� � ������� ����� � ����� ��������
                    string fileName = Path.GetFileName(path);
                    
                    // ��������� � images ��������
                    string inImages = Path.Combine(productFolder, ImagesSubfolder, fileName);
                    if (File.Exists(inImages))
                        return inImages;

                    // ��������� �������� � ����� ��������
                    string inFolder = Path.Combine(productFolder, fileName);
                    if (File.Exists(inFolder))
                        return inFolder;

                    // ������ �� ����� � ���������� �������� ����
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
            
            // ����������� ������������� ���� � ����������
            if (!string.IsNullOrEmpty(dto.CdfFilePath) && !string.IsNullOrEmpty(productFolder))
            {
                absoluteCdfPath = ResolveAbsolutePath(dto.CdfFilePath, productFolder);
            }

            // ����������� ���� � ������ 3D-�����
            if (!string.IsNullOrEmpty(dto.FilePreviewPngPath) && !string.IsNullOrEmpty(productFolder))
            {
                absoluteFilePreviewPath = ResolveAbsolutePath(dto.FilePreviewPngPath, productFolder);
            }
            
            var part = new PartModelFromStorage
            {
                Name = dto.Name,
                Marking = dto.Marking,
                DetailType = dto.DetailType,
                Material = dto.Material,
                Mass = dto.Mass,
                FilePath = dto.FilePath,
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
        /// ����������� ������������� ���� � ���������� ��� ����������.
        /// ���� ���� ��������� ������ ����� �������� � ������ ��� �������������.
        /// ���� ���� ��������� � ������ ����� products � ��������� ������ ������������� ����� images\filename.
        /// </summary>
        private static string MakeRelativePath(string absolutePath, string productFolder)
        {
            if (string.IsNullOrEmpty(absolutePath))
                return null;

            try
            {
                // ���� ���� ��� ��������� ������ ����� ��������, ������ ��� �������������
                if (absolutePath.StartsWith(productFolder, StringComparison.OrdinalIgnoreCase))
                {
                    return absolutePath.Substring(productFolder.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                }

                // ���� ���� �������� images\filename � ��������� ������������� �����
                // ��� ����� ����� ���� ��������� �� ��������� ����� products, � ��������� � ���������
                string fileName = Path.GetFileName(absolutePath);
                string parentDir = Path.GetFileName(Path.GetDirectoryName(absolutePath) ?? "");
                if (string.Equals(parentDir, ImagesSubfolder, StringComparison.OrdinalIgnoreCase))
                {
                    return Path.Combine(ImagesSubfolder, fileName);
                }

                // ��������� ������ ��� ����� ��� fallback
                return fileName;
            }
            catch
            {
                return absolutePath;
            }
        }

        #endregion
    }

    #region DTO � ��������������� ������

    /// <summary>
    /// ���������� � ����������� ����� ��������
    /// </summary>
    public class ProductFileInfo
    {
        public string FileName { get; set; }
        public string ProductName { get; set; }
        public string Marking { get; set; }
        public int DetailsCount { get; set; }
        public DateTime SavedDate { get; set; }

        public string DisplayName => $"{ProductName} ({Marking}) - {DetailsCount} ���.";
    }

    /// <summary>
    /// ��������� ��������
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
