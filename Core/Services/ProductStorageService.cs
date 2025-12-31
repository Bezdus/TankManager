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
    /// Сервис для сохранения и загрузки Product в локальную базу
    /// </summary>
    public class ProductStorageService
    {
        private static readonly string ProductsDirectory =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "products");

        private const string LastProductFileName = "_last_product.json";
        private const string ProductJsonFileName = "product.json";
        private const string ImagesSubfolder = "images";
        private const string FileExtension = ".json";

        public ProductStorageService()
        {
            Directory.CreateDirectory(ProductsDirectory);
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
        /// Сохранить продукт в базу с уникальным именем
        /// </summary>
        public string Save(Product product, string customName = null)
        {
            if (product == null) return null;

            // Проверяем, есть ли уже сохранённый продукт
            string existingFolderPath = FindExistingProductFolder(product);
            
            if (existingFolderPath != null)
            {
                // Перезаписываем существующий продукт
                string existingFilePath = Path.Combine(existingFolderPath, ProductJsonFileName);
                SaveToFile(product, existingFilePath);
                return existingFilePath;
            }

            // Создаём новую папку для продукта
            string productFolderName = GenerateProductFolderName(product, customName);
            string productFolderPath = Path.Combine(ProductsDirectory, productFolderName);
            Directory.CreateDirectory(productFolderPath);
            
            // Создаём папку для изображений
            string imagesFolder = Path.Combine(productFolderPath, ImagesSubfolder);
            Directory.CreateDirectory(imagesFolder);

            // Сохраняем JSON
            string filePath = Path.Combine(productFolderPath, ProductJsonFileName);
            SaveToFile(product, filePath);
            return filePath;
        }

        /// <summary>
        /// Найти существующую папку продукта по имени и обозначению
        /// </summary>
        private string FindExistingProductFolder(Product product)
        {
            if (!Directory.Exists(ProductsDirectory))
                return null;

            var folders = Directory.GetDirectories(ProductsDirectory)
                .Where(f => !Path.GetFileName(f).StartsWith("_"));

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

            return null;
        }

        /// <summary>
        /// Загрузить продукт по имени папки
        /// </summary>
        public Product Load(string folderName)
        {
            string folderPath = Path.Combine(ProductsDirectory, folderName);
            string filePath = Path.Combine(folderPath, ProductJsonFileName);
            return LoadFromFile(filePath, folderPath);
        }

        /// <summary>
        /// Получить список всех сохранённых продуктов
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
        /// Удалить продукт из базы
        /// </summary>
        public bool Delete(string folderName)
        {
            try
            {
                string folderPath = Path.Combine(ProductsDirectory, folderName);
                if (Directory.Exists(folderPath))
                {
                    Directory.Delete(folderPath, true);
                    return true;
                }
            }
            catch
            {
                // Ошибка удаления
            }
            return false;
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

        private string GenerateProductFolderName(Product product, string customName)
        {
            string baseName = customName ?? $"{product.Name}_{product.Marking}";
            // Убираем недопустимые символы
            baseName = string.Join("_", baseName.Split(Path.GetInvalidFileNameChars()));

            string folderPath = Path.Combine(ProductsDirectory, baseName);

            // Если папка существует, добавляем номер
            int counter = 1;
            while (Directory.Exists(folderPath))
            {
                string folderName = $"{baseName}_{counter}";
                folderPath = Path.Combine(ProductsDirectory, folderName);
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

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    serializer.WriteObject(stream, dto);
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
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var dto = (ProductDto)serializer.ReadObject(stream);
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