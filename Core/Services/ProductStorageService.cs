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
        private const string FileExtension = ".json";

        public ProductStorageService()
        {
            Directory.CreateDirectory(ProductsDirectory);
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
            return LoadFromFile(Path.Combine(ProductsDirectory, LastProductFileName));
        }

        /// <summary>
        /// Сохранить продукт в базу с уникальным именем
        /// </summary>
        public string Save(Product product, string customName = null)
        {
            if (product == null) return null;

            string fileName = GenerateFileName(product, customName);
            string filePath = Path.Combine(ProductsDirectory, fileName);

            SaveToFile(product, filePath);
            return filePath;
        }

        /// <summary>
        /// Загрузить продукт по имени файла
        /// </summary>
        public Product Load(string fileName)
        {
            string filePath = Path.Combine(ProductsDirectory, fileName);
            return LoadFromFile(filePath);
        }

        /// <summary>
        /// Получить список всех сохранённых продуктов
        /// </summary>
        public List<ProductFileInfo> GetSavedProducts()
        {
            var result = new List<ProductFileInfo>();

            if (!Directory.Exists(ProductsDirectory))
                return result;

            var files = Directory.GetFiles(ProductsDirectory, "*" + FileExtension)
                .Where(f => !Path.GetFileName(f).Equals(LastProductFileName, StringComparison.OrdinalIgnoreCase));

            foreach (var file in files)
            {
                try
                {
                    var fileInfo = new FileInfo(file);
                    var product = LoadFromFile(file);

                    if (product != null)
                    {
                        result.Add(new ProductFileInfo
                        {
                            FileName = Path.GetFileName(file),
                            ProductName = product.Name,
                            Marking = product.Marking,
                            DetailsCount = product.Details.Count,
                            SavedDate = fileInfo.LastWriteTime
                        });
                    }
                }
                catch
                {
                    // Пропускаем повреждённые файлы
                }
            }

            return result.OrderByDescending(p => p.SavedDate).ToList();
        }

        /// <summary>
        /// Удалить продукт из базы
        /// </summary>
        public bool Delete(string fileName)
        {
            try
            {
                string filePath = Path.Combine(ProductsDirectory, fileName);
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
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
        public bool Exists(string fileName)
        {
            return File.Exists(Path.Combine(ProductsDirectory, fileName));
        }

        #region Private Methods

        private string GenerateFileName(Product product, string customName)
        {
            string baseName = customName ?? $"{product.Name}_{product.Marking}";
            // Убираем недопустимые символы
            baseName = string.Join("_", baseName.Split(Path.GetInvalidFileNameChars()));

            string fileName = baseName + FileExtension;
            string filePath = Path.Combine(ProductsDirectory, fileName);

            // Если файл существует, добавляем номер
            int counter = 1;
            while (File.Exists(filePath))
            {
                fileName = $"{baseName}_{counter}{FileExtension}";
                filePath = Path.Combine(ProductsDirectory, fileName);
                counter++;
            }

            return fileName;
        }

        private void SaveToFile(Product product, string filePath)
        {
            try
            {
                var dto = ToDto(product);
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

        private Product LoadFromFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return null;

                var serializer = new DataContractJsonSerializer(typeof(ProductDto));
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var dto = (ProductDto)serializer.ReadObject(stream);
                    return FromDto(dto);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка загрузки Product: {ex.Message}");
                return null;
            }
        }

        private static ProductDto ToDto(Product product)
        {
            return new ProductDto
            {
                Name = product.Name,
                Marking = product.Marking,
                Mass = product.Mass,
                FilePath = product.FilePath,
                Details = product.Details.Select(d => ToPartDto(d)).ToList(),
                StandardParts = product.StandardParts.Select(d => ToPartDto(d)).ToList(),
                SheetMaterials = product.SheetMaterials.Select(m => ToMaterialDto(m)).ToList(),
                TubularProducts = product.TubularProducts.Select(m => ToMaterialDto(m)).ToList()
            };
        }

        private static PartModelDto ToPartDto(PartModel part)
        {
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
                ProductType = (int)part.ProductType
            };
        }

        private static MaterialInfoDto ToMaterialDto(MaterialInfo material)
        {
            return new MaterialInfoDto
            {
                Name = material.Name,
                TotalMass = material.TotalMass
            };
        }

        private static Product FromDto(ProductDto dto)
        {
            var product = new Product();
            product.Name = dto.Name;
            product.Marking = dto.Marking;
            product.Mass = dto.Mass;
            product.FilePath = dto.FilePath;

            foreach (var partDto in dto.Details ?? Enumerable.Empty<PartModelDto>())
            {
                product.Details.Add(FromPartDto(partDto));
            }

            foreach (var partDto in dto.StandardParts ?? Enumerable.Empty<PartModelDto>())
            {
                product.StandardParts.Add(FromPartDto(partDto));
            }

            foreach (var materialDto in dto.SheetMaterials ?? Enumerable.Empty<MaterialInfoDto>())
            {
                product.SheetMaterials.Add(FromMaterialDto(materialDto));
            }

            foreach (var materialDto in dto.TubularProducts ?? Enumerable.Empty<MaterialInfoDto>())
            {
                product.TubularProducts.Add(FromMaterialDto(materialDto));
            }

            return product;
        }

        private static PartModel FromPartDto(PartModelDto dto)
        {
            return new PartModelFromStorage
            {
                Name = dto.Name,
                Marking = dto.Marking,
                DetailType = dto.DetailType,
                Material = dto.Material,
                Mass = dto.Mass,
                FilePath = dto.FilePath,
                ProductType = (ProductType)(dto.ProductType)
            };
        }

        private static MaterialInfo FromMaterialDto(MaterialInfoDto dto)
        {
            return new MaterialInfo
            {
                Name = dto.Name,
                TotalMass = dto.TotalMass
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
    }

    [System.Runtime.Serialization.DataContract]
    public class MaterialInfoDto
    {
        [System.Runtime.Serialization.DataMember]
        public string Name { get; set; }

        [System.Runtime.Serialization.DataMember]
        public double TotalMass { get; set; }
    }

    #endregion
}