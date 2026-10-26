using System.Collections.Generic;
using System.Linq;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Агрегирует материалы из деталей изделия по типам
    /// </summary>
    public class MaterialAggregator
    {
        private readonly ILogger _logger;

        public MaterialAggregator(ILogger logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Агрегирует материалы из деталей в соответствующие коллекции продукта
        /// </summary>
        /// <param name="product">Продукт для агрегации материалов</param>
        public void AggregateMaterials(Product product)
        {
            if (product == null)
                return;

            product.SheetMaterials.Clear();
            product.TubularProducts.Clear();
            product.OtherMaterials.Clear();

            AggregateSheetMaterials(product);
            AggregateTubularProducts(product);
            AggregateOtherMaterials(product);

            _logger.LogInfo($"Aggregated materials: {product.SheetMaterials.Count} sheet, {product.TubularProducts.Count} tubular, {product.OtherMaterials.Count} other");
        }

        /// <summary>
        /// Детали, материал которых учитывается в сводках: покупные изделия не учитываются
        /// </summary>
        private static IEnumerable<PartModel> GetMaterialParts(Product product)
        {
            return product.Details
                .Where(p => p.ProductType != ProductType.PurchasedPart && !string.IsNullOrEmpty(p.Material));
        }

        /// <summary>
        /// Группирует листовой прокат по материалу
        /// </summary>
        private void AggregateSheetMaterials(Product product)
        {
            var sheetMaterials = GetMaterialParts(product)
                .Where(p => p.ProductType == ProductType.SheetMaterial)
                .GroupBy(p => p.Material)
                .Select(g => new MaterialInfo
                {
                    Name = g.Key,
                    TotalMass = g.Sum(p => p.Mass)
                })
                .OrderByDescending(m => m.TotalMass);

            foreach (var material in sheetMaterials)
            {
                product.SheetMaterials.Add(material);
            }
        }

        /// <summary>
        /// Группирует трубный прокат по материалу
        /// </summary>
        private void AggregateTubularProducts(Product product)
        {
            var tubularProducts = GetMaterialParts(product)
                .Where(p => p.ProductType == ProductType.TubularProduct)
                .GroupBy(p => p.Material)
                .Select(g => new MaterialInfo
                {
                    Name = g.Key,
                    TotalMass = g.Sum(p => p.Mass),
                    TotalLength = g.Sum(p => p.Length)
                })
                .OrderByDescending(m => m.TotalLength);

            foreach (var material in tubularProducts)
            {
                product.TubularProducts.Add(material);
            }
        }

        /// <summary>
        /// Группирует прочие материалы по материалу
        /// </summary>
        private void AggregateOtherMaterials(Product product)
        {
            var otherMaterials = GetMaterialParts(product)
                .Where(p => p.ProductType != ProductType.SheetMaterial
                         && p.ProductType != ProductType.TubularProduct)
                .GroupBy(p => p.Material)
                .Select(g => new MaterialInfo
                {
                    Name = g.Key,
                    TotalMass = g.Sum(p => p.Mass)
                })
                .OrderByDescending(m => m.TotalMass);

            foreach (var material in otherMaterials)
            {
                product.OtherMaterials.Add(material);
            }
        }
    }
}
