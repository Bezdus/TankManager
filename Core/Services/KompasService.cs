using Kompas6API5;
using Kompas6Constants3D;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    public class KompasService : IKompasService
    {
        private readonly ILogger _logger;
        private readonly ComObjectManager _comManager;

        public KompasService() : this(new FileLogger())
        {
        }

        public KompasService(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _comManager = new ComObjectManager(_logger);
        }

        public Product LoadDocument(string filePath)
        {
            _logger.LogInfo($"Loading document: {filePath}");

            // Создаём новый контекст для каждого документа
            var context = new KompasContext();
            EnsureKompasInitialized(context);

            try
            {
                context.LoadDocument(filePath);

                if (context.IsDocumentLoaded)
                {
                    var product = CreateProductFromTopPart(context);
                    _logger.LogInfo($"Successfully loaded product '{product.Name}' with {product.Details.Count} parts");
                    return product;
                }
                else
                {
                    _logger.LogWarning("Document loaded but TopPart is null");
                    return new Product();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to load document: {filePath}", ex);
                context.Dispose();
                throw;
            }
        }

        public Product LoadActiveDocument()
        {
            _logger.LogInfo("Loading active document from KOMPAS");

            var context = new KompasContext();
            EnsureKompasInitialized(context);

            try
            {
                context.LoadActiveDocument();

                if (context.IsDocumentLoaded)
                {
                    var product = CreateProductFromTopPart(context);
                    _logger.LogInfo($"Successfully loaded product '{product.Name}' with {product.Details.Count} parts from active document");
                    return product;
                }
                else
                {
                    _logger.LogWarning("Active document loaded but TopPart is null");
                    return new Product();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Failed to load active document", ex);
                context.Dispose();
                throw;
            }
        }

        private Product CreateProductFromTopPart(KompasContext context)
        {
            var topPart = context.TopPart;
            var product = new Product(topPart, context);

            var partExtractor = new PartExtractor(context, _logger, _comManager);
            var parts = new List<PartModel>();
            partExtractor.ExtractParts(topPart, parts);

            foreach (var part in parts)
            {
                product.Details.Add(part);

                switch (part.ProductType)
                {
                    case ProductType.PurchasedPart:
                        product.StandardParts.Add(part);
                        break;
                }
            }

            // Агрегируем материалы по типам
            AggregateMaterials(product);

            return product;
        }

        /// <summary>
        /// Агрегирует материалы из деталей в соответствующие коллекции
        /// </summary>
        private void AggregateMaterials(Product product)
        {
            // Группируем листовой прокат по материалу
            var sheetMaterials = product.Details
                .Where(p => p.ProductType == ProductType.SheetMaterial && !string.IsNullOrEmpty(p.Material))
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

            // Группируем трубный прокат по материалу
            var tubularProducts = product.Details
                .Where(p => p.ProductType == ProductType.TubularProduct && !string.IsNullOrEmpty(p.Material))
                .GroupBy(p => p.Material)
                .Select(g => new MaterialInfo
                {
                    Name = g.Key,
                    TotalMass = g.Sum(p => p.Mass)
                })
                .OrderByDescending(m => m.TotalMass);

            foreach (var material in tubularProducts)
            {
                product.TubularProducts.Add(material);
            }

            _logger.LogInfo($"Aggregated materials: {product.SheetMaterials.Count} sheet, {product.TubularProducts.Count} tubular");
        }

        public void ShowDetailInKompas(PartModel detail, Product product)
        {
            if (detail == null)
            {
                _logger.LogWarning("Detail is null");
                return;
            }

            if (product?.Context == null || !product.Context.IsDocumentLoaded)
            {
                _logger.LogWarning("Product context not loaded");
                return;
            }

            var context = product.Context;

            try
            {
                Debug.WriteLine(context.TopPart.Name);
                
                var partFinder = new PartFinder(_logger);
                IPart7 targetPart = partFinder.FindPartByModel(detail, context.TopPart);
                
                if (targetPart == null)
                {
                    _logger.LogWarning($"Could not find part: {detail.Name}");
                    return;
                }

                context.Document.Active = true;

                SelectDetail(context, targetPart);
                
                var cameraController = new KompasCameraController(context, _logger);
                cameraController.FocusOnPart(targetPart);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to show detail in Kompas: {detail.Name}", ex);
                throw;
            }
        }

        private void EnsureKompasInitialized(KompasContext context)
        {
            if (!context.IsInitialized)
            {
                const string errorMessage = "Не удалось подключиться к KOMPAS-3D";
                _logger.LogError(errorMessage);
                throw new InvalidOperationException(errorMessage);
            }
        }

        private void SelectDetail(KompasContext context, IPart7 part)
        {
            context.SelectionManager.UnselectAll();
            context.SelectionManager.Select(part);
        }

        public void Dispose()
        {
            _logger.LogInfo("Disposing KompasService");
            _comManager?.Dispose();
        }
    }
}
