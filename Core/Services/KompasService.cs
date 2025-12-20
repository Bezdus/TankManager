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
    /// <summary>
    /// Сервис для работы с KOMPAS-3D
    /// </summary>
    public class KompasService : IKompasService
    {
        private readonly ILogger _logger;
        private readonly ComObjectManager _comManager;
        private readonly MaterialAggregator _materialAggregator;
        private readonly DrawingPreviewService _previewService;

        public KompasService() : this(new FileLogger())
        {
        }

        public KompasService(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _comManager = new ComObjectManager(_logger);
            _materialAggregator = new MaterialAggregator(_logger);
            _previewService = new DrawingPreviewService();
        }

        /// <summary>
        /// Загружает документ из файла
        /// </summary>
        /// <param name="filePath">Путь к файлу документа</param>
        /// <returns>Загруженный продукт</returns>
        public Product LoadDocument(string filePath)
        {
            _logger.LogInfo($"Loading document: {filePath}");

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

        /// <summary>
        /// Загружает активный документ из KOMPAS
        /// </summary>
        /// <returns>Загруженный продукт</returns>
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

        /// <summary>
        /// Создает продукт из верхней детали контекста
        /// </summary>
        private Product CreateProductFromTopPart(KompasContext context)
        {
            var topPart = context.TopPart;
            var product = new Product(topPart, context);

            ExtractAllParts(product, topPart);
            _materialAggregator.AggregateMaterials(product);

            return product;
        }

        /// <summary>
        /// Извлекает все детали из верхней детали
        /// </summary>
        private void ExtractAllParts(Product product, IPart7 topPart)
        {
            var partExtractor = new PartExtractor(context: product.Context, logger: _logger, comManager: _comManager);
            var parts = new List<PartModel>();
            partExtractor.ExtractParts(topPart, parts);

            foreach (var part in parts)
            {
                product.Details.Add(part);

                if (part.ProductType == ProductType.PurchasedPart)
                {
                    product.StandardParts.Add(part);
                }
            }
        }

        /// <summary>
        /// Отображает деталь в KOMPAS (выделяет и центрирует камеру)
        /// </summary>
        /// <param name="detail">Деталь для отображения</param>
        /// <param name="product">Продукт, содержащий деталь</param>
        public void ShowDetailInKompas(PartModel detail, Product product)
        {
            if (!ValidateShowDetailParameters(detail, product))
                return;

            var context = product.Context;

            try
            {
                var targetPart = FindTargetPart(detail, context);
                if (targetPart == null)
                {
                    _logger.LogWarning($"Could not find part: {detail.Name}");
                    return;
                }

                ActivateDocumentAndShowPart(context, targetPart);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to show detail in Kompas: {detail.Name}", ex);
                throw;
            }
        }

        /// <summary>
        /// Валидирует параметры для отображения детали
        /// </summary>
        private bool ValidateShowDetailParameters(PartModel detail, Product product)
        {
            if (detail == null)
            {
                _logger.LogWarning("Detail is null");
                return false;
            }

            if (product?.Context == null || !product.Context.IsDocumentLoaded)
            {
                _logger.LogWarning("Product context not loaded");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Находит деталь в контексте по модели
        /// </summary>
        private IPart7 FindTargetPart(PartModel detail, KompasContext context)
        {
            Debug.WriteLine(context.TopPart.Name);
            
            var partFinder = new PartFinder(_logger);
            return partFinder.FindPartByModel(detail, context.TopPart);
        }

        /// <summary>
        /// Активирует документ, выделяет и центрирует камеру на детали
        /// </summary>
        private void ActivateDocumentAndShowPart(KompasContext context, IPart7 targetPart)
        {
            context.Document.Active = true;
            SelectDetail(context, targetPart);
            
            var cameraController = new KompasCameraController(context, _logger);
            cameraController.FocusOnPart(targetPart);
        }

        /// <summary>
        /// Проверяет инициализацию KOMPAS
        /// </summary>
        private void EnsureKompasInitialized(KompasContext context)
        {
            if (!context.IsInitialized)
            {
                const string errorMessage = "Не удалось подключиться к KOMPAS-3D";
                _logger.LogError(errorMessage);
                throw new InvalidOperationException(errorMessage);
            }
        }

        /// <summary>
        /// Выделяет деталь в документе
        /// </summary>
        private void SelectDetail(KompasContext context, IPart7 part)
        {
            context.SelectionManager.UnselectAll();
            context.SelectionManager.Select(part);
        }

        /// <summary>
        /// Загружает превью чертежа для детали
        /// </summary>
        /// <param name="detail">Деталь для загрузки превью</param>
        /// <param name="product">Продукт, содержащий деталь</param>
        public void LoadDrawingPreview(PartModel detail, Product product)
        {
            if (detail == null || product?.Context == null || !product.Context.IsDocumentLoaded)
                return;

            // Пропускаем Body-based детали (у них нет своего чертежа)
            if (detail.IsBodyBased)
                return;

            try
            {
                var targetPart = FindTargetPart(detail, product.Context);
                if (targetPart != null)
                {
                    detail.LoadDrawingPreview(targetPart, product.Context);
                    if (!string.IsNullOrEmpty(detail.CdfFilePath))
                    {
                        _logger.LogInfo($"Drawing preview loaded for: {detail.Name}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Failed to load drawing preview for {detail.Name}: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _logger.LogInfo("Disposing KompasService");
            _comManager?.Dispose();
        }
    }
}
