using Kompas6API5;
using Kompas6Constants3D;
using KompasAPI7;
using System;
using System.Collections.Generic;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    public class KompasService : IKompasService
    {
        private readonly KompasContext _context;
        private readonly ILogger _logger;
        private readonly ComObjectManager _comManager;
        private readonly PartExtractor _partExtractor;
        private readonly PartFinder _partFinder;
        private readonly KompasCameraController _cameraController;
        private readonly PartIntersectionDetector _intersectionDetector;

        public KompasService() : this(new KompasContext(), new FileLogger())
        {
        }

        public KompasService(KompasContext context, ILogger logger)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _comManager = new ComObjectManager(_logger);
            _partExtractor = new PartExtractor(_context, _logger, _comManager);
            _partFinder = new PartFinder(_logger);
            _cameraController = new KompasCameraController(_context, _logger);
            _intersectionDetector = new PartIntersectionDetector(_context, _logger, _cameraController);
        }

        public List<PartModel> LoadDocument(string filePath)
        {
            _logger.LogInfo($"Loading document: {filePath}");

            EnsureKompasInitialized();

            var result = new List<PartModel>();

            try
            {
                _context.LoadDocument(filePath);

                if (_context.IsDocumentLoaded)
                {
                    _partExtractor.ExtractParts(_context.TopPart, result);
                    _logger.LogInfo($"Successfully loaded {result.Count} parts");
                }
                else
                {
                    _logger.LogWarning("Document loaded but TopPart is null");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to load document: {filePath}", ex);
                throw;
            }

            return result;
        }

        public List<PartModel> LoadActiveDocument()
        {
            _logger.LogInfo("Loading active document from KOMPAS");

            EnsureKompasInitialized();

            var result = new List<PartModel>();

            try
            {
                _context.LoadActiveDocument();

                if (_context.IsDocumentLoaded)
                {
                    _partExtractor.ExtractParts(_context.TopPart, result);
                    _logger.LogInfo($"Successfully loaded {result.Count} parts from active document");
                }
                else
                {
                    _logger.LogWarning("Active document loaded but TopPart is null");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Failed to load active document", ex);
                throw;
            }

            return result;
        }

        public void ShowDetailInKompas(PartModel detail)
        {
            if (detail == null || !_context.IsDocumentLoaded)
            {
                _logger.LogWarning("Invalid detail or document not loaded");
                return;
            }

            try
            {
                IPart7 targetPart = _partFinder.FindPartByModel(detail, _context.TopPart);
                if (targetPart == null)
                {
                    _logger.LogWarning($"Could not find part: {detail.Name}");
                    return;
                }

                SelectDetail(targetPart);
                _cameraController.FocusOnPart(targetPart);
                
                // Находим пересекающиеся объекты (если необходимо)
                //List<IPart7> intersectingObjects = _intersectionDetector.FindIntersectingObjects(targetPart);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to show detail in Kompas: {detail.Name}", ex);
                throw;
            }
        }

        private void EnsureKompasInitialized()
        {
            if (!_context.IsInitialized)
            {
                const string errorMessage = "Не удалось подключиться к KOMPAS-3D";
                _logger.LogError(errorMessage);
                throw new InvalidOperationException(errorMessage);
            }
        }

        private void SelectDetail(IPart7 part)
        {
            _context.SelectionManager.UnselectAll();
            _context.SelectionManager.Select(part);
        }

        public void Dispose()
        {
            _logger.LogInfo("Disposing KompasService");
            _comManager?.Dispose();
            _context?.Dispose();
        }
    }
}
