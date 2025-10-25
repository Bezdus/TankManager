using Kompas6Constants3D;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using TankManager.Core.Constants;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    public class KompasService : IKompasService
    {
        private readonly KompasContext _context;
        private readonly ILogger _logger;
        private readonly ComObjectManager _comManager;

        public KompasService() : this(new KompasContext(), new FileLogger())
        {
        }

        public KompasService(KompasContext context, ILogger logger)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _comManager = new ComObjectManager(_logger);
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
                    ExtractParts(_context.TopPart, result);
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

        private void EnsureKompasInitialized()
        {
            if (!_context.IsInitialized)
            {
                const string errorMessage = "Не удалось подключиться к KOMPAS-3D";
                _logger.LogError(errorMessage);
                throw new InvalidOperationException(errorMessage);
            }
        }

        private void ExtractParts(IPart7 part, List<PartModel> details)
        {
            if (part == null) return;

            try
            {
                ExtractBodies(part, details);
                ExtractSubParts(part, details);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error extracting parts from {part.Name}", ex);
            }
        }

        private void ExtractBodies(IPart7 part, List<PartModel> details)
        {
            IFeature7 feature = null;
            object bodiesVariant = null;
            
            try
            {
                feature = part as IFeature7;
                if (feature == null)
                    return;

                _comManager.Track(feature);
                bodiesVariant = feature.ResultBodies;
                
                if (bodiesVariant == null)
                    return;

                int bodyIndex = 0; // Счётчик для тел
                if (bodiesVariant is Array bodiesArray)
                {
                    foreach (var bodyObj in bodiesArray)
                    {
                        IBody7 body = null;
                        try
                        {
                            body = bodyObj as IBody7;
                            if (body != null && IsDetailBody(body))
                            {
                                _comManager.Track(body);
                        
                                // Подсчитываем, сколько таких же тел уже добавлено
                                int instanceIndex = details.Count(d => 
                                    d.Name == body.Name && 
                                    d.Marking == body.Marking && 
                                    d.IsBodyBased);
                        
                                details.Add(new PartModel(body, _context, instanceIndex));
                                bodyIndex++;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning($"Error processing body: {ex.Message}");
                            if (body != null)
                            {
                                _comManager.Release(body);
                            }
                        }
                    }
            
                    if (Marshal.IsComObject(bodiesArray))
                    {
                        Marshal.ReleaseComObject(bodiesArray);
                    }
                }
                else if (bodiesVariant is IBody7 singleBody)
                {
                    if (IsDetailBody(singleBody))
                    {
                        _comManager.Track(singleBody);
                
                        int instanceIndex = details.Count(d => 
                            d.Name == singleBody.Name && 
                            d.Marking == singleBody.Marking && 
                            d.IsBodyBased);
                
                        details.Add(new PartModel(singleBody, _context, instanceIndex));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error extracting bodies: {ex.Message}");
            }
            finally
            {
                if (bodiesVariant != null && Marshal.IsComObject(bodiesVariant))
                {
                    Marshal.ReleaseComObject(bodiesVariant);
                }
                
                if (feature != null && feature != part)
                {
                    _comManager.Release(feature);
                }
            }
        }

        private bool IsDetailBody(IBody7 body)
        {
            IPropertyKeeper propertyKeeper = null;
            try
            {
                propertyKeeper = body as IPropertyKeeper;
                if (propertyKeeper == null)
                    return false;

                propertyKeeper.GetPropertyValue(
                    (_Property)_context.SpecificationSectionProperty,
                    out object val,
                    false,
                    out _);

                return (val as string) == KompasConstants.DetailsSectionName;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (propertyKeeper != null && propertyKeeper != (object)body)
                {
                    Marshal.ReleaseComObject(propertyKeeper);
                }
            }
        }

        private void ExtractSubParts(IPart7 part, List<PartModel> details)
        {
            IParts7 partsCollection = null;

            try
            {
                partsCollection = part.Parts;
                _comManager.Track(partsCollection);

                if (partsCollection == null)
                    return;

                foreach (IPart7 subPart in partsCollection)
                {
                    if (subPart == null) continue;

                    _comManager.Track(subPart);

                    try
                    {
                        if (subPart.Detail == true)
                        {
                            // Подсчитываем, сколько таких же деталей уже добавлено
                            int instanceIndex = details.Count(d => 
                                d.Name == subPart.Name && 
                                d.Marking == subPart.Marking && 
                                !d.IsBodyBased &&
                                d.FilePath == subPart.FileName);
                    
                            details.Add(new PartModel(subPart, _context, instanceIndex));
                        }
                        else
                        {
                            var detailType = _context.GetDetailType(subPart);
                            if (detailType == KompasConstants.OtherPartsType)
                            {
                                int instanceIndex = details.Count(d => 
                                    d.Name == subPart.Name && 
                                    d.Marking == subPart.Marking && 
                                    !d.IsBodyBased &&
                                    d.FilePath == subPart.FileName);
                        
                                details.Add(new PartModel(subPart, _context, instanceIndex));
                            }
                            else
                            {
                                ExtractParts(subPart, details);
                            }
                        }
                    }
                    catch (COMException ex)
                    {
                        _logger.LogWarning($"COM error accessing part: {ex.Message}");
                    }
                    finally
                    {
                        _comManager.Release(subPart);
                    }
                }
            }
            finally
            {
                if (partsCollection != null)
                {
                    _comManager.Release(partsCollection);
                }
            }
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
                IPart7 targetPart = FindPartByModel(detail, _context.TopPart);
                if (targetPart == null)
                {
                    _logger.LogWarning($"Could not find part: {detail.Name}");
                    return;
                }

                SelectDetail(targetPart);
                FocusOnDetail(targetPart);
                _logger.LogInfo($"Focused on detail: {detail.Name}");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to show detail in Kompas: {detail.Name}", ex);
                throw;
            }
        }

        private IPart7 FindPartByModel(PartModel model, IPart7 currentPart)
        {
            if (currentPart == null)
                return null;

            // Счётчик найденных совпадений
            int foundIndex = 0;

            try
            {
                if (MatchesPart(model, currentPart))
                {
                    if (foundIndex == model.InstanceIndex)
                        return currentPart;
                    foundIndex++;
                }

                if (model.IsBodyBased)
                {
                    IFeature7 feature = currentPart as IFeature7;
                    if (feature != null)
                    {
                        object bodiesVariant = feature.ResultBodies;
                        if (bodiesVariant is Array bodiesArray)
                        {
                            foreach (var bodyObj in bodiesArray)
                            {
                                IBody7 body = bodyObj as IBody7;
                                if (body != null && MatchesBody(model, body))
                                {
                                    if (foundIndex == model.InstanceIndex)
                                        return currentPart;
                                    foundIndex++;
                                }
                            }
                        }
                    }
                }

                IParts7 partsCollection = currentPart.Parts;
                if (partsCollection != null)
                {
                    foreach (IPart7 subPart in partsCollection)
                    {
                        IPart7 found = FindPartByModelRecursive(model, subPart, ref foundIndex);
                        if (found != null)
                            return found;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error finding part: {ex.Message}");
            }

            return null;
        }

        private bool MatchesPart(PartModel model, IPart7 part)
        {
            return part.Name == model.Name && 
                   part.Marking == model.Marking &&
                   (part.FileName == model.FilePath || string.IsNullOrEmpty(model.FilePath));
        }

        private bool MatchesBody(PartModel model, IBody7 body)
        {
            return body.Name == model.Name && body.Marking == model.Marking;
        }

        private void SelectDetail(IPart7 part)
        {
            _context.SelectionManager.UnselectAll();
            _context.SelectionManager.Select(part);
        }

        private void FocusOnDetail(IPart7 part)
        {
            var gabarit = GetPartGabarit(part);
            var center = CalculateGlobalCenter(part);
            var scale = CalculateScale(gabarit);
            UpdateCamera(center, scale);
        }

        private struct Gabarit
        {
            public double X1, Y1, Z1, X2, Y2, Z2;
        }

        private (double x, double y, double z) CalculateGlobalCenter(IPart7 detail)
        {
            // Начинаем с локального центра детали
            var gabarit = GetPartGabarit(detail);
            var localCenter = CalculateLocalCenter(gabarit);

            double globalX = localCenter.x;
            double globalY = localCenter.y;
            double globalZ = localCenter.z;

            // Начинаем с родителя детали (пропускаем саму деталь)
            IPart7 currentPart = TryGetParent(detail);

            // Проходим по всей иерархии родителей до TopPart
            while (currentPart != null)
            {
                IPart7 parentPart = TryGetParent(currentPart);

                // Если есть родитель, применяем смещение текущей части относительно родителя
                if (parentPart != null)
                {
                    double originX, originY, originZ;
                    currentPart.Placement.GetOrigin(out originX, out originY, out originZ);

                    globalX += originX;
                    globalY += originY;
                    globalZ += originZ;
                }

                currentPart = parentPart;
            }

            return (globalX, globalY, globalZ);
        }

        private IPart7 TryGetParent(IPart7 part)
        {
            try
            {
                return part.Parent as IPart7;
            }
            catch
            {
                return null;
            }
        }

        private Gabarit GetPartGabarit(IPart7 part)
        {
            var gabarit = new Gabarit();
            part.GetGabarit(true, true,
                out gabarit.X1, out gabarit.Y1, out gabarit.Z1,
                out gabarit.X2, out gabarit.Y2, out gabarit.Z2);
            return gabarit;
        }

        private (double x, double y, double z) CalculateLocalCenter(Gabarit g)
        {
            return ((g.X1 + g.X2) / 2.0, (g.Y1 + g.Y2) / 2.0, (g.Z1 + g.Z2) / 2.0);
        }

        private double CalculateScale(Gabarit g)
        {
            double maxSize = Math.Max(
                Math.Max(Math.Abs(g.X2 - g.X1), Math.Abs(g.Y2 - g.Y1)),
                Math.Abs(g.Z2 - g.Z1));

            return maxSize > 0 
                ? KompasConstants.ScaleFactor / maxSize 
                : KompasConstants.DefaultScale;
        }

        private void UpdateCamera((double x, double y, double z) center, double scale)
        {
            object matrixObj = null;
            try
            {
                matrixObj = _context.ViewProjectionManager.Matrix3D;
                
                if (matrixObj is Array matrix)
                {
                    matrix.SetValue(center.x, KompasConstants.CameraMatrixXIndex);
                    matrix.SetValue(center.y, KompasConstants.CameraMatrixYIndex);
                    matrix.SetValue(center.z, KompasConstants.CameraMatrixZIndex);
                    _context.ViewProjectionManager.SetMatrix3D(matrix, scale);
                }
            }
            finally
            {
                if (matrixObj != null && Marshal.IsComObject(matrixObj))
                {
                    Marshal.ReleaseComObject(matrixObj);
                }
            }
        }

        // Вспомогательный рекурсивный метод с передачей счётчика
        private IPart7 FindPartByModelRecursive(PartModel model, IPart7 currentPart, ref int foundIndex)
        {
            if (currentPart == null)
                return null;

            try
            {
                if (MatchesPart(model, currentPart))
                {
                    if (foundIndex == model.InstanceIndex)
                        return currentPart;
                    foundIndex++;
                }

                if (model.IsBodyBased)
                {
                    IFeature7 feature = currentPart as IFeature7;
                    if (feature != null)
                    {
                        object bodiesVariant = feature.ResultBodies;
                        if (bodiesVariant is Array bodiesArray)
                        {
                            foreach (var bodyObj in bodiesArray)
                            {
                                IBody7 body = bodyObj as IBody7;
                                if (body != null && MatchesBody(model, body))
                                {
                                    if (foundIndex == model.InstanceIndex)
                                        return currentPart;
                                    foundIndex++;
                                }
                            }
                        }
                    }
                }

                IParts7 partsCollection = currentPart.Parts;
                if (partsCollection != null)
                {
                    foreach (IPart7 subPart in partsCollection)
                    {
                        IPart7 found = FindPartByModelRecursive(model, subPart, ref foundIndex);
                        if (found != null)
                            return found;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error finding part: {ex.Message}");
            }

            return null;
        }

        public void Dispose()
        {
            _logger.LogInfo("Disposing KompasService");
            _comManager?.Dispose();
            _context?.Dispose();
        }
    }
}
