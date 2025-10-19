using Kompas6Constants3D;
using KompasAPI7;
using System;
using System.Collections.Generic;
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
                                details.Add(new PartModel(body, _context));
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
                    
                    // Освобождаем массив
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
                        details.Add(new PartModel(singleBody, _context));
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
                            details.Add(new PartModel(subPart, _context));
                        }
                        else
                        {
                            var detailType = _context.GetDetailType(subPart);
                            if (detailType == KompasConstants.OtherPartsType)
                            {
                                details.Add(new PartModel(subPart, _context));
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
            if (!IsValidDetailForDisplay(detail))
            {
                _logger.LogWarning("Invalid detail for display");
                return;
            }

            try
            {
                SelectDetail(detail);
                FocusOnDetail(detail);
                _logger.LogInfo($"Focused on detail: {detail.Name}");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to show detail in Kompas: {detail.Name}", ex);
                throw;
            }
        }

        private bool IsValidDetailForDisplay(PartModel detail)
        {
            return _context.IsDocumentLoaded && detail?.Part != null;
        }

        private void SelectDetail(PartModel detail)
        {
            _context.SelectionManager.UnselectAll();
            _context.SelectionManager.Select(detail.Part);
        }

        private void FocusOnDetail(PartModel detail)
        {
            var gabarit = GetPartGabarit(detail.Part);
            var center = CalculateCenter(gabarit);
            var scale = CalculateScale(gabarit);
            UpdateCamera(center, scale);
        }

        private struct Gabarit
        {
            public double X1, Y1, Z1, X2, Y2, Z2;
        }

        private Gabarit GetPartGabarit(IPart7 part)
        {
            var gabarit = new Gabarit();
            part.GetGabarit(false, false,
                out gabarit.X1, out gabarit.Y1, out gabarit.Z1,
                out gabarit.X2, out gabarit.Y2, out gabarit.Z2);
            return gabarit;
        }

        private (double x, double y, double z) CalculateCenter(Gabarit g)
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

        public void Dispose()
        {
            _logger.LogInfo("Disposing KompasService");
            _comManager?.Dispose();
            _context?.Dispose();
        }
    }
}
