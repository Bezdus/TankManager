using Kompas6Constants;
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
    /// <summary>
    /// Извлекает детали из документа KOMPAS
    /// </summary>
    internal class PartExtractor
    {
        private readonly KompasContext _context;
        private readonly ILogger _logger;
        private readonly ComObjectManager _comManager;

        public PartExtractor(KompasContext context, ILogger logger, ComObjectManager comManager)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _comManager = comManager ?? throw new ArgumentNullException(nameof(comManager));
        }

        public void ExtractParts(IPart7 part, List<PartModel> details)
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
                    ProcessBodiesArray(bodiesArray, details);
                }
                else if (bodiesVariant is IBody7 singleBody)
                {
                    ProcessSingleBody(singleBody, details);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error extracting bodies: {ex.Message}");
            }
            finally
            {
                CleanupComObjects(bodiesVariant, feature, part);
            }
        }

        private void ProcessBodiesArray(Array bodiesArray, List<PartModel> details)
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
                        int instanceIndex = CountMatchingParts(details, body.Name, body.Marking, isBodyBased: true);
                        details.Add(new PartModel(body, _context, instanceIndex));
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

        private void ProcessSingleBody(IBody7 body, List<PartModel> details)
        {
            if (IsDetailBody(body))
            {
                _comManager.Track(body);
                int instanceIndex = CountMatchingParts(details, body.Name, body.Marking, isBodyBased: true);
                details.Add(new PartModel(body, _context, instanceIndex));
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
                        ProcessSubPart(subPart, details);
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

        private void ProcessSubPart(IPart7 subPart, List<PartModel> details)
        {
            if (subPart.Detail == true)
            {
                AddPartToList(subPart, details);
            }
            else
            {
                var detailType = _context.GetDetailType(subPart);
                if (detailType == KompasConstants.OtherPartsType)
                {
                    AddPartToList(subPart, details);
                }
                else
                {
                    ExtractParts(subPart, details);
                }
            }
        }

        private void AddPartToList(IPart7 subPart, List<PartModel> details)
        {
            int instanceIndex = CountMatchingParts(details, subPart.Name, subPart.Marking,
                isBodyBased: false, filePath: subPart.FileName);
            details.Add(new PartModel(subPart, _context, instanceIndex));
        }

        private int CountMatchingParts(List<PartModel> details, string name, string marking,
            bool isBodyBased, string filePath = null)
        {
            return details.Count(d =>
                d.Name == name &&
                d.Marking == marking &&
                d.IsBodyBased == isBodyBased &&
                (filePath == null || d.FilePath == filePath));
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

        private void CleanupComObjects(object bodiesVariant, IFeature7 feature, IPart7 part)
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
}