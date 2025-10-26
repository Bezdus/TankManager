using KompasAPI7;
using System;
using System.Collections.Generic;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Отвечает за поиск деталей в иерархии KOMPAS
    /// </summary>
    internal class PartFinder
    {
        private readonly ILogger _logger;

        public PartFinder(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public IPart7 FindPartByModel(PartModel model, IPart7 rootPart)
        {
            if (rootPart == null || model == null)
                return null;

            int foundIndex = 0;
            return FindPartByModelRecursive(model, rootPart, ref foundIndex);
        }

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
                    var foundPart = SearchInBodies(model, currentPart, ref foundIndex);
                    if (foundPart != null)
                        return foundPart;
                }

                return SearchInSubParts(model, currentPart, ref foundIndex);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error finding part: {ex.Message}");
                return null;
            }
        }

        private IPart7 SearchInBodies(PartModel model, IPart7 currentPart, ref int foundIndex)
        {
            if (!(currentPart is IFeature7 feature))
                return null;

            object bodiesVariant = feature.ResultBodies;
            if (!(bodiesVariant is Array bodiesArray))
                return null;

            foreach (var bodyObj in bodiesArray)
            {
                if (bodyObj is IBody7 body && MatchesBody(model, body))
                {
                    if (foundIndex == model.InstanceIndex)
                        return currentPart;
                    foundIndex++;
                }
            }

            return null;
        }

        private IPart7 SearchInSubParts(PartModel model, IPart7 currentPart, ref int foundIndex)
        {
            IParts7 partsCollection = currentPart.Parts;
            if (partsCollection == null)
                return null;

            foreach (IPart7 subPart in partsCollection)
            {
                IPart7 found = FindPartByModelRecursive(model, subPart, ref foundIndex);
                if (found != null)
                    return found;
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
    }
}