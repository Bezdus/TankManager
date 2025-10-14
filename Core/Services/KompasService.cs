using Kompas6Constants3D;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    public class KompasService : IKompasService
    {
        private const double DefaultScale = 1.0;
        private const double ScaleFactor = 100.0;
        private const string DetailsSection = "Детали";
        
        private readonly KompasContext _context;

        public KompasService()
        {
            _context = new KompasContext();
        }

        public KompasService(KompasContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public List<PartModel> LoadDocument(string filePath)
        {
            EnsureKompasInitialized();

            var result = new List<PartModel>();
            _context.LoadDocument(filePath);

            if (_context.IsDocumentLoaded)
            {
                ExtractParts(_context.TopPart, result);
            }

            return result;
        }

        private void EnsureKompasInitialized()
        {
            if (!_context.IsInitialized)
            {
                throw new InvalidOperationException("Не удалось подключиться к KOMPAS-3D");
            }
        }

        private void ExtractParts(IPart7 part, List<PartModel> details)
        {
            ExtractBodies(part, details);
            ExtractSubParts(part, details);
        }

        private void ExtractBodies(IPart7 part, List<PartModel> details)
        {
            var bodies = ((IFeature7)part).ResultBodies;
            
            if (bodies == null)
                return;
            foreach (IBody7 body in bodies)
            {
                if (IsDetailBody(body))
                {
                    details.Add(new PartModel(body, _context));
                }
            }
        }

        private bool IsDetailBody(IBody7 body)
        {
            object val;
            bool fromSource;
            ((IPropertyKeeper)body).GetPropertyValue(
                (_Property)_context.SpecificationSectionProperty, 
                out val, 
                false, 
                out fromSource);
            
            return (val as string) == DetailsSection;
        }

        private void ExtractSubParts(IPart7 part, List<PartModel> details)
        {
            IParts7 partsCollection = null;

            try
            {
                partsCollection = part.Parts;
                
                if (partsCollection == null)
                    return;

                foreach (IPart7 subPart in partsCollection)
                {
                    if (subPart == null)
                        continue;

                    try
                    {
                        if (subPart.Detail == true)
                        {
                            details.Add(new PartModel(subPart, _context));
                        }
                        else
                        {
                            ExtractParts(subPart, details);
                        }
                    }
                    catch (COMException)
                    {
                        // Игнорируем ошибки доступа к недоступным частям
                    }
                }
            }
            finally
            {
                ReleaseComObject(partsCollection);
            }
        }

        public void ShowDetailInKompas(PartModel detail)
        {
            if (!IsValidDetailForDisplay(detail))
                return;

            SelectDetail(detail);
            FocusOnDetail(detail);
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
            public double X1, Y1, Z1;
            public double X2, Y2, Z2;
        }

        private Gabarit GetPartGabarit(IPart7 part)
        {
            var gabarit = new Gabarit();
            part.GetGabarit(false, false,
                out gabarit.X1, out gabarit.Y1, out gabarit.Z1,
                out gabarit.X2, out gabarit.Y2, out gabarit.Z2);
            return gabarit;
        }

        private (double x, double y, double z) CalculateCenter(Gabarit gabarit)
        {
            return (
                x: (gabarit.X1 + gabarit.X2) / 2.0,
                y: (gabarit.Y1 + gabarit.Y2) / 2.0,
                z: (gabarit.Z1 + gabarit.Z2) / 2.0
            );
        }

        private double CalculateScale(Gabarit gabarit)
        {
            double width = Math.Abs(gabarit.X2 - gabarit.X1);
            double height = Math.Abs(gabarit.Y2 - gabarit.Y1);
            double depth = Math.Abs(gabarit.Z2 - gabarit.Z1);

            double maxSize = Math.Max(Math.Max(width, height), depth);

            return maxSize > 0 ? ScaleFactor / maxSize : DefaultScale;
        }

        private void UpdateCamera((double x, double y, double z) center, double scale)
        {
            var matrix = _context.ViewProjectionManager.Matrix3D;

            // Элементы [12], [13], [14] отвечают за позицию камеры в матрице 4x4
            matrix[12] = center.x;
            matrix[13] = center.y;
            matrix[14] = center.z;

            _context.ViewProjectionManager.SetMatrix3D(matrix, scale);
        }

        private void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
        }

        public void Dispose()
        {
            _context?.Dispose();
        }
    }
}
