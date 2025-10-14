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
            var result = new List<PartModel>();

            if (!_context.IsInitialized)
                throw new InvalidOperationException("Не удалось подключиться к KOMPAS-3D");

            _context.LoadDocument(filePath);

            if (_context.IsDocumentLoaded)
            {
                ExtractParts(_context.TopPart, result);
            }

            return result;
        }

        private void ExtractParts(IPart7 part, List<PartModel> details)
        {
            object partsEnum = null;
            try
            {
                // Получаем коллекцию безопасно
                var partsCollection = part.Parts;
                partsEnum = partsCollection.GetEnumerator();
                
                while (true)
                {
                    IPart7 subPart = null;
                    try
                    {
                        // Используем явный вызов MoveNext для контроля
                        if (partsCollection.Count == 0)
                            break;

                        foreach (IPart7 p in partsCollection)
                        {
                            subPart = p;
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
                            finally
                            {
                                // Не освобождаем subPart здесь, т.к. он используется в PartModel
                            }
                        }
                        break;
                    }
                    catch
                    {
                        if (subPart != null && Marshal.IsComObject(subPart))
                            Marshal.ReleaseComObject(subPart);
                        throw;
                    }
                }

                // Освобождаем коллекцию
                if (partsCollection != null && Marshal.IsComObject(partsCollection))
                    Marshal.ReleaseComObject(partsCollection);
            }
            finally
            {
                if (partsEnum != null && Marshal.IsComObject(partsEnum))
                    Marshal.ReleaseComObject(partsEnum);
            }
        }

        public void ShowDetailInKompas(PartModel detail)
        {
            if (!_context.IsDocumentLoaded || detail?.Part == null)
                return;

            _context.SelectionManager.UnselectAll();
            _context.SelectionManager.Select(detail.Part);

            // Получаем габариты детали
            detail.Part.GetGabarit(false, false, 
                out double X1, out double Y1, out double Z1, 
                out double X2, out double Y2, out double Z2);

            // Вычисляем центр габарита
            double centerX = (X1 + X2) / 2.0;
            double centerY = (Y1 + Y2) / 2.0;
            double centerZ = (Z1 + Z2) / 2.0;

            // Вычисляем размеры габарита
            double width = Math.Abs(X2 - X1);
            double height = Math.Abs(Y2 - Y1);
            double depth = Math.Abs(Z2 - Z1);

            // Находим максимальный размер для расчета масштаба
            double maxSize = Math.Max(Math.Max(width, height), depth);

            // Вычисляем масштаб (обратная логика - больше размер = меньше масштаб)
            // Типичный диапазон масштабов: 0.1 - 10.0
            double scale = maxSize > 0 ? 100.0 / maxSize : 1.0;

            // Получаем существующую матрицу
            var matrix = _context.ViewProjectionManager.Matrix3D;
            
            // Модифицируем позицию камеры (элементы трансляции)
            // В матрице 4x4 элементы [12], [13], [14] отвечают за позицию
            matrix[12] = centerX;
            matrix[13] = centerY;
            matrix[14] = centerZ;

            _context.ViewProjectionManager.SetMatrix3D(matrix, scale);
        }

        public void Dispose()
        {
            _context?.Dispose();
        }
    }
}
