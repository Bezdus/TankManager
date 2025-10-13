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
        private IApplication _app;
        private IPart7 _part;
        private IProperty _property;
        private IChooseManager _chooseManager;
        private ISelectionManager _selectionManager;
        private IViewProjectionManager _viewProjectionManager;

        public KompasService()
        {
            _app = GetKompas();
        }

        public List<PartModel> LoadDocument(string filePath)
        {
            var result = new List<PartModel>();

            if (_app == null)
                throw new InvalidOperationException("Не удалось подключиться к KOMPAS-3D");

            IKompasDocument3D kompasDocument3D = _app.Documents.Open(filePath) as IKompasDocument3D;
            _chooseManager = kompasDocument3D.ChooseManager;
            _selectionManager = kompasDocument3D.SelectionManager;
            _viewProjectionManager = ((IKompasDocument3D1)kompasDocument3D).ViewProjectionManager;

            if (kompasDocument3D != null)
            {
                _part = kompasDocument3D.TopPart;
                _property = ((IPropertyMng)_app).GetProperty(kompasDocument3D, "Раздел спецификации");
                ExtractParts(_part, result);
            }

            return result;
        }

        private void ExtractParts(IPart7 part, List<PartModel> details)
        {
            foreach (IPart7 subPart in part.Parts)
            {
                if (subPart.Detail == true)
                {
                    details.Add(new PartModel(subPart, GetDetailType(subPart)));
                }
                else
                {
                    ExtractParts(subPart, details);
                }
            }
        }

        private string GetDetailType(IPart7 part)
        {
            object markingObj;
            bool fromSource;
            IPropertyKeeper propertyKeeper = (IPropertyKeeper)part;
            propertyKeeper.GetPropertyValue((KompasAPI7._Property)_property, out markingObj, false, out fromSource);
            return markingObj?.ToString();
        }

        public void ShowDetailInKompas(PartModel detail)
        {
            if (_app == null || detail?.Part == null)
                return;

            _selectionManager.UnselectAll();
            _selectionManager.Select(detail.Part);



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
            var matrix = _viewProjectionManager.Matrix3D;
            
            // Модифицируем позицию камеры (элементы трансляции)
            // В матрице 4x4 элементы [12], [13], [14] отвечают за позицию
            matrix[12] = centerX;
            matrix[13] = centerY;
            matrix[14] = centerZ;

            _viewProjectionManager.SetMatrix3D(matrix, scale);
        }

        private static IApplication GetKompas()
        {
            try
            {
                return (IApplication)Marshal.GetActiveObject("KOMPAS.Application.7");
            }
            catch
            {
                return null;
            }
        }

        public void Dispose()
        {
            _app = null;
            _property = null;
        }
    }
}
