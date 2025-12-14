using Kompas6API5;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Контекст работы с KOMPAS-3D документом
    /// </summary>
    public class KompasContext : IDisposable
    {
        public IApplication Application { get; private set; }
        public IKompasDocument3D Document { get; private set; }
        public IPart7 TopPart { get; private set; }
        public IChooseManager ChooseManager { get; private set; }
        public ISelectionManager SelectionManager { get; private set; }
        public IViewProjectionManager ViewProjectionManager { get; private set; }
        public IPropertyMng PropertyManager => (IPropertyMng)Application;
        public IProperty SpecificationSectionProperty { get; private set; }

        public bool IsInitialized => Application != null;
        public bool IsDocumentLoaded => Document != null && TopPart != null;

        public KompasContext()
        {
            Application = GetKompasApplication();
        }

        /// <summary>
        /// Загружает документ из файла
        /// </summary>
        /// <param name="filePath">Путь к файлу документа</param>
        public void LoadDocument(string filePath)
        {
            ValidateApplication();

            Document = Application.Documents.Open(filePath) as IKompasDocument3D;

            if (Document != null)
            {
                InitializeDocumentComponents();
            }
        }

        /// <summary>
        /// Загружает активный документ
        /// </summary>
        public void LoadActiveDocument()
        {
            ValidateApplication();

            Document = Application.ActiveDocument as IKompasDocument3D;

            if (Document == null)
                throw new InvalidOperationException("Нет активного 3D документа в KOMPAS-3D");

            InitializeDocumentComponents();
        }

        /// <summary>
        /// Инициализирует компоненты документа
        /// </summary>
        private void InitializeDocumentComponents()
        {
            TopPart = Document.TopPart;
            ChooseManager = Document.ChooseManager;
            SelectionManager = Document.SelectionManager;
            ViewProjectionManager = ((IKompasDocument3D1)Document).ViewProjectionManager;
            SpecificationSectionProperty = PropertyManager.GetProperty(Document, "Раздел спецификации");
        }

        /// <summary>
        /// Получает тип детали из свойств спецификации
        /// </summary>
        /// <param name="part">Деталь для анализа</param>
        /// <returns>Тип детали или null</returns>
        public string GetDetailType(IPart7 part)
        {
            if (SpecificationSectionProperty == null || part == null)
                return null;

            IPropertyKeeper propertyKeeper = null;
            try
            {
                propertyKeeper = part as IPropertyKeeper;
                if (propertyKeeper == null)
                    return null;

                propertyKeeper.GetPropertyValue(
                    (KompasAPI7._Property)SpecificationSectionProperty, 
                    out object markingObj, 
                    false, 
                    out bool fromSource);
                
                return markingObj?.ToString();
            }
            finally
            {
                ReleaseComObjectIfNeeded(propertyKeeper, part);
            }
        }

        /// <summary>
        /// Получает длину детали по выдавливанию
        /// </summary>
        /// <param name="part">Деталь для анализа</param>
        /// <returns>Длина детали</returns>
        public double GetDetailLengthByExtrusion(IPart7 part)
        {
            if (part == null)
                return 0.0;

            // Пытаемся получить длину из выдавливания
            double extrusionDepth = TryGetExtrusionDepth(part);
            if (extrusionDepth > 0)
                return extrusionDepth;

            // Пытаемся получить длину из свойства "Длина профиля"
            IFeature7 feature = part as IFeature7;
            if (feature != null)
            {
                IBody7 body = feature.ResultBodies;
                string lengthStr = GetBodyPropertyValue(body, "Длина профиля");
                if (double.TryParse(lengthStr, out double length))
                    return length;
            }

            // Возвращаем максимальный габарит
            var partGabarit = KompasCameraController.GetPartGabarit(part);
            return partGabarit.MaxSize;
        }

        /// <summary>
        /// Пытается получить глубину выдавливания
        /// </summary>
        private double TryGetExtrusionDepth(IPart7 part)
        {
            try
            {
                IModelContainer modelContainer = part as IModelContainer;
                if (modelContainer == null || modelContainer.Extrusions.Count == 0)
                    return 0.0;

                IExtrusion extrusion = (IExtrusion)modelContainer.Extrusions[0];
                if (extrusion == null)
                    return 0.0;

                var opRes = (extrusion as IExtrusion1).OperationResult;
                bool isValidOperation = opRes == Kompas6Constants3D.ksOperationResultEnum.ksOperationNewBody 
                                     || opRes == Kompas6Constants3D.ksOperationResultEnum.ksOperationUnion;

                if (extrusion.Depth[true] != 0 && isValidOperation)
                    return extrusion.Depth[true];

                return 0.0;
            }
            catch
            {
                return 0.0;
            }
        }

        /// <summary>
        /// Получает значение свойства тела
        /// </summary>
        /// <param name="body">Тело детали</param>
        /// <param name="propertyName">Имя свойства</param>
        /// <returns>Значение свойства или null</returns>
        public string GetBodyPropertyValue(IBody7 body, string propertyName)
        {
            if (body == null || string.IsNullOrEmpty(propertyName))
                return null;

            IProperty property = null;
            IPropertyKeeper propertyKeeper = null;
            IKompasDocument3D parentDocument3D = null;
            
            try
            {
                IPart7 parentPart = body.Parent as IPart7;
                parentDocument3D = Application.Documents.Open(parentPart.FileName, false, true) as IKompasDocument3D;

                property = PropertyManager.GetProperty(parentDocument3D, propertyName);
                if (property == null)
                    return null;

                propertyKeeper = body as IPropertyKeeper;
                if (propertyKeeper == null)
                    return null;

                propertyKeeper.GetPropertyValue(
                    (KompasAPI7._Property)property, 
                    out object markingObj, 
                    false, 
                    out bool fromSource);
                
                return markingObj?.ToString();
            }
            finally
            {
                ReleaseComObject(parentDocument3D);
                ReleaseComObject(property);
                ReleaseComObjectIfNeeded(propertyKeeper, body);
            }
        }

        /// <summary>
        /// Закрывает документ и освобождает ресурсы
        /// </summary>
        public void CloseDocument()
        {
            if (Document != null)
            {
                ReleaseComObject(SpecificationSectionProperty);
                ReleaseComObject(ViewProjectionManager);
                ReleaseComObject(SelectionManager);
                ReleaseComObject(ChooseManager);
                ReleaseComObject(TopPart);
                ReleaseComObject(Document);

                SpecificationSectionProperty = null;
                ViewProjectionManager = null;
                SelectionManager = null;
                ChooseManager = null;
                TopPart = null;
                Document = null;
            }
        }

        /// <summary>
        /// Проверяет, что приложение инициализировано
        /// </summary>
        private void ValidateApplication()
        {
            if (Application == null)
                throw new InvalidOperationException("Не удалось подключиться к KOMPAS-3D");
        }

        /// <summary>
        /// Получает экземпляр приложения KOMPAS
        /// </summary>
        private static IApplication GetKompasApplication()
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

        /// <summary>
        /// Освобождает COM-объект
        /// </summary>
        private void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch { }
            }
        }

        /// <summary>
        /// Освобождает COM-объект, если он отличается от базового объекта
        /// </summary>
        private void ReleaseComObjectIfNeeded(object comObject, object baseObject)
        {
            if (comObject != null && comObject != (object)baseObject && Marshal.IsComObject(comObject))
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch { }
            }
        }

        public void Dispose()
        {
            CloseDocument();
            
            // Не освобождаем Application, т.к. получен через GetActiveObject
            Application = null;
            
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}