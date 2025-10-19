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
        public bool IsDocumentLoaded => Document != null;

        public KompasContext()
        {
            Application = GetKompasApplication();
        }

        public void LoadDocument(string filePath)
        {
            if (Application == null)
                throw new InvalidOperationException("Не удалось подключиться к KOMPAS-3D");

            Document = Application.Documents.Open(filePath) as IKompasDocument3D;

            if (Document != null)
            {
                TopPart = Document.TopPart;
                ChooseManager = Document.ChooseManager;
                SelectionManager = Document.SelectionManager;
                ViewProjectionManager = ((IKompasDocument3D1)Document).ViewProjectionManager;
                SpecificationSectionProperty = PropertyManager.GetProperty(Document, "Раздел спецификации");
            }
        }

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
                // Приведение типа (as) не создает новый RCW, поэтому освобождать не нужно
                // Проверяем на всякий случай
                if (propertyKeeper != null && propertyKeeper != (object)part && Marshal.IsComObject(propertyKeeper))
                {
                    try
                    {
                        Marshal.ReleaseComObject(propertyKeeper);
                    }
                    catch { }
                }
            }
        }

        public string GetBodyPropertyValue(IBody7 body, string propertyName)
        {
            if (body == null || string.IsNullOrEmpty(propertyName))
                return null;

            IProperty property = null;
            IPropertyKeeper propertyKeeper = null;
            
            try
            {
                property = PropertyManager.GetProperty(Document, propertyName);
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
                // Освобождаем property, если это COM-объект
                if (property != null && Marshal.IsComObject(property))
                {
                    try
                    {
                        Marshal.ReleaseComObject(property);
                    }
                    catch { }
                }

                // Приведение типа не создает новый RCW
                if (propertyKeeper != null && propertyKeeper != (object)body && Marshal.IsComObject(propertyKeeper))
                {
                    try
                    {
                        Marshal.ReleaseComObject(propertyKeeper);
                    }
                    catch { }
                }
            }
        }

        public void CloseDocument()
        {
            if (Document != null)
            {
                // Освобождаем COM-объекты в обратном порядке
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