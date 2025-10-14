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

        public void CloseDocument()
        {
            if (Document != null)
            {
                // Освобождаем COM-объекты
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