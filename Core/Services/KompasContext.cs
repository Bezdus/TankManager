using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Контекст работы с KOMPAS-3D документом
    /// </summary>
    public class KompasContext : IDisposable
    {
        private static readonly ILogger Logger = new FileLogger();

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

        /// <summary>
        /// Проверяет, что COM-ссылка на документ ещё жива (документ не закрыт, КОМПАС не перезапущен)
        /// </summary>
        public bool IsDocumentAlive()
        {
            if (Document == null || TopPart == null)
                return false;

            try
            {
                var name = Document.Name;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            catch (InvalidComObjectException)
            {
                return false;
            }
        }

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
                IBody7 body = feature.ResultBodies as IBody7;
                string lengthStr = GetBodyPropertyValue(body, "Длина профиля");
                if (double.TryParse(lengthStr, out double length))
                    return length;

                // Возвращаем максимальный габарит
                var partGabarit = KompasCameraController.GetPartGabarit(body);
                return partGabarit.MaxSize;
            }

            return 0.0;
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
            return GetBodyPropertyValues(body, propertyName)[0];
        }

        /// <summary>
        /// Получает значения нескольких свойств тела за одно открытие родительского документа
        /// </summary>
        /// <param name="body">Тело детали</param>
        /// <param name="propertyNames">Имена свойств</param>
        /// <returns>Значения в том же порядке (null, если свойство не найдено)</returns>
        public string[] GetBodyPropertyValues(IBody7 body, params string[] propertyNames)
        {
            var result = new string[propertyNames.Length];
            if (body == null || propertyNames.Length == 0)
                return result;

            IKompasDocument3D parentDocument3D = null;
            bool ownsParentDocument = false;
            IPropertyKeeper propertyKeeper = null;

            try
            {
                object parent = body.Parent;
                IPart7 parentPart = parent as IPart7;
                string parentFileName = parentPart?.FileName;

                if (parentPart == null)
                    Logger.LogWarning($"Родитель тела {body.Name} не IPart7 ({parent?.GetType().Name ?? "null"}), свойства читаются из текущего документа");
                else if (string.IsNullOrEmpty(parentFileName))
                    Logger.LogWarning($"У родителя тела {body.Name} ({parentPart.Name}) пустой FileName, свойства читаются из текущего документа");
                else
                {
                    parentDocument3D = GetPropertySourceDocument(parentFileName, out ownsParentDocument);
                    if (parentDocument3D == null)
                        Logger.LogWarning($"Не удалось получить документ {parentFileName} для тела {body.Name}, свойства читаются из текущего документа");
                }

                // Запасной источник описаний свойств — загруженный документ сборки (он не закрывается и не освобождается здесь)
                IKompasDocument3D propertySource = parentDocument3D ?? Document;
                if (propertySource == null)
                    return result;

                propertyKeeper = body as IPropertyKeeper;
                if (propertyKeeper == null)
                {
                    Logger.LogWarning($"Тело {body.Name} не поддерживает IPropertyKeeper");
                    return result;
                }

                for (int i = 0; i < propertyNames.Length; i++)
                {
                    if (string.IsNullOrEmpty(propertyNames[i]))
                        continue;

                    IProperty property = null;
                    try
                    {
                        property = PropertyManager.GetProperty(propertySource, propertyNames[i]);
                        if (property == null)
                        {
                            Logger.LogWarning($"Свойство '{propertyNames[i]}' не найдено в документе {parentFileName} (тело {body.Name})");
                            continue;
                        }

                        propertyKeeper.GetPropertyValue(
                            (KompasAPI7._Property)property,
                            out object value,
                            false,
                            out bool fromSource);

                        result[i] = value?.ToString();

                        if (string.IsNullOrWhiteSpace(result[i]))
                            Logger.LogWarning($"Пустое свойство '{propertyNames[i]}' у тела {body.Name} ({parentFileName}), fromSource={fromSource}, тип={value?.GetType().Name ?? "null"}");
                    }
                    finally
                    {
                        ReleaseComObject(property);
                    }
                }

                return result;
            }
            finally
            {
                if (ownsParentDocument)
                    CloseIfHidden(parentDocument3D);
                if (parentDocument3D != null && !ReferenceEquals(parentDocument3D, Document))
                    ReleaseComObject(parentDocument3D);
                ReleaseComObjectIfNeeded(propertyKeeper, body);
            }
        }

        /// <summary>
        /// Документ, из которого берутся описания свойств тела: уже открытый в КОМПАС (не закрываем)
        /// или открытый нами без окна. Подсборку (.a3d), загруженную в текущую сборку, КОМПАС может
        /// не открыть только на чтение — тогда пробуем обычное открытие.
        /// </summary>
        private IKompasDocument3D GetPropertySourceDocument(string filePath, out bool openedByUs)
        {
            openedByUs = false;

            if (Document != null && string.Equals(GetPathName(Document), filePath, StringComparison.OrdinalIgnoreCase))
                return Document;

            try
            {
                foreach (object docObj in Application.Documents)
                {
                    var doc3D = docObj as IKompasDocument3D;
                    if (doc3D != null && string.Equals(GetPathName(doc3D), filePath, StringComparison.OrdinalIgnoreCase))
                        return doc3D;

                    ReleaseComObject(docObj);
                }
            }
            catch (COMException ex)
            {
                Logger.LogWarning($"Не удалось перебрать открытые документы КОМПАС: {ex.Message}");
            }

            foreach (bool readOnly in new[] { true, false })
            {
                object opened = null;
                try
                {
                    opened = Application.Documents.Open(filePath, false, readOnly);
                }
                catch (COMException ex)
                {
                    Logger.LogWarning($"Ошибка открытия {filePath} (readOnly={readOnly}): {ex.Message}");
                }

                var opened3D = opened as IKompasDocument3D;
                if (opened3D != null)
                {
                    openedByUs = true;
                    return opened3D;
                }

                Logger.LogWarning($"Documents.Open вернул {opened?.GetType().Name ?? "null"} для {filePath} (readOnly={readOnly})");
                if (opened != null)
                {
                    try
                    {
                        var openedDoc = opened as IKompasDocument;
                        if (openedDoc != null && !openedDoc.Visible)
                            openedDoc.Close(DocumentCloseOptions.kdDoNotSaveChanges);
                    }
                    catch (COMException) { }
                    ReleaseComObject(opened);
                }
            }

            return null;
        }

        private static string GetPathName(IKompasDocument3D document)
        {
            try
            {
                return (document as IKompasDocument)?.PathName;
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <summary>
        /// Закрывает документ, открытый без окна (иначе он остаётся загруженным в КОМПАС).
        /// Видимые документы (сборка пользователя) не трогаем.
        /// </summary>
        private static void CloseIfHidden(IKompasDocument3D document)
        {
            if (document == null)
                return;

            try
            {
                var kompasDocument = document as IKompasDocument;
                if (kompasDocument != null && !kompasDocument.Visible)
                    kompasDocument.Close(DocumentCloseOptions.kdDoNotSaveChanges);
            }
            catch (COMException ex)
            {
                Logger.LogWarning($"Не удалось закрыть скрытый документ: {ex.Message}");
            }
        }

        /// <summary>
        /// Получает список технологических операций для листовой детали из дерева построения
        /// </summary>
        /// <param name="part">Деталь для анализа</param>
        /// <returns>Список операций изготовления</returns>
        public List<ManufacturingOperationBase> GetManufacturingOperations(IPart7 part)
        {
            var operations = new List<ManufacturingOperationBase>();

            if (part == null)
                return operations;

            IFeature7 feature = null;
            object modelObjects = null;

            try
            {
                feature = part as IFeature7;
                if (feature == null)
                    return operations;

                modelObjects = feature.SubFeatures[Kompas6Constants.ksTreeTypeEnum.ksOperTree, false, false];
                if (modelObjects == null)
                    return operations;

                if (modelObjects is Array modelArray)
                {
                    foreach (IModelObject modelObject in modelArray)
                    {
                        try
                        {
                            ProcessModelObject(modelObject, operations);
                        }
                        finally
                        {
                            ReleaseComObject(modelObject);
                        }
                    }
                }

                if (operations.Count > 0)
                {
                    double thickness = PartModel.ParseSheetThickness(part.Material);
                    foreach (var op in operations)
                        op.MaterialThickness = thickness;
                }
            }
            catch (Exception ex)
            {
                // Деталь остаётся с частичным списком операций — фиксируем это в логе
                Logger.LogWarning($"Ошибка чтения операций детали {part.Name}: {ex.Message}");
            }
            finally
            {
                ReleaseComObject(modelObjects);
                ReleaseComObjectIfNeeded(feature, part);
            }

            return operations;
        }

        private void ProcessModelObject(IModelObject modelObject, List<ManufacturingOperationBase> operations)
        {
            if (modelObject == null)
                return;

            switch (modelObject.ModelObjectType)
            {
                case Kompas6Constants3D.ksObj3dTypeEnum.o3d_sheetMetalBody:
                    object childrenContainer = null;

                    try
                    {
                        IModelObject1 modelObject1 = modelObject as IModelObject1;
                        if (modelObject1 == null)
                            return;

                        childrenContainer = modelObject1.Childrens[Kompas6Constants3D.ksRelationTypeEnum.ksRTIndifferent];
                        if (childrenContainer is Array childrenArray)
                        {
                            foreach (IModelObject child in childrenArray)
                            {
                                try
                                {
                                    if (child.ModelObjectType == Kompas6Constants3D.ksObj3dTypeEnum.o3d_sheetMetalBendObject)
                                    {
                                        operations.Add(new BendingOperation());
                                    }
                                }
                                finally
                                {
                                    ReleaseComObject(child);
                                }
                            }
                        }
                    }
                    finally
                    {
                        ReleaseComObject(childrenContainer);
                    }
                    break;

                case Kompas6Constants3D.ksObj3dTypeEnum.o3d_sheetMetalBend:
                    ISheetMetalBend sheetMetalBend = modelObject as ISheetMetalBend;
                    if (sheetMetalBend != null)
                    {
                        operations.Add(new BendingOperation
                        {
                            BendAngle = sheetMetalBend.Angle,
                            BendLength = sheetMetalBend.BendValue
                        });
                        ReleaseComObjectIfNeeded(sheetMetalBend, modelObject);
                    }
                    break;

                case Kompas6Constants3D.ksObj3dTypeEnum.o3d_sheetMetalCowling:
                    ISheetMetalRuledShell cowling = modelObject as ISheetMetalRuledShell;
                    if (cowling != null)
                    {
                        // Стоимость вальцовки считается по массе детали (см. RollingOperation.CalculateCost),
                        // поэтому геометрические параметры здесь не нужны
                        operations.Add(new RollingOperation());
                        ReleaseComObjectIfNeeded(cowling, modelObject);
                    }
                    break;

                case Kompas6Constants3D.ksObj3dTypeEnum.o3d_sheetMetalFlanging:
                    ISheetMetalBend flanging = modelObject as ISheetMetalBend;
                    if (flanging != null)
                    {
                        operations.Add(new FlangingOperation
                        {
                            Radius = flanging.Radius
                        });
                        ReleaseComObjectIfNeeded(flanging, modelObject);
                    }
                    break;
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
        }
    }
}