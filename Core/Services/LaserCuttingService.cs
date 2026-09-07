using Kompas6Constants;
using KompasAPI7;
using System;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Считает длину реза и гравировки листовой детали по её DXF-файлу.
    /// DXF конвертируется в FRW через конвертор KOMPAS, затем анализируются
    /// объекты чертежа: кривые слоя "Scribe" — гравировка, остальные — рез.
    /// </summary>
    public class LaserCuttingService
    {
        private const string ScribeLayerName = "Scribe";
        private const string TempFolderName = "TempFrw";

        private readonly ILogger _logger;

        public LaserCuttingService(ILogger logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Вычисляет операцию лазерной резки по DXF-файлу.
        /// Возвращает null при любой ошибке (нет KOMPAS, нет конвертора и т.п.)
        /// </summary>
        public LaserCuttingOperation Compute(string dxfFilePath, IApplication application)
        {
            if (string.IsNullOrEmpty(dxfFilePath) || application == null)
                return null;

            string converterLibrary = KompasPathResolver.ResolveConverterLibrary();
            if (string.IsNullOrEmpty(converterLibrary))
            {
                _logger.LogWarning("Не найдена библиотека конвертора dwgdxfImp.rtw");
                return null;
            }

            IConverter converter = null;
            IKompasDocument2D document = null;
            IView view = null;
            IDrawingContainer drawingContainer = null;
            ILayers layers = null;
            object drawingObjects = null;
            string frwFilePath = null;

            try
            {
                converter = application.Converter[converterLibrary];
                if (converter == null)
                    return null;

                string dxfDirectory = Path.GetDirectoryName(dxfFilePath);
                if (string.IsNullOrEmpty(dxfDirectory))
                    dxfDirectory = Path.GetTempPath();

                string tempFolder = Path.Combine(dxfDirectory, TempFolderName);
                Directory.CreateDirectory(tempFolder);

                frwFilePath = Path.Combine(
                    tempFolder,
                    Path.GetFileNameWithoutExtension(dxfFilePath) + ".frw");
                converter.Convert(dxfFilePath, frwFilePath, 0, false);
                if (!File.Exists(frwFilePath))
                    return null;

                document = application.Documents.Open(frwFilePath, false, false) as IKompasDocument2D;
                if (document == null)
                    return null;

                view = document.ViewsAndLayersManager.Views.ActiveView;
                if (view == null)
                    return null;

                drawingContainer = view as IDrawingContainer;
                if (drawingContainer == null)
                    return null;

                layers = view.Layers;
                drawingObjects = drawingContainer.Objects[DrawingObjectTypeEnum.ksAllObj];

                double cuttingLength = 0;
                double scribeLength = 0;

                if (drawingObjects is IEnumerable enumerable)
                {
                    foreach (object obj in enumerable)
                    {
                        IDrawingObject drawingObject = obj as IDrawingObject;
                        if (drawingObject == null)
                            continue;

                        try
                        {
                            Curve2D curve = ((IDrawingObject1)drawingObject).GetCurve2D();
                            if (curve == null)
                                continue;

                            ILayer layer = layers.Layer[drawingObject.LayerNumber];
                            if (layer != null &&
                                string.Equals(layer.Name, ScribeLayerName, StringComparison.OrdinalIgnoreCase))
                                scribeLength += curve.Length;
                            else
                                cuttingLength += curve.Length;
                        }
                        finally
                        {
                            ReleaseComObject(drawingObject);
                        }
                    }
                }

                return new LaserCuttingOperation
                {
                    CutLength = cuttingLength,
                    EngravingLength = scribeLength
                };
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось посчитать рез/гравировку для {dxfFilePath}: {ex.Message}");
                return null;
            }
            finally
            {
                if (document != null)
                {
                    try { document.Close(DocumentCloseOptions.kdDoNotSaveChanges); }
                    catch { }
                }

                ReleaseComObject(drawingObjects);
                ReleaseComObject(layers);
                ReleaseComObject(drawingContainer);
                ReleaseComObject(view);
                ReleaseComObject(document);
                ReleaseComObject(converter);

                if (!string.IsNullOrEmpty(frwFilePath))
                {
                    try { if (File.Exists(frwFilePath)) File.Delete(frwFilePath); }
                    catch { }
                }
            }
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try { Marshal.ReleaseComObject(obj); }
                catch { }
            }
        }
    }
}
