using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Kompas6Constants;
using Kompas6Constants3D;
using KompasAPI7;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Создаёт PNG-превью изделий из библиотек КОМПАС по геометрии вхождения в сборку.
    /// Файл такого изделия — общий шаблон, а типоразмер задаётся в сборке, поэтому вхождение
    /// копируется (аналог Ctrl+C / Ctrl+V) в новую сборку без окна, которая закрывается без сохранения.
    /// Исходная сборка при этом не меняется.
    /// </summary>
    public class StandardPartPreviewService
    {
        /// <summary>
        /// Префикс имени PNG-превью изделия из библиотеки
        /// </summary>
        public const string FilePrefix = "std_";

        private const int RasterResolution = 200;
        private const byte BackgroundThreshold = 245;
        private const int CropPadding = 10;

        private static readonly ksObj3dTypeEnum[] DefaultObjectsToHide =
        {
            ksObj3dTypeEnum.o3d_planeXOY,
            ksObj3dTypeEnum.o3d_planeXOZ,
            ksObj3dTypeEnum.o3d_planeYOZ,
            ksObj3dTypeEnum.o3d_pointCS
        };

        private readonly ILogger _logger;

        // Выключается, если копирование неожиданно пометило исходную сборку изменённой
        private static volatile bool _disabled;

        public StandardPartPreviewService(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Путь к PNG-превью вхождения: из кэша или созданный заново. null, если создать не удалось.
        /// </summary>
        public string GetOrCreatePreview(IPart7 part, PartModel model, KompasContext context, string imagesFolder)
        {
            if (part == null || model == null || context?.Application == null || string.IsNullOrEmpty(imagesFolder))
                return null;

            string pngPath = Path.Combine(imagesFolder, GetPreviewFileName(model));
            if (File.Exists(pngPath))
                return pngPath;

            if (_disabled)
                return null;

            Directory.CreateDirectory(imagesFolder);
            string rawPath = Path.Combine(Path.GetTempPath(), $"TankManager_std_{Guid.NewGuid():N}.png");
            bool sourceChangedBefore = IsChanged(context.Document);
            IKompasDocument3D document = null;
            IPart7 copy = null;

            try
            {
                document = context.Application.Documents.Add(DocumentTypeEnum.ksDocumentAssembly, false) as IKompasDocument3D;
                if (document?.TopPart == null)
                {
                    _logger.LogWarning($"Не удалось создать временную сборку для превью {model.Name}");
                    return null;
                }

                copy = document.TopPart.Parts.CopyPart((Part7)part, document.TopPart.Placement);
                if (copy == null)
                {
                    _logger.LogWarning($"Не удалось скопировать изделие {model.Name} во временную сборку");
                    return null;
                }

                HideDefaultObjects(document.TopPart);
                document.TopPart.RebuildModel(true);

                if (!SaveRaster(document, rawPath))
                {
                    _logger.LogWarning($"Не удалось сохранить растр изделия {model.Name}");
                    return null;
                }

                if (!CropToContent(rawPath, pngPath))
                {
                    _logger.LogWarning($"Пустой растр изделия {model.Name}");
                    return null;
                }

                return pngPath;
            }
            catch (COMException ex)
            {
                _logger.LogWarning($"Ошибка КОМПАС при создании превью изделия {model.Name}: {ex.Message}");
                return null;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ExternalException)
            {
                _logger.LogWarning($"Ошибка файла при создании превью изделия {model.Name}: {ex.Message}");
                return null;
            }
            finally
            {
                Release(copy);
                CloseDocument(document);
                TryDelete(rawPath);

                if (!sourceChangedBefore && IsChanged(context.Document))
                {
                    _disabled = true;
                    _logger.LogError($"После копирования изделия {model.Name} исходная сборка помечена изменённой. " +
                                     "Генерация превью изделий из библиотек отключена до перезапуска.");
                }
            }
        }

        /// <summary>
        /// Имя PNG: типоразмеры одного шаблона различаются наименованием и обозначением
        /// </summary>
        public static string GetPreviewFileName(PartModel model)
        {
            string key = $"{model.FilePath}|{model.Name}|{model.Marking}".ToLowerInvariant();
            string hash;
            using (var md5 = MD5.Create())
            {
                hash = BitConverter.ToString(md5.ComputeHash(Encoding.UTF8.GetBytes(key))).Replace("-", "").Substring(0, 12);
            }

            string baseName = Path.GetFileNameWithoutExtension(model.FilePath ?? string.Empty);
            if (string.IsNullOrEmpty(baseName))
                baseName = "part";
            baseName = string.Join("_", baseName.Split(Path.GetInvalidFileNameChars()));

            return $"{FilePrefix}{baseName}_{hash}.png";
        }

        /// <summary>
        /// Превью создано этим сервисом (а не миниатюра файла-шаблона)
        /// </summary>
        public static bool IsStandardPartPreview(string pngPath)
        {
            return !string.IsNullOrEmpty(pngPath) &&
                   Path.GetFileName(pngPath).StartsWith(FilePrefix, StringComparison.OrdinalIgnoreCase);
        }

        private void HideDefaultObjects(IPart7 topPart)
        {
            foreach (var type in DefaultObjectsToHide)
            {
                IModelObject obj = null;
                try
                {
                    obj = topPart.DefaultObject[type];
                    if (obj == null)
                        continue;

                    obj.Hidden = true;
                    obj.Update();
                }
                catch (COMException ex)
                {
                    _logger.LogWarning($"Не удалось скрыть системный объект {type}: {ex.Message}");
                }
                finally
                {
                    Release(obj);
                }
            }
        }

        private static bool SaveRaster(IKompasDocument3D document, string rawPath)
        {
            var document1 = document as IKompasDocument1;
            if (document1 == null)
                return false;

            var rasterParams = (IRasterConvertParameters)document1.GetInterface(KompasAPIObjectTypeEnum.ksObjectRasterConvertParameters);
            try
            {
                rasterParams.ColorType = ksObjectColorTypeEnum.ksColorObject;
                rasterParams.Resolution = RasterResolution;
                rasterParams.ColorBPP = ksColorBPPEnum.ksColorBPP_24;
                rasterParams.RasterFormat = ksRasterFormatEnum.ksRasterFormatPNG;
                rasterParams.Scale = 1;
                rasterParams.Uncompressed = false;

                return document1.SaveAsToRasterFormat(rawPath, (RasterConvertParameters)rasterParams) && File.Exists(rawPath);
            }
            finally
            {
                Release(rasterParams);
            }
        }

        /// <summary>
        /// Обрезает белые поля: КОМПАС рисует модель в углу растра размером с рабочую область
        /// </summary>
        private static bool CropToContent(string sourcePath, string targetPath)
        {
            using (var source = new Bitmap(sourcePath))
            {
                var bounds = FindContentBounds(source);
                if (bounds.IsEmpty)
                    return false;

                var crop = Rectangle.FromLTRB(
                    Math.Max(0, bounds.Left - CropPadding),
                    Math.Max(0, bounds.Top - CropPadding),
                    Math.Min(source.Width, bounds.Right + CropPadding),
                    Math.Min(source.Height, bounds.Bottom + CropPadding));

                using (var cropped = source.Clone(crop, PixelFormat.Format24bppRgb))
                {
                    cropped.Save(targetPath, ImageFormat.Png);
                }
            }

            return true;
        }

        private static Rectangle FindContentBounds(Bitmap bitmap)
        {
            var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                var row = new byte[data.Width * 4];
                int minX = int.MaxValue, minY = int.MaxValue, maxX = -1, maxY = -1;

                for (int y = 0; y < data.Height; y++)
                {
                    Marshal.Copy(data.Scan0 + y * data.Stride, row, 0, row.Length);
                    for (int x = 0; x < data.Width; x++)
                    {
                        int i = x * 4; // BGRA
                        if (row[i] >= BackgroundThreshold && row[i + 1] >= BackgroundThreshold && row[i + 2] >= BackgroundThreshold)
                            continue;

                        if (x < minX) minX = x;
                        if (x > maxX) maxX = x;
                        if (y < minY) minY = y;
                        if (y > maxY) maxY = y;
                    }
                }

                return maxX < 0 ? Rectangle.Empty : Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }

        private static bool IsChanged(IKompasDocument3D document)
        {
            try
            {
                return (document as IKompasDocument)?.Changed ?? false;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private void CloseDocument(IKompasDocument3D document)
        {
            if (document == null)
                return;

            try
            {
                (document as IKompasDocument)?.Close(DocumentCloseOptions.kdDoNotSaveChanges);
            }
            catch (COMException ex)
            {
                _logger.LogWarning($"Не удалось закрыть временную сборку: {ex.Message}");
            }
            finally
            {
                Release(document);
            }
        }

        private static void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
                Marshal.ReleaseComObject(comObject);
        }

        private void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                _logger.LogWarning($"Не удалось удалить временный файл {path}: {ex.Message}");
            }
        }
    }
}
