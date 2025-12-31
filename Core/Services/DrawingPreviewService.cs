using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Media.Imaging;
using Kompas6Constants;
using KompasAPI7;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Сервис для создания и кэширования PNG-превью чертежей
    /// </summary>
    public class DrawingPreviewService
    {
        /// <summary>
        /// Получает или создаёт PNG-превью чертежа детали
        /// </summary>
        /// <param name="part">Деталь</param>
        /// <param name="context">Контекст KOMPAS</param>
        /// <param name="sourceCdwPath">Выходной параметр: путь к исходному файлу чертежа</param>
        /// <param name="targetDirectory">Целевая папка для сохранения превью</param>
        /// <returns>Путь к PNG-файлу превью</returns>
        public string GetOrCreatePreview(IPart7 part, KompasContext context, out string sourceCdwPath, string targetDirectory)
        {
            sourceCdwPath = null;

            if (part == null || context == null || string.IsNullOrEmpty(targetDirectory))
                return null;

            IKompasDocument3D kompasDocument3D = null;
            IKompasDocument cdwDocument = null;

            try
            {
                // Получаем путь к чертежу
                string cdwPath = GetAttachedDrawingPath(part, ref kompasDocument3D);
                if (string.IsNullOrEmpty(cdwPath) || !File.Exists(cdwPath))
                    return null;

                sourceCdwPath = cdwPath;

                // Создаём целевую папку если её нет
                Directory.CreateDirectory(targetDirectory);

                // Проверяем кэш
                string pngPath = GetCachedPngPath(cdwPath, targetDirectory);
                if (IsCacheValid(cdwPath, pngPath))
                    return pngPath;

                // Удаляем существующий файл, чтобы КОМПАС не показывал диалог подтверждения
                if (File.Exists(pngPath))
                {
                    try { File.Delete(pngPath); }
                    catch { /* Игнорируем ошибку удаления */ }
                }

                // Генерируем PNG
                cdwDocument = context.Application.Documents.Open(cdwPath, false, false);
                if (cdwDocument == null)
                    return null;

                var cdwDocument1 = cdwDocument as IKompasDocument1;
                if (cdwDocument1 == null)
                    return null;

                var rasterParams = (IRasterConvertParameters)cdwDocument1
                    .GetInterface(KompasAPIObjectTypeEnum.ksObjectRasterConvertParameters);
                rasterParams.ColorType = ksObjectColorTypeEnum.ksColorObject;

                cdwDocument1.SaveAsToRasterFormat(pngPath, (RasterConvertParameters)rasterParams);

                return pngPath;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                CloseAndReleaseDocument(cdwDocument);
                CloseAndReleaseDocument(kompasDocument3D);
            }
        }

        /// <summary>
        /// Загружает PNG-изображение для отображения в UI.
        /// Проверяет актуальность кэша перед загрузкой.
        /// </summary>
        /// <param name="pngPath">Путь к PNG-файлу</param>
        /// <param name="sourceCdwPath">Путь к исходному файлу чертежа для проверки актуальности</param>
        /// <returns>Изображение или null, если кэш устарел или файл не существует</returns>
        public BitmapImage LoadPreviewImage(string pngPath, string sourceCdwPath = null)
        {
            if (string.IsNullOrEmpty(pngPath) || !File.Exists(pngPath))
            {
                return null;
            }

            // Если указан исходный файл, проверяем актуальность кэша
            if (!string.IsNullOrEmpty(sourceCdwPath) && !IsCacheValid(sourceCdwPath, pngPath))
            {
                return null;
            }

            try
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(pngPath);
                bitmap.EndInit();
                bitmap.Freeze(); // Для потокобезопасности WPF
                return bitmap;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private string GetAttachedDrawingPath(IPart7 part, ref IKompasDocument3D kompasDocument3D)
        {
            OpenDocumentParam param = part.GetOpenDocumentParam();
            param.Visible = false;
            kompasDocument3D = part.OpenSourceDocument(param);

            if (kompasDocument3D == null)
                return null;

            var propertyKeeper = kompasDocument3D as IPropertyKeeper;
            var productDataManager = kompasDocument3D as IProductDataManager;

            if (propertyKeeper == null || productDataManager == null)
                return null;

            var arrAttachDoc = productDataManager.ObjectAttachedDocuments[propertyKeeper];
            if (arrAttachDoc == null)
                return null;

            return ((object[])arrAttachDoc)
                .Cast<string>()
                .FirstOrDefault(path => Path.GetExtension(path)
                    .Equals(".cdw", StringComparison.OrdinalIgnoreCase));
        }

        private string GetCachedPngPath(string cdwPath, string cacheDirectory)
        {
            string hash = ComputeHash(cdwPath);
            string fileName = Path.GetFileNameWithoutExtension(cdwPath);
            // Убираем недопустимые символы из имени файла
            string safeFileName = string.Join("_", fileName.Split(Path.GetInvalidFileNameChars()));
            return Path.Combine(cacheDirectory, $"{safeFileName}_{hash}.png");
        }

        private static bool IsCacheValid(string sourcePath, string cachedPath)
        {
            if (!File.Exists(cachedPath))
                return false;

            return File.GetLastWriteTimeUtc(cachedPath) >= File.GetLastWriteTimeUtc(sourcePath);
        }

        private static string ComputeHash(string input)
        {
            using (var md5 = MD5.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(input.ToLowerInvariant());
                byte[] hash = md5.ComputeHash(bytes);
                return BitConverter.ToString(hash).Replace("-", "").Substring(0, 8);
            }
        }

        private static void CloseAndReleaseDocument(object document)
        {
            if (document == null)
                return;

            try
            {
                var kompasDoc = document as IKompasDocument;
                if (kompasDoc != null && !kompasDoc.Visible)
                    kompasDoc.Close(DocumentCloseOptions.kdDoNotSaveChanges);
            }
            catch { /* Игнорируем ошибки закрытия */ }
            finally
            {
                if (Marshal.IsComObject(document))
                {
                    try { Marshal.ReleaseComObject(document); }
                    catch { }
                }
            }
        }
    }
}
