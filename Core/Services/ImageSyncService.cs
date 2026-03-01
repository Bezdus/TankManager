using System;
using System.Diagnostics;
using System.IO;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Сервис синхронизации изображений между локальной и серверной папками
    /// </summary>
    public class ImageSyncService
    {
        /// <summary>
        /// Двусторонняя синхронизация изображений между двумя директориями.
        /// Более новые файлы перезаписывают старые.
        /// </summary>
        public void SyncImageDirectories(string dir1, string dir2)
        {
            if (string.IsNullOrEmpty(dir1) || string.IsNullOrEmpty(dir2))
                return;
            if (!Directory.Exists(dir1) && !Directory.Exists(dir2))
                return;

            Directory.CreateDirectory(dir1);
            Directory.CreateDirectory(dir2);

            SyncOneWay(dir1, dir2);
            SyncOneWay(dir2, dir1);
        }

        private void SyncOneWay(string sourceDir, string destDir)
        {
            if (!Directory.Exists(sourceDir))
                return;

            foreach (var sourceFile in Directory.GetFiles(sourceDir))
            {
                try
                {
                    string fileName = Path.GetFileName(sourceFile);
                    string destFile = Path.Combine(destDir, fileName);

                    if (!File.Exists(destFile))
                    {
                        File.Copy(sourceFile, destFile, false);
                    }
                    else
                    {
                        var sourceTime = File.GetLastWriteTimeUtc(sourceFile);
                        var destTime = File.GetLastWriteTimeUtc(destFile);
                        if (sourceTime > destTime)
                        {
                            File.Copy(sourceFile, destFile, true);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка синхронизации изображения {Path.GetFileName(sourceFile)}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Проверяет, устарело ли превью чертежа (PNG) по сравнению с .cdw файлом
        /// </summary>
        public static bool IsDrawingPreviewStale(string pngPath, string cdwPath)
        {
            if (string.IsNullOrEmpty(cdwPath) || !File.Exists(cdwPath))
                return false;

            if (string.IsNullOrEmpty(pngPath) || !File.Exists(pngPath))
                return true;

            return File.GetLastWriteTimeUtc(cdwPath) > File.GetLastWriteTimeUtc(pngPath);
        }

        /// <summary>
        /// Проверяет, устарело ли превью 3D-файла по сравнению с исходным файлом
        /// </summary>
        public static bool IsFilePreviewStale(string previewPngPath, string sourceFilePath)
        {
            if (string.IsNullOrEmpty(sourceFilePath) || !File.Exists(sourceFilePath))
                return false;

            if (string.IsNullOrEmpty(previewPngPath) || !File.Exists(previewPngPath))
                return true;

            return File.GetLastWriteTimeUtc(sourceFilePath) > File.GetLastWriteTimeUtc(previewPngPath);
        }
    }
}
