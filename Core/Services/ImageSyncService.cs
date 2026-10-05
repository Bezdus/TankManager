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
        /// <param name="localDir">Локальная папка изображений</param>
        /// <param name="serverDir">Серверная папка изображений</param>
        /// <param name="downloadOnly">Только с сервера в локальную папку (режим просмотра)</param>
        /// <returns>Количество файлов, скачанных с сервера в локальную папку</returns>
        public int SyncImageDirectories(string localDir, string serverDir, bool downloadOnly = false)
        {
            if (string.IsNullOrEmpty(localDir) || string.IsNullOrEmpty(serverDir))
                return 0;

            if (downloadOnly)
            {
                if (!Directory.Exists(serverDir))
                    return 0;

                Directory.CreateDirectory(localDir);
                return SyncOneWay(serverDir, localDir);
            }

            if (!Directory.Exists(localDir) && !Directory.Exists(serverDir))
                return 0;

            Directory.CreateDirectory(localDir);
            Directory.CreateDirectory(serverDir);

            SyncOneWay(localDir, serverDir);
            return SyncOneWay(serverDir, localDir);
        }

        private int SyncOneWay(string sourceDir, string destDir)
        {
            int copied = 0;

            if (!Directory.Exists(sourceDir))
                return copied;

            foreach (var sourceFile in Directory.GetFiles(sourceDir))
            {
                try
                {
                    string fileName = Path.GetFileName(sourceFile);
                    string destFile = Path.Combine(destDir, fileName);

                    if (!File.Exists(destFile))
                    {
                        File.Copy(sourceFile, destFile, false);
                        copied++;
                    }
                    else
                    {
                        var sourceTime = File.GetLastWriteTimeUtc(sourceFile);
                        var destTime = File.GetLastWriteTimeUtc(destFile);
                        if (sourceTime > destTime)
                        {
                            File.Copy(sourceFile, destFile, true);
                            copied++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка синхронизации изображения {Path.GetFileName(sourceFile)}: {ex.Message}");
                }
            }

            return copied;
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
    }
}
