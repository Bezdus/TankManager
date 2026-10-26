using System;
using System.IO;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Атомарная запись файла: данные пишутся во временный файл, который затем подменяет целевой.
    /// При сбое на середине записи исходный файл остаётся нетронутым, а читатель видит либо старый,
    /// либо новый файл целиком.
    /// </summary>
    public static class AtomicFile
    {
        public static void WriteAllBytes(string path, byte[] data)
        {
            string tempPath = GetTempPath(path);
            try
            {
                File.WriteAllBytes(tempPath, data);
                ReplaceWith(tempPath, path, null);
            }
            finally
            {
                TryDelete(tempPath);
            }
        }

        /// <summary>
        /// Копирует файл с подменой целевого. Дата изменения берётся у исходного (по ней работает синхронизация).
        /// backupPath — куда сохранить прежнюю версию целевого файла (null — не сохранять)
        /// </summary>
        public static void Copy(string sourcePath, string destPath, string backupPath = null)
        {
            string tempPath = GetTempPath(destPath);
            try
            {
                File.Copy(sourcePath, tempPath, true);
                ReplaceWith(tempPath, destPath, backupPath);
                File.SetLastWriteTimeUtc(destPath, File.GetLastWriteTimeUtc(sourcePath));
            }
            finally
            {
                TryDelete(tempPath);
            }
        }

        // Своё имя у каждой записи: в общую папку один и тот же файл могут писать несколько компьютеров
        private static string GetTempPath(string path) =>
            path + "." + Guid.NewGuid().ToString("N").Substring(0, 8) + ".tmp";

        private static void ReplaceWith(string tempPath, string path, string backupPath)
        {
            if (File.Exists(path))
            {
                if (backupPath != null && File.Exists(backupPath))
                    File.Delete(backupPath);
                File.Replace(tempPath, path, backupPath);
            }
            else
            {
                File.Move(tempPath, path);
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // Временный файл с уникальным именем никому не мешает
            }
        }
    }
}
