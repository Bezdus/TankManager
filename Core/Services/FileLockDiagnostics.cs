using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Утилита для диагностики проблем с блокировкой файлов
    /// </summary>
    public static class FileLockDiagnostics
    {
        /// <summary>
        /// Проверяет, какие файлы в указанной папке заблокированы
        /// </summary>
        public static List<string> FindLockedFiles(string directoryPath)
        {
            var lockedFiles = new List<string>();

            if (!Directory.Exists(directoryPath))
                return lockedFiles;

            try
            {
                var files = Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories);
                
                foreach (var file in files)
                {
                    if (IsFileLocked(file))
                    {
                        lockedFiles.Add(file);
                        Debug.WriteLine($"?? Заблокирован: {file}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка проверки блокировки файлов: {ex.Message}");
            }

            return lockedFiles;
        }

        /// <summary>
        /// Проверяет, заблокирован ли конкретный файл
        /// </summary>
        public static bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            FileStream stream = null;
            try
            {
                stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return false; // Файл не заблокирован
            }
            catch (IOException)
            {
                return true; // Файл заблокирован
            }
            finally
            {
                stream?.Dispose();
            }
        }

        /// <summary>
        /// Удаляет папку с несколькими попытками (файлы могут быть кратковременно заняты).
        /// Вызывать вне UI-потока: между попытками поток засыпает.
        /// </summary>
        public static bool ForceDeleteDirectory(string directoryPath, int maxAttempts = 5, int delayMs = 500)
        {
            if (!Directory.Exists(directoryPath))
                return true;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                // Снимаем атрибуты «только чтение», иначе Directory.Delete упадёт
                try
                {
                    foreach (var file in Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories))
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Не удалось снять атрибуты файлов: {ex.Message}");
                }

                try
                {
                    Directory.Delete(directoryPath, true);
                    return true;
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    Debug.WriteLine($"Попытка {attempt}/{maxAttempts} не удалась: {ex.Message}");

                    if (attempt == maxAttempts)
                    {
                        var blockedFiles = FindLockedFiles(directoryPath);
                        Debug.WriteLine($"Заблокированных файлов: {blockedFiles.Count}");
                        foreach (var file in blockedFiles.Take(10))
                        {
                            Debug.WriteLine($"   - {Path.GetFileName(file)}");
                        }
                    }
                    else
                    {
                        System.Threading.Thread.Sleep(delayMs);
                    }
                }
            }

            return false;
        }
    }
}
