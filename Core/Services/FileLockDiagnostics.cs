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
        /// Пытается удалить папку с диагностикой блокированных файлов
        /// </summary>
        public static bool TryDeleteDirectoryWithDiagnostics(string directoryPath, out List<string> blockedFiles)
        {
            blockedFiles = new List<string>();

            if (!Directory.Exists(directoryPath))
                return true;

            // Сначала проверяем блокировки
            blockedFiles = FindLockedFiles(directoryPath);

            if (blockedFiles.Count > 0)
            {
                Debug.WriteLine($"? Не удалось удалить папку. Заблокировано файлов: {blockedFiles.Count}");
                foreach (var file in blockedFiles)
                {
                    Debug.WriteLine($"   - {Path.GetFileName(file)}");
                }
                return false;
            }

            // Пытаемся удалить
            try
            {
                Directory.Delete(directoryPath, true);
                Debug.WriteLine($"? Папка успешно удалена: {directoryPath}");
                return true;
            }
            catch (IOException ex)
            {
                Debug.WriteLine($"? Ошибка удаления папки: {ex.Message}");
                // Повторно проверяем, что именно заблокировано
                blockedFiles = FindLockedFiles(directoryPath);
                return false;
            }
        }

        /// <summary>
        /// Принудительная очистка с ожиданием и повторными попытками
        /// </summary>
        public static bool ForceDeleteDirectory(string directoryPath, int maxAttempts = 5, int delayMs = 500)
        {
            if (!Directory.Exists(directoryPath))
                return true;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                Debug.WriteLine($"Попытка {attempt}/{maxAttempts} удаления папки...");

                // Принудительная сборка мусора
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Ждём освобождения ресурсов
                System.Threading.Thread.Sleep(delayMs);

                // Снимаем атрибуты только для чтения
                try
                {
                    foreach (var file in Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories))
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                    }
                }
                catch { }

                // Пытаемся удалить
                try
                {
                    Directory.Delete(directoryPath, true);
                    Debug.WriteLine($"? Папка успешно удалена на попытке {attempt}");
                    return true;
                }
                catch (IOException ex)
                {
                    Debug.WriteLine($"? Попытка {attempt} не удалась: {ex.Message}");

                    // На последней попытке выводим детальную диагностику
                    if (attempt == maxAttempts)
                    {
                        var blockedFiles = FindLockedFiles(directoryPath);
                        Debug.WriteLine($"Заблокировано файлов: {blockedFiles.Count}");
                        foreach (var file in blockedFiles.Take(10))
                        {
                            Debug.WriteLine($"   - {Path.GetFileName(file)}");
                        }
                    }
                }
            }

            return false;
        }
    }
}
