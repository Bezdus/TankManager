using System;
using System.IO;

namespace TankManager.Core.Services
{
    public class FileLogger : ILogger
    {
        private readonly string _logFilePath;
        private static readonly object _lockObject = new object();

        // Лог не растёт бесконечно: больше 5 МБ — переименовывается в .old (прежний .old удаляется)
        private const long MaxLogSize = 5 * 1024 * 1024;
        private const int SizeCheckInterval = 200;
        private static int _writesSinceSizeCheck = SizeCheckInterval;

        public FileLogger(string logFilePath = null)
        {
            _logFilePath = logFilePath ?? GetDefaultLogPath();
        }

        private static string GetDefaultLogPath()
        {
            string directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "TankManager");

            try { Directory.CreateDirectory(directory); }
            catch { }

            return Path.Combine(directory, "TankManager.log");
        }

        public void LogInfo(string message) => WriteLog("INFO", message);

        public void LogWarning(string message) => WriteLog("WARNING", message);

        public void LogError(string message, Exception exception = null)
        {
            var fullMessage = exception != null
                ? $"{message}\n{exception}"
                : message;
            WriteLog("ERROR", fullMessage);
        }

        private void WriteLog(string level, string message)
        {
            lock (_lockObject)
            {
                try
                {
                    RotateIfNeeded();
                }
                catch
                {
                    // Не удалось переименовать (файл открыт другим процессом) — пишем в прежний
                }

                try
                {
                    var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
                    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
                }
                catch
                {
                    // Игнорируем ошибки логирования
                }
            }
        }

        private void RotateIfNeeded()
        {
            // Размер проверяется не на каждой записи: лог пишется часто
            if (++_writesSinceSizeCheck < SizeCheckInterval)
                return;
            _writesSinceSizeCheck = 0;

            var info = new FileInfo(_logFilePath);
            if (!info.Exists || info.Length < MaxLogSize)
                return;

            string oldPath = _logFilePath + ".old";
            if (File.Exists(oldPath))
                File.Delete(oldPath);
            File.Move(_logFilePath, oldPath);
        }
    }
}