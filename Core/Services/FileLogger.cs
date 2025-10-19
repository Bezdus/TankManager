using System;
using System.IO;

namespace TankManager.Core.Services
{
    public class FileLogger : ILogger
    {
        private readonly string _logFilePath;
        private readonly object _lockObject = new object();

        public FileLogger(string logFilePath = "TankManager.log")
        {
            _logFilePath = logFilePath;
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
                    var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
                    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
                }
                catch
                {
                    // Игнорируем ошибки логирования
                }
            }
        }
    }
}