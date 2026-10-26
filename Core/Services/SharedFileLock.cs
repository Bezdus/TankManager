using System;
using System.IO;
using System.Threading;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Блокировка общего файла между компьютерами на время «прочитать — изменить — записать».
    /// Рядом с файлом открывается &lt;файл&gt;.lock без совместного доступа; блокировку снимает закрытие
    /// дескриптора (в том числе при падении программы или обрыве сети), поэтому «зависших» блокировок нет.
    /// </summary>
    public sealed class SharedFileLock : IDisposable
    {
        public const string Extension = ".lock";

        private static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(10);
        private static readonly Random Jitter = new Random();

        private FileStream _stream;

        private SharedFileLock(FileStream stream)
        {
            _stream = stream;
        }

        /// <summary>
        /// Захватывает блокировку файла; если её держит другой компьютер дольше timeout — IOException
        /// </summary>
        public static SharedFileLock Acquire(string targetPath, TimeSpan? timeout = null)
        {
            string lockPath = targetPath + Extension;
            var deadline = DateTime.UtcNow + (timeout ?? DefaultTimeout);

            while (true)
            {
                try
                {
                    var stream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                        FileShare.None, 1, FileOptions.DeleteOnClose);
                    return new SharedFileLock(stream);
                }
                // Занят другим (или ещё удаляется после закрытия) — ждём
                catch (Exception ex) when ((ex is IOException || ex is UnauthorizedAccessException) && DateTime.UtcNow < deadline)
                {
                    int delay;
                    lock (Jitter)
                        delay = 100 + Jitter.Next(150);
                    Thread.Sleep(delay);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    throw new IOException($"Файл {Path.GetFileName(targetPath)} занят другим пользователем, повторите позже", ex);
                }
            }
        }

        public void Dispose()
        {
            _stream?.Dispose();
            _stream = null;
        }
    }
}
