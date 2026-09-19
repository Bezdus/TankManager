using System.IO;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Атомарная запись файла: данные пишутся во временный файл, который затем подменяет целевой.
    /// При сбое на середине записи исходный файл остаётся нетронутым.
    /// </summary>
    public static class AtomicFile
    {
        public static void WriteAllBytes(string path, byte[] data)
        {
            string tempPath = path + ".tmp";
            File.WriteAllBytes(tempPath, data);

            if (File.Exists(path))
                File.Replace(tempPath, path, null);
            else
                File.Move(tempPath, path);
        }
    }
}
