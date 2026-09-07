using System;
using System.Collections.Generic;
using System.IO;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Поиск DXF-файлов для лазерной резки относительно файла сборки.
    /// Папка Dxf ищется либо рядом со сборкой, либо на один уровень выше.
    /// Сопоставление детали с файлом — по обозначению (Marking) в имени файла.
    /// </summary>
    public static class DxfResolver
    {
        private const string DxfFolderName = "Dxf";

        /// <summary>
        /// Возвращает упорядоченный список папок-кандидатов (ближняя приоритетнее)
        /// </summary>
        public static IReadOnlyList<string> GetCandidateFolders(string assemblyFilePath)
        {
            var folders = new List<string>();

            if (string.IsNullOrEmpty(assemblyFilePath))
                return folders;

            string assemblyDir = Path.GetDirectoryName(assemblyFilePath);
            if (string.IsNullOrEmpty(assemblyDir))
                return folders;

            folders.Add(Path.Combine(assemblyDir, DxfFolderName));

            var parent = Directory.GetParent(assemblyDir);
            if (parent != null)
            {
                folders.Add(Path.Combine(parent.FullName, DxfFolderName));
            }

            return folders;
        }

        /// <summary>
        /// Находит DXF-файл по обозначению: имя файла должно содержать обозначение (без учёта регистра).
        /// Папки перебираются в порядке приоритета.
        /// </summary>
        public static string FindDxfForMarking(string marking, IEnumerable<string> folders)
        {
            if (string.IsNullOrEmpty(marking))
                return null;

            foreach (var folder in folders)
            {
                if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
                    continue;

                IEnumerable<string> files;
                try
                {
                    files = Directory.EnumerateFiles(folder, "*.dxf", SearchOption.TopDirectoryOnly);
                }
                catch
                {
                    continue;
                }

                foreach (var file in files)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    if (fileName.IndexOf(marking, StringComparison.OrdinalIgnoreCase) >= 0)
                        return file;
                }
            }

            return null;
        }
    }
}
