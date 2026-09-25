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
        /// Находит DXF-файл по обозначению. Обозначение должно входить в имя файла как отдельный токен:
        /// по краям — начало/конец имени или разделитель (не буква, не цифра, не '.' и не '-'),
        /// чтобы «АБ.01.001» не находило «АБ.01.0010» или «АБ.01.001.1».
        /// Точное совпадение имени приоритетнее, затем самое короткое имя. Папки — в порядке приоритета.
        /// </summary>
        public static string FindDxfForMarking(string marking, IEnumerable<string> folders)
        {
            if (string.IsNullOrEmpty(marking))
                return null;

            foreach (var folder in folders)
            {
                if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
                    continue;

                string[] files;
                try
                {
                    files = Directory.GetFiles(folder, "*.dxf", SearchOption.TopDirectoryOnly);
                }
                catch (Exception)
                {
                    continue;
                }

                Array.Sort(files, StringComparer.OrdinalIgnoreCase);

                string best = null;
                int bestLength = int.MaxValue;

                foreach (var file in files)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);

                    if (string.Equals(fileName, marking, StringComparison.OrdinalIgnoreCase))
                        return file;

                    if (ContainsToken(fileName, marking) && fileName.Length < bestLength)
                    {
                        best = file;
                        bestLength = fileName.Length;
                    }
                }

                if (best != null)
                    return best;
            }

            return null;
        }

        private static bool ContainsToken(string fileName, string marking)
        {
            int start = 0;
            while (true)
            {
                int index = fileName.IndexOf(marking, start, StringComparison.OrdinalIgnoreCase);
                if (index < 0)
                    return false;

                int end = index + marking.Length;
                bool leftOk = index == 0 || IsBoundary(fileName[index - 1]);
                bool rightOk = end >= fileName.Length || IsBoundary(fileName[end]);
                if (leftOk && rightOk)
                    return true;

                start = index + 1;
            }
        }

        private static bool IsBoundary(char c)
        {
            return !char.IsLetterOrDigit(c) && c != '.' && c != '-';
        }
    }
}
