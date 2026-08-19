using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Определяет путь к установленному KOMPAS-3D без хардкода версии.
    /// Используется для поиска библиотеки импорта/экспорта DXF/DWG.
    /// Работает и из 32-битного процесса.
    /// </summary>
    public static class KompasPathResolver
    {
        /// <summary>
        /// Библиотека импорта/экспорта DXF/DWG относительно корня установки KOMPAS
        /// </summary>
        private const string ConverterRelativePath = @"Libs\ImpExp\dwgdxfImp.rtw";

        /// <summary>
        /// Имя исполняемого файла KOMPAS
        /// </summary>
        private const string KompasExecutable = "KOMPAS.exe";

        /// <summary>
        /// Возвращает полный путь к dwgdxfImp.rtw или null, если не найден
        /// </summary>
        public static string ResolveConverterLibrary()
        {
            string root =
                ResolveFromAppPaths() ??
                ResolveFromRunningProcess() ??
                ResolveFromRegistry() ??
                ResolveFromCommonFolders();

            if (string.IsNullOrEmpty(root))
                return null;

            string lib = Path.Combine(root, ConverterRelativePath);
            return File.Exists(lib) ? lib : null;
        }

        /// <summary>
        /// Определяет корень установки по ключу реестра App Paths (не зависит от разрядности процесса)
        /// </summary>
        private static string ResolveFromAppPaths()
        {
            string[] regPaths =
            {
                $@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{KompasExecutable}",
                $@"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\{KompasExecutable}"
            };

            foreach (var regPath in regPaths)
            {
                try
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(regPath))
                    {
                        string path = key?.GetValue("")?.ToString();
                        if (!string.IsNullOrEmpty(path) && File.Exists(path))
                        {
                            string root = FindRootByWalkingUp(Path.GetDirectoryName(path));
                            if (root != null)
                                return root;
                        }
                    }
                }
                catch { }
            }

            return null;
        }

        /// <summary>
        /// Определяет корень установки по пути запущенного процесса KOMPAS
        /// </summary>
        private static string ResolveFromRunningProcess()
        {
            try
            {
                var proc = Process.GetProcessesByName("KOMPAS")
                    .FirstOrDefault(p =>
                    {
                        try { return !string.IsNullOrEmpty(p.MainModule?.FileName); }
                        catch { return false; }
                    });

                if (proc == null)
                    return null;

                return FindRootByWalkingUp(Path.GetDirectoryName(proc.MainModule.FileName));
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Определяет корень установки по реестру ASCON.
        /// Имя значения с путём различается между версиями — перебираем кандидатов.
        /// </summary>
        private static string ResolveFromRegistry()
        {
            string[] valueNames = { "CurrentExe", "RootPath", "InstallPath", "SetupPath", "Path", "InstallDir", "BinDir" };

            try
            {
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                    using (var ascon = baseKey.OpenSubKey(@"SOFTWARE\ASCON"))
                    {
                        if (ascon == null)
                            continue;

                        foreach (var subName in ascon.GetSubKeyNames())
                        {
                            using (var key = ascon.OpenSubKey(subName))
                            {
                                if (key == null)
                                    continue;

                                foreach (var v in valueNames)
                                {
                                    var val = key.GetValue(v) as string;
                                    if (string.IsNullOrEmpty(val) || !Path.IsPathRooted(val))
                                        continue;

                                    string startDir = File.Exists(val) ? Path.GetDirectoryName(val) : val;
                                    string root = FindRootByWalkingUp(startDir);
                                    if (root != null)
                                        return root;
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Ищет установку в типовых папках Program Files\ASCON.
        /// Учитывает, что в 32-битном процессе ProgramFiles перенаправляется на x86.
        /// </summary>
        private static string ResolveFromCommonFolders()
        {
            var roots = new List<string>
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                GetProgramFilesDirectory64()
            };

            foreach (var root in roots.Where(r => !string.IsNullOrEmpty(r)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    string ascon = Path.Combine(root, "ASCON");
                    if (!Directory.Exists(ascon))
                        continue;

                    foreach (var dir in Directory.GetDirectories(ascon, "KOMPAS-3D*"))
                    {
                        if (File.Exists(Path.Combine(dir, ConverterRelativePath)))
                            return dir;
                    }
                }
                catch { }
            }

            return null;
        }

        /// <summary>
        /// Возвращает настоящий 64-битный Program Files (не перенаправляется в x86 в 32-битном процессе)
        /// </summary>
        private static string GetProgramFilesDirectory64()
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion"))
                {
                    return key?.GetValue("ProgramFilesDir") as string;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Идёт вверх от каталога (обычно <root>\Bin или <root>\Bin64),
        /// пока не найдёт каталог с Libs\ImpExp\dwgdxfImp.rtw
        /// </summary>
        private static string FindRootByWalkingUp(string startDir)
        {
            if (string.IsNullOrEmpty(startDir))
                return null;

            try
            {
                var dir = new DirectoryInfo(startDir);
                while (dir != null)
                {
                    if (File.Exists(Path.Combine(dir.FullName, ConverterRelativePath)))
                        return dir.FullName;
                    dir = dir.Parent;
                }
            }
            catch { }

            return null;
        }
    }
}
