using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Вид изменения в журнале
    /// </summary>
    public enum AuditAction
    {
        ProductSaved,
        ProductDeleted,
        OperationsEdited,
        PricingChanged,
        UsersChanged,
        UserRegistered
    }

    /// <summary>
    /// Запись журнала изменений: кто, когда и что изменил
    /// </summary>
    [DataContract]
    public class AuditEntry
    {
        [DataMember]
        public long UtcTicks { get; set; }

        [DataMember]
        public string Login { get; set; }

        [DataMember]
        public string Name { get; set; }

        [DataMember(Name = "Action")]
        private string ActionText
        {
            get => Action.ToString();
            set => Action = Enum.TryParse(value, true, out AuditAction action) ? action : AuditAction.ProductSaved;
        }

        public AuditAction Action { get; set; }

        [DataMember(EmitDefaultValue = false)]
        public string ProductName { get; set; }

        [DataMember(EmitDefaultValue = false)]
        public string Marking { get; set; }

        [DataMember(EmitDefaultValue = false)]
        public string PartName { get; set; }

        [DataMember(EmitDefaultValue = false)]
        public string PartMarking { get; set; }

        /// <summary>
        /// Что изменилось («было → стало»), по строке на изменение
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string Details { get; set; }

        public DateTime LocalTime => new DateTime(UtcTicks, DateTimeKind.Utc).ToLocalTime();

        public string UserDisplayName => string.IsNullOrWhiteSpace(Name) ? Login : Name;

        public string ActionTitle => GetActionTitle(Action);

        public string ProductDisplay => string.IsNullOrEmpty(Marking) ? ProductName : $"{ProductName} ({Marking})";

        public string PartDisplay => string.IsNullOrEmpty(PartMarking) ? PartName : $"{PartName} ({PartMarking})";

        public static string GetActionTitle(AuditAction action)
        {
            switch (action)
            {
                case AuditAction.ProductSaved: return "Сохранение изделия";
                case AuditAction.ProductDeleted: return "Удаление изделия";
                case AuditAction.OperationsEdited: return "Правка операций";
                case AuditAction.PricingChanged: return "Изменение расценок";
                case AuditAction.UsersChanged: return "Изменение сотрудников";
                case AuditAction.UserRegistered: return "Новый сотрудник";
                default: return action.ToString();
            }
        }
    }

    /// <summary>
    /// Журнал изменений. У каждого сотрудника свой файл на месяц (&lt;логин&gt;_&lt;гггг-ММ&gt;.jsonl, строка — запись),
    /// поэтому в один файл одновременно никто не пишет. Запись дописывается в локальную папку
    /// products\_audit и копируется в &lt;серверная папка&gt;\_audit; без сервера — при следующей синхронизации
    /// </summary>
    public class AuditService
    {
        private const string AuditFolderName = "_audit";
        private const string FileExtension = ".jsonl";

        private static readonly string LocalFolder =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "products", AuditFolderName);

        private static readonly object FileLock = new object();
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        private readonly ILogger _logger = new FileLogger();
        private readonly Func<string> _serverFolderProvider;

        /// <param name="serverFolderProvider">Серверная папка изделий (null — не задана)</param>
        public AuditService(Func<string> serverFolderProvider)
        {
            _serverFolderProvider = serverFolderProvider;
        }

        private string ServerAuditFolder
        {
            get
            {
                string server = _serverFolderProvider?.Invoke();
                return string.IsNullOrEmpty(server) ? null : Path.Combine(server, AuditFolderName);
            }
        }

        private static string OwnFilePrefix => SafeFileName(CurrentUser.Login) + "_";

        /// <summary>
        /// Записывает изменение от имени текущего сотрудника. Ошибки журнала не мешают работе: только логируются
        /// </summary>
        public void Record(AuditAction action, string productName = null, string marking = null,
            string partName = null, string partMarking = null, string details = null)
        {
            var entry = new AuditEntry
            {
                UtcTicks = DateTime.UtcNow.Ticks,
                Login = CurrentUser.Login,
                Name = CurrentUser.Name,
                Action = action,
                ProductName = productName,
                Marking = marking,
                PartName = partName,
                PartMarking = partMarking,
                Details = string.IsNullOrWhiteSpace(details) ? null : details.Trim()
            };

            try
            {
                string line = Serialize(entry);
                string fileName = OwnFilePrefix + DateTime.UtcNow.ToString("yyyy-MM") + FileExtension;

                lock (FileLock)
                {
                    Directory.CreateDirectory(LocalFolder);
                    File.AppendAllText(Path.Combine(LocalFolder, fileName), line + "\n", Utf8NoBom);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Не удалось записать в журнал изменений: {action}", ex);
                return;
            }

            // Сетевая папка может отвечать медленно — копируем в фоне
            Task.Run(() => UploadPending());
        }

        /// <summary>
        /// Копирует свои файлы журнала на сервер, если там они короче (файлы только дописываются)
        /// </summary>
        public void UploadPending()
        {
            string serverFolder = ServerAuditFolder;
            if (serverFolder == null)
                return;

            lock (FileLock)
            {
                try
                {
                    if (!Directory.Exists(LocalFolder) || !Directory.Exists(Path.GetDirectoryName(serverFolder)))
                        return;

                    foreach (var localFile in Directory.GetFiles(LocalFolder, OwnFilePrefix + "*" + FileExtension))
                    {
                        string serverFile = Path.Combine(serverFolder, Path.GetFileName(localFile));
                        var localInfo = new FileInfo(localFile);
                        var serverInfo = new FileInfo(serverFile);
                        if (serverInfo.Exists && serverInfo.Length >= localInfo.Length)
                            continue;

                        Directory.CreateDirectory(serverFolder);
                        byte[] data = File.ReadAllBytes(localFile);
                        AtomicFile.WriteAllBytes(serverFile, data);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Не удалось выложить журнал изменений на сервер: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Все записи журнала (с сервера и свои локальные), новые сверху
        /// </summary>
        public List<AuditEntry> ReadAll()
        {
            var result = new List<AuditEntry>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string serverFolder = ServerAuditFolder;
            if (serverFolder != null)
            {
                try
                {
                    if (Directory.Exists(serverFolder))
                    {
                        foreach (var file in Directory.GetFiles(serverFolder, "*" + FileExtension))
                            ReadFile(file, result, seen);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Не удалось прочитать журнал изменений на сервере: {ex.Message}");
                }
            }

            // Свои записи, ещё не попавшие на сервер
            try
            {
                if (Directory.Exists(LocalFolder))
                {
                    foreach (var file in Directory.GetFiles(LocalFolder, "*" + FileExtension))
                        ReadFile(file, result, seen);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось прочитать локальный журнал изменений: {ex.Message}");
            }

            return result.OrderByDescending(e => e.UtcTicks).ToList();
        }

        private void ReadFile(string path, List<AuditEntry> result, HashSet<string> seen)
        {
            try
            {
                string[] lines;
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                    lines = reader.ReadToEnd().Split('\n');

                foreach (var raw in lines)
                {
                    string line = raw.Trim();
                    if (line.Length == 0)
                        continue;

                    try
                    {
                        var entry = Deserialize(line);
                        if (entry == null)
                            continue;

                        string key = entry.Login + "|" + entry.UtcTicks + "|" + entry.Action;
                        if (seen.Add(key))
                            result.Add(entry);
                    }
                    catch (Exception)
                    {
                        // Недописанная строка (файл копировался во время записи) — пропускаем
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось прочитать журнал {Path.GetFileName(path)}: {ex.Message}");
            }
        }

        private static string Serialize(AuditEntry entry)
        {
            var serializer = new DataContractJsonSerializer(typeof(AuditEntry));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, entry);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        private static AuditEntry Deserialize(string line)
        {
            var serializer = new DataContractJsonSerializer(typeof(AuditEntry));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(line)))
                return (AuditEntry)serializer.ReadObject(stream);
        }

        private static string SafeFileName(string name)
        {
            return string.Join("_", (name ?? "unknown").Split(Path.GetInvalidFileNameChars()));
        }
    }
}
