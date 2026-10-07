using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Ручные правки операций хранятся отдельно от product.json — в operations.json папки изделия.
    /// Их пишут конструктор и технолог (у технолога product.json только для чтения), поэтому
    /// пересохранение изделия конструктором правки не затирает. Файл сливается по деталям:
    /// для каждой детали побеждает более поздняя правка.
    /// </summary>
    public partial class ProductStorageService
    {
        private const string OperationEditsFileName = "operations.json";

        /// <summary>
        /// Сохраняет правки операций детали локально и на сервер. Ошибка локальной записи пробрасывается,
        /// ошибка сервера возвращается текстом (null — всё записано)
        /// </summary>
        public string SaveOperationEdits(Product product, PartModel part, IEnumerable<ManufacturingOperationBase> operations,
            string removedNote = null)
        {
            if (product == null || part == null || string.IsNullOrEmpty(product.Name))
                return null;

            var ops = operations?.ToList() ?? new List<ManufacturingOperationBase>();

            // Храним все операции детали: операции из КОМПАС сопоставляются по порядковому номеру.
            // Без правок — пустой список: он сбрасывает прежние правки детали
            var entry = new PartOperationEditsDto
            {
                Key = OperationEditsMerger.GetPartKey(part),
                PartName = part.Name,
                Marking = part.Marking,
                ModifiedUtcTicks = DateTime.UtcNow.Ticks,
                ModifiedBy = CurrentUser.Login,
                ModifiedByName = CurrentUser.Name,
                RemovedNote = removedNote,
                Operations = ops.Any(op => op.IsModified)
                    ? ops.Select(ToOperationDto).ToList()
                    : new List<OperationDto>()
            };

            lock (_ioLock)
            {
                string localFolder = FindExistingProductFolder(product, ProductsDirectory)
                                     ?? Path.Combine(ProductsDirectory, GetProductFolderName(product));
                Directory.CreateDirectory(localFolder);

                var local = ReadOperationEdits(localFolder, out bool _);
                local.RemoveAll(e => string.Equals(e.Key, entry.Key, StringComparison.OrdinalIgnoreCase));
                local.Add(entry);
                WriteOperationEdits(localFolder, local);

                if (!HasServerFolder)
                    return null;

                if (!IsServerAvailable)
                    return "серверная папка недоступна, правки будут отправлены при синхронизации";

                string serverFolder = FindExistingProductFolder(product, _serverStorageFolder);
                if (serverFolder == null)
                    return "изделие ещё не сохранено на сервере, правки будут отправлены при его сохранении";

                return SyncOperationEditsFolder(localFolder, serverFolder, upload: true);
            }
        }

        /// <summary>
        /// Накладывает сохранённые правки операций на изделие (например, только что прочитанное из КОМПАС)
        /// </summary>
        public void ApplyOperationEdits(Product product)
        {
            if (product == null || string.IsNullOrEmpty(product.Name))
                return;

            string folder = FindExistingProductFolder(product, ProductsDirectory);
            if (folder != null)
                ApplyOperationEdits(product, folder);
        }

        private void ApplyOperationEdits(Product product, string productFolder)
        {
            try
            {
                var edits = ReadOperationEdits(productFolder, out bool _);
                if (edits.Count == 0 || product.Details == null)
                    return;

                var byKey = new Dictionary<string, PartOperationEditsDto>(StringComparer.OrdinalIgnoreCase);
                foreach (var entry in edits.Where(e => !string.IsNullOrEmpty(e.Key)))
                    byKey[entry.Key] = entry;

                var parts = product.Details
                    .Concat(product.StandardParts ?? Enumerable.Empty<PartModel>())
                    .Distinct();
                var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var part in parts)
                {
                    string key = OperationEditsMerger.GetPartKey(part);
                    if (!byKey.TryGetValue(key, out var entry))
                        continue;

                    var operations = (entry.Operations ?? new List<OperationDto>())
                        .Select(FromOperationDto)
                        .Where(op => op != null)
                        .ToList();

                    int lost = OperationEditsMerger.ReplaceEdits(part, operations);
                    part.SetOperationsModified(
                        string.IsNullOrWhiteSpace(entry.ModifiedByName) ? entry.ModifiedBy : entry.ModifiedByName,
                        entry.ModifiedUtcTicks > 0 ? new DateTime(entry.ModifiedUtcTicks, DateTimeKind.Utc) : (DateTime?)null,
                        entry.RemovedNote);
                    if (lost > 0 && reported.Add(key))
                        _logger.LogWarning($"Правки операций детали {part.Name} {part.Marking}: {lost} шт. не применены — операции больше нет в модели КОМПАС");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка применения правок операций: {ex.Message}");
            }
        }

        /// <summary>
        /// Сливает правки операций всех изделий, которые есть и локально, и на сервере
        /// </summary>
        private void SyncAllOperationEdits(HashSet<string> deleted)
        {
            bool upload = AppMode.CanEditOperations;

            foreach (var localFolder in Directory.GetDirectories(ProductsDirectory))
            {
                string folderName = Path.GetFileName(localFolder);
                if (folderName.StartsWith("_") || deleted.Contains(folderName))
                    continue;

                string serverFolder = Path.Combine(_serverStorageFolder, folderName);
                if (!Directory.Exists(serverFolder))
                    continue;

                string error = SyncOperationEditsFolder(localFolder, serverFolder, upload);
                if (error != null)
                    _logger.LogWarning($"Правки операций {folderName}: {error}");
            }
        }

        /// <summary>
        /// Сливает operations.json локальной и серверной папок изделия: по каждой детали побеждает
        /// более поздняя правка. Возвращает текст ошибки сервера или null
        /// </summary>
        private string SyncOperationEditsFolder(string localFolder, string serverFolder, bool upload)
        {
            try
            {
                var local = ReadOperationEdits(localFolder, out bool localOk);
                var server = ReadOperationEdits(serverFolder, out bool serverOk);

                var merged = new Dictionary<string, PartOperationEditsDto>(StringComparer.OrdinalIgnoreCase);
                foreach (var entry in local.Concat(server).Where(e => !string.IsNullOrEmpty(e.Key)))
                {
                    if (!merged.TryGetValue(entry.Key, out var existing) || entry.ModifiedUtcTicks > existing.ModifiedUtcTicks)
                        merged[entry.Key] = entry;
                }

                var result = merged.Values.OrderBy(e => e.Key, StringComparer.OrdinalIgnoreCase).ToList();

                if (!localOk || !SameEdits(local, result))
                    WriteOperationEdits(localFolder, result);

                if (upload && !SameEdits(server, result))
                {
                    // Не перезаписываем файл на сервере, который не удалось прочитать (иначе потеряем чужие правки)
                    if (!serverOk)
                        return "не удалось прочитать правки операций на сервере";

                    WriteOperationEdits(serverFolder, result);
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Ошибка синхронизации правок операций {Path.GetFileName(localFolder)}: {ex.Message}");
                return ex.Message;
            }
        }

        private static bool SameEdits(List<PartOperationEditsDto> a, List<PartOperationEditsDto> b)
        {
            if (a.Count != b.Count)
                return false;

            var ticks = a.Where(e => e.Key != null)
                .GroupBy(e => e.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.Max(e => e.ModifiedUtcTicks), StringComparer.OrdinalIgnoreCase);

            return b.All(e => e.Key != null && ticks.TryGetValue(e.Key, out long t) && t == e.ModifiedUtcTicks);
        }

        /// <summary>
        /// Читает правки операций папки изделия. ok = false, если файл есть, но прочитать его не удалось
        /// </summary>
        private List<PartOperationEditsDto> ReadOperationEdits(string productFolder, out bool ok)
        {
            ok = true;
            string path = Path.Combine(productFolder, OperationEditsFileName);

            try
            {
                if (!File.Exists(path))
                    return new List<PartOperationEditsDto>();

                var serializer = new DataContractJsonSerializer(typeof(OperationEditsFile));
                using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var file = (OperationEditsFile)serializer.ReadObject(fileStream);
                    return file?.Parts ?? new List<PartOperationEditsDto>();
                }
            }
            catch (Exception ex)
            {
                ok = false;
                _logger.LogWarning($"Не удалось прочитать правки операций {path}: {ex.Message}");
                return new List<PartOperationEditsDto>();
            }
        }

        private static void WriteOperationEdits(string productFolder, List<PartOperationEditsDto> parts)
        {
            var serializer = new DataContractJsonSerializer(typeof(OperationEditsFile));
            using (var memoryStream = new MemoryStream())
            {
                serializer.WriteObject(memoryStream, new OperationEditsFile { Parts = parts });
                AtomicFile.WriteAllBytes(Path.Combine(productFolder, OperationEditsFileName), memoryStream.ToArray());
            }
        }
    }

    [DataContract]
    public class OperationEditsFile
    {
        [DataMember]
        public List<PartOperationEditsDto> Parts { get; set; }
    }

    /// <summary>
    /// Правки операций одной детали (общие для всех её экземпляров в сборке)
    /// </summary>
    [DataContract]
    public class PartOperationEditsDto
    {
        /// <summary>
        /// Ключ детали: <see cref="OperationEditsMerger.GetPartKey"/>
        /// </summary>
        [DataMember]
        public string Key { get; set; }

        [DataMember]
        public string PartName { get; set; }

        [DataMember]
        public string Marking { get; set; }

        [DataMember]
        public long ModifiedUtcTicks { get; set; }

        [DataMember]
        public string ModifiedBy { get; set; }

        [DataMember(EmitDefaultValue = false)]
        public string ModifiedByName { get; set; }

        /// <summary>
        /// «Удалил операцию «Сварка»», если при этой правке операции удаляли (их подписать уже негде)
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string RemovedNote { get; set; }

        /// <summary>
        /// Все операции детали после правки; пустой список — правки сброшены
        /// </summary>
        [DataMember]
        public List<OperationDto> Operations { get; set; }
    }
}
