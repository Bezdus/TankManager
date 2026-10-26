using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Накладывает ручные правки операций (добавленные, исключённые, с заданной ценой) на операции детали,
    /// в том числе только что прочитанные из КОМПАС
    /// </summary>
    public static class OperationEditsMerger
    {
        /// <summary>
        /// Ключ детали для правок: одинаковые экземпляры в сборке правятся вместе.
        /// Из пути к файлу берётся только имя: тот же проект у разных сотрудников может лежать
        /// по разным путям (буква диска, сетевой путь), а после переноса папки правки не должны теряться
        /// </summary>
        public static string GetPartKey(PartModel part)
        {
            return $"{part.Name}|{part.Marking}|{FileNameOf(part.FilePath)}";
        }

        /// <summary>
        /// Ключ из файлов прежних версий («Имя|Обозначение|полный путь») в нынешнем виде
        /// </summary>
        public static string NormalizeKey(string key)
        {
            if (string.IsNullOrEmpty(key))
                return key;

            int separator = key.LastIndexOf('|');
            return separator < 0 ? key : key.Substring(0, separator + 1) + FileNameOf(key.Substring(separator + 1));
        }

        // Без Path.GetFileName: путь из чужого файла может содержать недопустимые символы
        private static string FileNameOf(string path)
        {
            if (string.IsNullOrEmpty(path))
                return string.Empty;

            int slash = path.LastIndexOfAny(new[] { '\\', '/' });
            return slash < 0 ? path : path.Substring(slash + 1);
        }

        /// <summary>
        /// Заменить правки детали: сбросить имеющиеся (ручные операции убрать, у операций из КОМПАС
        /// снять признаки правок) и наложить edited. Пустой edited — просто сброс правок.
        /// Возвращает число правок операций из КОМПАС, для которых не нашлось пары
        /// </summary>
        public static int ReplaceEdits(PartModel target, IEnumerable<ManufacturingOperationBase> edited)
        {
            foreach (var manual in target.Operations.Where(op => op.Origin == OperationOrigin.Manual).ToList())
                target.Operations.Remove(manual);

            foreach (var op in target.Operations)
            {
                op.Quantity = 1;
                op.IsExcluded = false;
                op.HasManualCost = false;
                op.ManualCost = 0;
                op.IsEdited = false;
                op.ClearChange();
            }

            return MergeOperations(target, edited ?? Enumerable.Empty<ManufacturingOperationBase>());
        }

        // Не входят в сравнение: стоимость и производные значения пересчитываются, подпись — не правка
        private static readonly HashSet<string> IgnoredProperties = new HashSet<string>
        {
            nameof(ManufacturingOperationBase.Cost),
            nameof(ManufacturingOperationBase.TimeMinutes),
            nameof(ManufacturingOperationBase.MaterialThickness),
            nameof(RollingOperation.PartMass)
        };

        /// <summary>
        /// Подписывает операции, изменённые в окне операций: добавленные, исключённые, изменённые.
        /// Операции сопоставляются по <see cref="ManufacturingOperationBase.InstanceId"/>. У операций,
        /// вернувшихся к значениям из КОМПАС, подпись снимается; у неизменённых остаётся прежняя
        /// </summary>
        public static void StampChanges(IEnumerable<ManufacturingOperationBase> before,
            IEnumerable<ManufacturingOperationBase> after, string login, string name, long utcTicks)
        {
            var previous = (before ?? Enumerable.Empty<ManufacturingOperationBase>())
                .GroupBy(op => op.InstanceId)
                .ToDictionary(g => g.Key, g => g.First());

            foreach (var op in after ?? Enumerable.Empty<ManufacturingOperationBase>())
            {
                if (!op.IsModified)
                {
                    op.ClearChange();
                    continue;
                }

                if (!previous.TryGetValue(op.InstanceId, out var was))
                {
                    op.SetChange(op.IsManual ? OperationChangeKind.Added : OperationChangeKind.Edited, login, name, utcTicks);
                    continue;
                }

                if (op.IsExcluded && !was.IsExcluded)
                    op.SetChange(OperationChangeKind.Excluded, login, name, utcTicks);
                else if (Signature(op) != Signature(was))
                    op.SetChange(op.IsExcluded ? OperationChangeKind.Excluded : OperationChangeKind.Edited, login, name, utcTicks);
            }
        }

        /// <summary>
        /// Значения, которые правит пользователь (параметры, количество, цена, признаки правок)
        /// </summary>
        private static string Signature(ManufacturingOperationBase op)
        {
            var values = op.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.CanWrite && p.GetSetMethod() != null && !IgnoredProperties.Contains(p.Name))
                .Where(p => p.PropertyType.IsPrimitive || p.PropertyType.IsEnum || p.PropertyType == typeof(string))
                .OrderBy(p => p.Name)
                .Select(p => p.Name + "=" + Convert.ToString(p.GetValue(op), CultureInfo.InvariantCulture));

            return string.Join(";", values);
        }

        /// <summary>
        /// Возвращает число правок операций из КОМПАС, для которых не нашлось пары
        /// </summary>
        private static int MergeOperations(PartModel target, IEnumerable<ManufacturingOperationBase> sourceOperations)
        {
            var source = sourceOperations.ToList();

            // Операции из КОМПАС сопоставляем по типу и порядковому номеру среди операций этого типа
            var freshByType = target.Operations
                .Where(op => op.Origin == OperationOrigin.Kompas)
                .GroupBy(op => op.Type)
                .ToDictionary(g => g.Key, g => g.ToList());

            int lostEdits = 0;

            foreach (var group in source.Where(op => op.Origin == OperationOrigin.Kompas).GroupBy(op => op.Type))
            {
                freshByType.TryGetValue(group.Key, out var candidates);

                int index = 0;
                foreach (var old in group)
                {
                    var match = candidates != null && index < candidates.Count ? candidates[index] : null;
                    index++;

                    if (!old.IsModified)
                        continue;

                    if (match == null)
                    {
                        lostEdits++;
                        continue;
                    }

                    if (old.IsEdited)
                        match.CopyParametersFrom(old);

                    match.Quantity = old.Quantity;
                    match.IsExcluded = old.IsExcluded;
                    match.HasManualCost = old.HasManualCost;
                    match.ManualCost = old.ManualCost;
                    match.IsEdited = old.IsEdited;
                    match.CopyChangeFrom(old);
                }
            }

            foreach (var manual in source.Where(op => op.Origin == OperationOrigin.Manual))
                target.Operations.Add(manual.Clone());

            target.RecalculateOperationsCost();
            return lostEdits;
        }
    }
}
