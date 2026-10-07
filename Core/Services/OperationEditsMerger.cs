using System.Collections.Generic;
using System.Linq;
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
        /// Ключ детали для правок: одинаковые экземпляры в сборке правятся вместе
        /// </summary>
        public static string GetPartKey(PartModel part)
        {
            return $"{part.Name}|{part.Marking}|{part.FilePath}";
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
            }

            return MergeOperations(target, edited ?? Enumerable.Empty<ManufacturingOperationBase>());
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
                }
            }

            foreach (var manual in source.Where(op => op.Origin == OperationOrigin.Manual))
                target.Operations.Add(manual.Clone());

            target.RecalculateOperationsCost();
            return lostEdits;
        }
    }
}
