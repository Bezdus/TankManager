using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Текст «было → стало» для журнала изменений
    /// </summary>
    public static class ChangeDescriber
    {
        private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

        private static string Num(double value) => value.ToString("0.##", Ru);

        #region Расценки

        public static string DescribePricing(PricingSettings before, PricingSettings after)
        {
            if (after == null) return null;
            before = before ?? new PricingSettings();

            var lines = new List<string>();
            void Compare(string title, double was, double now)
            {
                if (Math.Abs(was - now) > 0.0001)
                    lines.Add($"{title}: {Num(was)} → {Num(now)}");
            }

            Compare("Листовой прокат, руб/кг", before.SheetMetalPricePerKg, after.SheetMetalPricePerKg);
            Compare("Прочий металл, руб/кг", before.OtherMetalPricePerKg, after.OtherMetalPricePerKg);
            Compare("Лазерная резка по умолчанию, руб/м", before.LaserCuttingPricePerMeter, after.LaserCuttingPricePerMeter);
            Compare("Гравировка, руб/мм", before.EngravingPricePerMm, after.EngravingPricePerMm);
            Compare("Гибка, руб/операция", before.BendingPricePerOperation, after.BendingPricePerOperation);
            Compare("Вальцовка, руб/кг", before.RollingPricePerKg, after.RollingPricePerKg);
            Compare("Отбортовка, руб/операция", before.FlangingPricePerOperation, after.FlangingPricePerOperation);

            CompareTable(lines, "Резка, толщина {0} мм, руб/м",
                ToTable(before.LaserCuttingPricing, e => Num(e.Thickness), e => e.PricePerMeter),
                ToTable(after.LaserCuttingPricing, e => Num(e.Thickness), e => e.PricePerMeter));

            CompareTable(lines, "Труба {0}, руб/м",
                ToTable(before.TubularPricing, e => e.Size?.Trim(), e => e.PricePerMeter),
                ToTable(after.TubularPricing, e => e.Size?.Trim(), e => e.PricePerMeter));

            return lines.Count == 0 ? null : string.Join("\n", lines);
        }

        private static Dictionary<string, double> ToTable<T>(IEnumerable<T> items, Func<T, string> key, Func<T, double> price)
        {
            var table = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in items ?? Enumerable.Empty<T>())
            {
                string k = item == null ? null : key(item);
                if (!string.IsNullOrEmpty(k))
                    table[k] = price(item);
            }
            return table;
        }

        private static void CompareTable(List<string> lines, string titleFormat,
            Dictionary<string, double> before, Dictionary<string, double> after)
        {
            foreach (var key in before.Keys.Union(after.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                string title = string.Format(titleFormat, key);
                bool had = before.TryGetValue(key, out double was);
                bool has = after.TryGetValue(key, out double now);

                if (had && !has)
                    lines.Add($"{title}: {Num(was)} → удалено");
                else if (!had && has)
                    lines.Add($"{title}: добавлено {Num(now)}");
                else if (Math.Abs(was - now) > 0.0001)
                    lines.Add($"{title}: {Num(was)} → {Num(now)}");
            }
        }

        #endregion

        #region Операции

        public static string DescribeOperations(IEnumerable<ManufacturingOperationBase> before, IEnumerable<ManufacturingOperationBase> after)
        {
            var was = (before ?? Enumerable.Empty<ManufacturingOperationBase>()).Select(DescribeOperation).ToList();
            var now = (after ?? Enumerable.Empty<ManufacturingOperationBase>()).Select(DescribeOperation).ToList();

            if (was.SequenceEqual(now))
                return null;

            var sb = new StringBuilder();
            sb.AppendLine("Было:");
            AppendList(sb, was);
            sb.AppendLine("Стало:");
            AppendList(sb, now);
            return sb.ToString().TrimEnd();
        }

        private static void AppendList(StringBuilder sb, List<string> items)
        {
            if (items.Count == 0)
                sb.AppendLine("  (нет операций)");
            foreach (var item in items)
                sb.AppendLine("  • " + item);
        }

        private static string DescribeOperation(ManufacturingOperationBase op)
        {
            string name = string.IsNullOrWhiteSpace(op.Name) ? op.TypeTitle : op.Name;
            var parts = new List<string> { name };

            if (op.Quantity != 1)
                parts.Add($"× {op.Quantity}");
            if (op.IsManual)
                parts.Add("добавлена вручную");
            if (op.IsEdited)
                parts.Add("параметры изменены");
            if (op.IsExcluded)
                parts.Add("исключена");
            else if (op.HasManualCost)
                parts.Add($"цена вручную {Num(op.ManualCost)} руб");
            else
                parts.Add($"{Num(op.Cost)} руб");

            return string.Join(", ", parts);
        }

        #endregion

        #region Изделие

        /// <summary>
        /// Сравнение с прежней сохранённой версией изделия (null — изделие сохраняется впервые)
        /// </summary>
        public static string DescribeProduct(Product before, Product after)
        {
            if (after == null) return null;

            if (before == null)
                return $"Новое изделие: {CountParts(after.Details)} дет., масса {Num(after.Mass)} кг";

            var lines = new List<string>();

            if (Math.Abs(before.Mass - after.Mass) > 0.001)
                lines.Add($"Масса изделия: {Num(before.Mass)} → {Num(after.Mass)} кг");

            var was = GroupParts(before.Details);
            var now = GroupParts(after.Details);

            var added = now.Keys.Where(k => !was.ContainsKey(k)).ToList();
            var removed = was.Keys.Where(k => !now.ContainsKey(k)).ToList();

            const int maxListed = 15;
            foreach (var key in added.Take(maxListed))
                lines.Add($"Добавлена: {PartTitle(now[key].Part)} × {now[key].Count}");
            foreach (var key in removed.Take(maxListed))
                lines.Add($"Удалена: {PartTitle(was[key].Part)} × {was[key].Count}");
            if (added.Count > maxListed || removed.Count > maxListed)
                lines.Add($"… всего добавлено {added.Count}, удалено {removed.Count}");

            int changedListed = 0;
            foreach (var key in now.Keys.Where(was.ContainsKey))
            {
                var a = was[key];
                var b = now[key];
                var diffs = new List<string>();

                if (a.Count != b.Count)
                    diffs.Add($"количество {a.Count} → {b.Count}");
                if (Math.Abs(a.Part.Mass - b.Part.Mass) > 0.001)
                    diffs.Add($"масса {Num(a.Part.Mass)} → {Num(b.Part.Mass)} кг");
                if (!string.Equals(a.Part.Material ?? "", b.Part.Material ?? "", StringComparison.Ordinal))
                    diffs.Add($"материал «{a.Part.Material}» → «{b.Part.Material}»");

                if (diffs.Count == 0)
                    continue;

                if (++changedListed <= maxListed)
                    lines.Add($"Изменена: {PartTitle(b.Part)}: {string.Join(", ", diffs)}");
            }
            if (changedListed > maxListed)
                lines.Add($"… всего изменено деталей: {changedListed}");

            return lines.Count == 0 ? "Без изменений состава" : string.Join("\n", lines);
        }

        private class PartGroup
        {
            public PartModel Part;
            public int Count;
        }

        private static Dictionary<string, PartGroup> GroupParts(IEnumerable<PartModel> parts)
        {
            var result = new Dictionary<string, PartGroup>(StringComparer.OrdinalIgnoreCase);
            foreach (var part in parts ?? Enumerable.Empty<PartModel>())
            {
                string key = $"{part.Name}|{part.Marking}";
                if (result.TryGetValue(key, out var group))
                    group.Count++;
                else
                    result[key] = new PartGroup { Part = part, Count = 1 };
            }
            return result;
        }

        private static int CountParts(IEnumerable<PartModel> parts) => parts?.Count() ?? 0;

        private static string PartTitle(PartModel part) =>
            string.IsNullOrEmpty(part.Marking) ? part.Name : $"{part.Name} ({part.Marking})";

        #endregion
    }
}
