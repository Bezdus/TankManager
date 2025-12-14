using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    public class ExcelService
    {
        /// <summary>
        /// Копирует список материалов в буфер обмена (масса)
        /// </summary>
        public void CopyMaterialsToClipboard(IEnumerable<MaterialInfo> materials)
        {
            CopyToClipboard(
                materials,
                "Список материалов пуст",
                "Материал\tМасса (кг)",
                m => $"{m.Name}\t{m.TotalMass:F2}");
        }

        /// <summary>
        /// Копирует список трубного проката в буфер обмена (длина)
        /// </summary>
        public void CopyTubularProductsToClipboard(IEnumerable<MaterialInfo> materials)
        {
            CopyToClipboard(
                materials,
                "Список материалов пуст",
                "Материал\tДлина (мм)",
                m => $"{m.Name}\t{m.TotalLength:F2}");
        }

        /// <summary>
        /// Копирует список деталей в буфер обмена с группировкой по уникальным деталям
        /// </summary>
        public void CopyPartsToClipboard(IEnumerable<PartModel> parts)
        {
            if (parts == null || !parts.Any())
            {
                MessageBox.Show("Список деталей пуст", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Наименование\tОбозначение\tМатериал\tКоличество\tМасса ед. (кг)\tМасса общ. (кг)");

                var groupedParts = parts
                    .GroupBy(p => new { p.Name, p.Marking, p.Material })
                    .OrderBy(g => g.Key.Name)
                    .ThenBy(g => g.Key.Marking);

                foreach (var group in groupedParts)
                {
                    int count = group.Count();
                    double unitMass = group.First().Mass;
                    double totalMass = group.Sum(p => p.Mass);

                    sb.AppendLine($"{group.Key.Name}\t{group.Key.Marking}\t{group.Key.Material}\t{count}\t{unitMass:F3}\t{totalMass:F3}");
                }

                Clipboard.SetText(sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при копировании в буфер обмена: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Универсальный метод копирования коллекции в буфер обмена
        /// </summary>
        private void CopyToClipboard<T>(
            IEnumerable<T> items,
            string emptyMessage,
            string header,
            Func<T, string> formatRow)
        {
            if (items == null || !items.Any())
            {
                MessageBox.Show(emptyMessage, "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var sb = new StringBuilder();
                sb.AppendLine(header);

                foreach (var item in items)
                {
                    sb.AppendLine(formatRow(item));
                }

                Clipboard.SetText(sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при копировании в буфер обмена: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
