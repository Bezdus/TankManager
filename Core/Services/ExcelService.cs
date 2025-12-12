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
        /// Копирует список материалов в буфер обмена в формате, подходящем для вставки в Excel
        /// </summary>
        /// <param name="materials">Список материалов для копирования</param>
        public void CopyMaterialsToClipboard(IEnumerable<MaterialInfo> materials)
        {
            if (materials == null || !materials.Any())
            {
                MessageBox.Show("Список материалов пуст", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var sb = new StringBuilder();

                // Заголовки столбцов
                sb.AppendLine("Материал\tМасса (кг)");

                // Данные
                foreach (var material in materials)
                {
                    sb.AppendLine($"{material.Name}\t{material.TotalMass:F2}");
                }

                // Копируем в буфер обмена
                Clipboard.SetText(sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при копировании в буфер обмена: {ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Копирует список деталей в буфер обмена в формате, подходящем для вставки в Excel
        /// </summary>
        /// <param name="parts">Список деталей для копирования</param>
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

                // Заголовки столбцов
                sb.AppendLine("Наименование\tОбозначение\tМатериал\tМасса (кг)");

                // Данные
                foreach (var part in parts)
                {
                    sb.AppendLine($"{part.Name}\t{part.Marking}\t{part.Material}\t{part.Mass:F3}");
                }

                // Копируем в буфер обмена
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
