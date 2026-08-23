using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using ClosedXML.Excel;
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
                sb.AppendLine("Наименование\tОбозначение\tМатериал\tКоличество\tМасса ед. (кг)\tМасса общ. (кг)\tСтоимость металла (руб)\tСтоимость операций (руб)\tОбщая стоимость (руб)");

                var groupedParts = parts
                    .GroupBy(p => new { p.Name, p.Marking, p.Material })
                    .OrderBy(g => g.Key.Name)
                    .ThenBy(g => g.Key.Marking);

                foreach (var group in groupedParts)
                {
                    int count = group.Count();
                    double unitMass = group.First().Mass;
                    double totalMass = group.Sum(p => p.Mass);
                    double metalCost = group.Sum(p => p.MetalCost);
                    double opsCost = group.Sum(p => p.OperationsCost);
                    double totalCost = group.Sum(p => p.TotalCost);

                    sb.AppendLine($"{group.Key.Name}\t{group.Key.Marking}\t{group.Key.Material}\t{count}\t{unitMass:F3}\t{totalMass:F3}\t{metalCost:F2}\t{opsCost:F2}\t{totalCost:F2}");
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
        /// Копирует все данные в буфер обмена: покупные детали, листовые материалы, трубные материалы, прочие материалы через пустой столбец
        /// </summary>
        public void CopyAllDataToClipboard(
            IEnumerable<PartModel> standardParts,
            IEnumerable<MaterialInfo> sheetMaterials,
            IEnumerable<MaterialInfo> tubularProducts,
            IEnumerable<MaterialInfo> otherMaterials)
        {
            try
            {
                var standardPartsList = standardParts?.ToList() ?? new List<PartModel>();
                var sheetMaterialsList = sheetMaterials?.ToList() ?? new List<MaterialInfo>();
                var tubularProductsList = tubularProducts?.ToList() ?? new List<MaterialInfo>();
                var otherMaterialsList = otherMaterials?.ToList() ?? new List<MaterialInfo>();

                if (!standardPartsList.Any() && !sheetMaterialsList.Any() && !tubularProductsList.Any() && !otherMaterialsList.Any())
                {
                    MessageBox.Show("Нет данных для копирования", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Подготовка данных покупных деталей с группировкой
                var groupedParts = standardPartsList
                    .GroupBy(p => new { p.Name, p.Marking, p.Material })
                    .OrderBy(g => g.Key.Name)
                    .ThenBy(g => g.Key.Marking)
                    .Select(g => new
                    {
                        Name = g.Key.Name,
                        Marking = g.Key.Marking,
                        Material = g.Key.Material,
                        Count = g.Count(),
                        UnitMass = g.First().Mass,
                        TotalMass = g.Sum(p => p.Mass),
                        MetalCost = g.Sum(p => p.MetalCost),
                        OperationsCost = g.Sum(p => p.OperationsCost),
                        TotalCost = g.Sum(p => p.TotalCost)
                    })
                    .ToList();

                // Определяем максимальное количество строк
                int maxRows = Math.Max(Math.Max(Math.Max(groupedParts.Count, sheetMaterialsList.Count), tubularProductsList.Count), otherMaterialsList.Count);

                var sb = new StringBuilder();

                // Заголовки: Покупные детали | пустой столбец | Листовые материалы | пустой столбец | Трубные материалы | пустой столбец | Прочие материалы
                sb.AppendLine("Наименование\tОбозначение\tМатериал\tКоличество\tМасса ед. (кг)\tМасса общ. (кг)\tСтоим. металла (руб)\tСтоим. операций (руб)\tОбщая стоим. (руб)\t\tМатериал\tМасса (кг)\t\tМатериал\tДлина (мм)\t\tМатериал\tМасса (кг)");

                for (int i = 0; i < maxRows; i++)
                {
                    var row = new List<string>();

                    // Покупные детали (9 столбцов)
                    if (i < groupedParts.Count)
                    {
                        var part = groupedParts[i];
                        row.Add(part.Name ?? "");
                        row.Add(part.Marking ?? "");
                        row.Add(part.Material ?? "");
                        row.Add(part.Count.ToString());
                        row.Add(part.UnitMass.ToString("F3"));
                        row.Add(part.TotalMass.ToString("F3"));
                        row.Add(part.MetalCost.ToString("F2"));
                        row.Add(part.OperationsCost.ToString("F2"));
                        row.Add(part.TotalCost.ToString("F2"));
                    }
                    else
                    {
                        row.AddRange(new[] { "", "", "", "", "", "", "", "", "" });
                    }

                    // Пустой столбец
                    row.Add("");

                    // Листовые материалы (2 столбца)
                    if (i < sheetMaterialsList.Count)
                    {
                        var material = sheetMaterialsList[i];
                        row.Add(material.Name ?? "");
                        row.Add(material.TotalMass.ToString("F2"));
                    }
                    else
                    {
                        row.AddRange(new[] { "", "" });
                    }

                    // Пустой столбец
                    row.Add("");

                    // Трубные материалы (2 столбца)
                    if (i < tubularProductsList.Count)
                    {
                        var tubular = tubularProductsList[i];
                        row.Add(tubular.Name ?? "");
                        row.Add(tubular.TotalLength.ToString("F2"));
                    }
                    else
                    {
                        row.AddRange(new[] { "", "" });
                    }

                    // Пустой столбец
                    row.Add("");

                    // Прочие материалы (2 столбца)
                    if (i < otherMaterialsList.Count)
                    {
                        var other = otherMaterialsList[i];
                        row.Add(other.Name ?? "");
                        row.Add(other.TotalMass.ToString("F2"));
                    }
                    else
                    {
                        row.AddRange(new[] { "", "" });
                    }

                    sb.AppendLine(string.Join("\t", row));
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
        /// Экспортирует все данные изделия в Excel файл
        /// </summary>
        public string ExportToExcelFile(
            string productName,
            IEnumerable<PartModel> allDetails,
            IEnumerable<PartModel> standardParts,
            IEnumerable<MaterialInfo> sheetMaterials,
            IEnumerable<MaterialInfo> tubularProducts,
            IEnumerable<MaterialInfo> otherMaterials)
        {
            var allDetailsList = allDetails?.ToList() ?? new List<PartModel>();
            var standardPartsList = standardParts?.ToList() ?? new List<PartModel>();
            var sheetMaterialsList = sheetMaterials?.ToList() ?? new List<MaterialInfo>();
            var tubularProductsList = tubularProducts?.ToList() ?? new List<MaterialInfo>();
            var otherMaterialsList = otherMaterials?.ToList() ?? new List<MaterialInfo>();

            if (!allDetailsList.Any() && !standardPartsList.Any() && !sheetMaterialsList.Any() && !tubularProductsList.Any() && !otherMaterialsList.Any())
            {
                throw new InvalidOperationException("Нет данных для экспорта");
            }

            var safeName = string.Join("_", (productName ?? "Изделие").Split(Path.GetInvalidFileNameChars()));
            var fileName = $"Ведомость материалов {safeName}.xlsx";

            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = fileName,
                DefaultExt = ".xlsx",
                Filter = "Excel файлы (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() != true)
                return null;

            var filePath = dialog.FileName;

            using (var workbook = new XLWorkbook())
            {
                // Лист 1: Все детали
                if (allDetailsList.Any())
                {
                    var wsDetails = workbook.Worksheets.Add("Детали");
                    var groupedDetails = allDetailsList
                        .GroupBy(p => new { p.Name, p.Marking, p.Material })
                        .OrderBy(g => g.Key.Name)
                        .ThenBy(g => g.Key.Marking)
                        .ToList();

                    wsDetails.Cell(1, 1).Value = "Наименование";
                    wsDetails.Cell(1, 2).Value = "Обозначение";
                    wsDetails.Cell(1, 3).Value = "Материал";
                    wsDetails.Cell(1, 4).Value = "Количество";
                    wsDetails.Cell(1, 5).Value = "Масса ед. (кг)";
                    wsDetails.Cell(1, 6).Value = "Масса общ. (кг)";
                    wsDetails.Cell(1, 7).Value = "Стоимость металла (руб)";
                    wsDetails.Cell(1, 8).Value = "Стоимость операций (руб)";
                    wsDetails.Cell(1, 9).Value = "Общая стоимость (руб)";
                    StyleHeaderRow(wsDetails, 1, 9);

                    for (int i = 0; i < groupedDetails.Count; i++)
                    {
                        var g = groupedDetails[i];
                        int row = i + 2;
                        int count = g.Count();
                        wsDetails.Cell(row, 1).Value = g.Key.Name ?? "";
                        wsDetails.Cell(row, 2).Value = g.Key.Marking ?? "";
                        wsDetails.Cell(row, 3).Value = g.Key.Material ?? "";
                        wsDetails.Cell(row, 4).Value = count;
                        wsDetails.Cell(row, 5).Value = Math.Round(g.First().Mass, 3);
                        wsDetails.Cell(row, 6).Value = Math.Round(g.Sum(p => p.Mass), 3);
                        wsDetails.Cell(row, 7).Value = Math.Round(g.Sum(p => p.MetalCost), 2);
                        wsDetails.Cell(row, 8).Value = Math.Round(g.Sum(p => p.OperationsCost), 2);
                        wsDetails.Cell(row, 9).Value = Math.Round(g.Sum(p => p.TotalCost), 2);
                    }

                    wsDetails.Columns().AdjustToContents();
                }

                // Лист 2: Покупные детали
                if (standardPartsList.Any())
                {
                    var wsStandard = workbook.Worksheets.Add("Покупные");
                    var groupedParts = standardPartsList
                        .GroupBy(p => new { p.Name, p.Marking, p.Material })
                        .OrderBy(g => g.Key.Name)
                        .ThenBy(g => g.Key.Marking)
                        .ToList();

                    wsStandard.Cell(1, 1).Value = "Наименование";
                    wsStandard.Cell(1, 2).Value = "Обозначение";
                    wsStandard.Cell(1, 3).Value = "Материал";
                    wsStandard.Cell(1, 4).Value = "Количество";
                    wsStandard.Cell(1, 5).Value = "Масса ед. (кг)";
                    wsStandard.Cell(1, 6).Value = "Масса общ. (кг)";
                    wsStandard.Cell(1, 7).Value = "Стоимость металла (руб)";
                    wsStandard.Cell(1, 8).Value = "Стоимость операций (руб)";
                    wsStandard.Cell(1, 9).Value = "Общая стоимость (руб)";
                    StyleHeaderRow(wsStandard, 1, 9);

                    for (int i = 0; i < groupedParts.Count; i++)
                    {
                        var g = groupedParts[i];
                        int row = i + 2;
                        int count = g.Count();
                        wsStandard.Cell(row, 1).Value = g.Key.Name ?? "";
                        wsStandard.Cell(row, 2).Value = g.Key.Marking ?? "";
                        wsStandard.Cell(row, 3).Value = g.Key.Material ?? "";
                        wsStandard.Cell(row, 4).Value = count;
                        wsStandard.Cell(row, 5).Value = Math.Round(g.First().Mass, 3);
                        wsStandard.Cell(row, 6).Value = Math.Round(g.Sum(p => p.Mass), 3);
                        wsStandard.Cell(row, 7).Value = Math.Round(g.Sum(p => p.MetalCost), 2);
                        wsStandard.Cell(row, 8).Value = Math.Round(g.Sum(p => p.OperationsCost), 2);
                        wsStandard.Cell(row, 9).Value = Math.Round(g.Sum(p => p.TotalCost), 2);
                    }

                    wsStandard.Columns().AdjustToContents();
                }

                // Лист 3: Листовой прокат
                if (sheetMaterialsList.Any())
                {
                    var wsSheet = workbook.Worksheets.Add("Листовой прокат");
                    wsSheet.Cell(1, 1).Value = "Материал";
                    wsSheet.Cell(1, 2).Value = "Масса (кг)";
                    StyleHeaderRow(wsSheet, 1, 2);

                    for (int i = 0; i < sheetMaterialsList.Count; i++)
                    {
                        var m = sheetMaterialsList[i];
                        wsSheet.Cell(i + 2, 1).Value = m.Name ?? "";
                        wsSheet.Cell(i + 2, 2).Value = Math.Round(m.TotalMass, 2);
                    }

                    wsSheet.Columns().AdjustToContents();
                }

                // Лист 4: Трубный прокат
                if (tubularProductsList.Any())
                {
                    var wsTubular = workbook.Worksheets.Add("Трубный прокат");
                    wsTubular.Cell(1, 1).Value = "Материал";
                    wsTubular.Cell(1, 2).Value = "Длина (мм)";
                    StyleHeaderRow(wsTubular, 1, 2);

                    for (int i = 0; i < tubularProductsList.Count; i++)
                    {
                        var t = tubularProductsList[i];
                        wsTubular.Cell(i + 2, 1).Value = t.Name ?? "";
                        wsTubular.Cell(i + 2, 2).Value = Math.Round(t.TotalLength, 2);
                    }

                    wsTubular.Columns().AdjustToContents();
                }

                // Лист 5: Прочие материалы
                if (otherMaterialsList.Any())
                {
                    var wsOther = workbook.Worksheets.Add("Прочие материалы");
                    wsOther.Cell(1, 1).Value = "Материал";
                    wsOther.Cell(1, 2).Value = "Масса (кг)";
                    StyleHeaderRow(wsOther, 1, 2);

                    for (int i = 0; i < otherMaterialsList.Count; i++)
                    {
                        var o = otherMaterialsList[i];
                        wsOther.Cell(i + 2, 1).Value = o.Name ?? "";
                        wsOther.Cell(i + 2, 2).Value = Math.Round(o.TotalMass, 2);
                    }

                    wsOther.Columns().AdjustToContents();
                }

                workbook.SaveAs(filePath);
            }

            return filePath;
        }

        private static void StyleHeaderRow(IXLWorksheet ws, int row, int columnCount)
        {
            var headerRange = ws.Range(row, 1, row, columnCount);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF505050");
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
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
