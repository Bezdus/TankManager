using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
                m => $"{Cell(m.Name)}\t{m.TotalMass:F2}");
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
                m => $"{Cell(m.Name)}\t{m.TotalLength:F2}");
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

                    sb.AppendLine($"{Cell(group.Key.Name)}\t{Cell(group.Key.Marking)}\t{Cell(group.Key.Material)}\t{count}\t{unitMass:F3}\t{totalMass:F3}\t{metalCost:F2}\t{opsCost:F2}\t{totalCost:F2}");
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
                        row.Add(Cell(part.Name));
                        row.Add(Cell(part.Marking));
                        row.Add(Cell(part.Material));
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
                        row.Add(Cell(material.Name));
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
                        row.Add(Cell(tubular.Name));
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
                        row.Add(Cell(other.Name));
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
        /// Открывает ведомость материалов в новой книге Excel (без сохранения файла — пользователь сохраняет сам)
        /// </summary>
        public void OpenInExcel(
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

            // Название листа, данные (первая строка — заголовок), количество текстовых столбцов слева
            var sheets = new List<Tuple<string, object[,], int>>();
            if (allDetailsList.Any())
                sheets.Add(Tuple.Create("Детали", BuildPartsSheet(allDetailsList), 3));
            if (standardPartsList.Any())
                sheets.Add(Tuple.Create("Покупные", BuildPartsSheet(standardPartsList), 3));
            if (sheetMaterialsList.Any())
                sheets.Add(Tuple.Create("Листовой прокат", BuildMaterialsSheet(sheetMaterialsList, "Масса (кг)", m => m.TotalMass), 1));
            if (tubularProductsList.Any())
                sheets.Add(Tuple.Create("Трубный прокат", BuildMaterialsSheet(tubularProductsList, "Длина (мм)", m => m.TotalLength), 1));
            if (otherMaterialsList.Any())
                sheets.Add(Tuple.Create("Прочие материалы", BuildMaterialsSheet(otherMaterialsList, "Масса (кг)", m => m.TotalMass), 1));

            if (sheets.Count == 0)
            {
                throw new InvalidOperationException("Нет данных для экспорта");
            }

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException("Microsoft Excel не установлен");
            }

            var comObjects = new List<object>();
            dynamic excel = null;
            dynamic workbook = null;
            bool shown = false;

            try
            {
                excel = Activator.CreateInstance(excelType);
                excel.ScreenUpdating = false;

                dynamic workbooks = Track(comObjects, excel.Workbooks);
                workbook = Track(comObjects, workbooks.Add());
                dynamic worksheets = Track(comObjects, workbook.Worksheets);

                // Новая книга может содержать несколько пустых листов — оставляем один
                excel.DisplayAlerts = false;
                while (worksheets.Count > 1)
                {
                    dynamic extra = worksheets[worksheets.Count];
                    extra.Delete();
                    Release(extra);
                }
                excel.DisplayAlerts = true;

                dynamic lastSheet = Track(comObjects, worksheets[1]);
                for (int i = 0; i < sheets.Count; i++)
                {
                    dynamic ws = i == 0 ? lastSheet : Track(comObjects, worksheets.Add(After: lastSheet));
                    WriteSheet(ws, sheets[i].Item1, sheets[i].Item2, sheets[i].Item3, comObjects);
                    lastSheet = ws;
                }

                workbook.Title = $"Ведомость материалов {productName ?? "Изделие"}";

                dynamic firstSheet = Track(comObjects, worksheets[1]);
                firstSheet.Activate();

                excel.ScreenUpdating = true;
                excel.Visible = true;
                excel.UserControl = true;
                shown = true;
            }
            catch
            {
                if (excel != null && !shown)
                {
                    try
                    {
                        if (workbook != null)
                            workbook.Close(false);
                        excel.Quit();
                    }
                    catch
                    {
                        // Excel уже недоступен — закрывать нечего
                    }
                }
                throw;
            }
            finally
            {
                // Освобождаем COM-ссылки, иначе процесс Excel не завершится после закрытия пользователем
                for (int i = comObjects.Count - 1; i >= 0; i--)
                    Release(comObjects[i]);
                Release(excel);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Данные листа деталей с группировкой по уникальным деталям
        /// </summary>
        private static object[,] BuildPartsSheet(List<PartModel> parts)
        {
            var groups = parts
                .GroupBy(p => new { p.Name, p.Marking, p.Material })
                .OrderBy(g => g.Key.Name)
                .ThenBy(g => g.Key.Marking)
                .ToList();

            string[] header =
            {
                "Наименование", "Обозначение", "Материал", "Количество", "Масса ед. (кг)", "Масса общ. (кг)",
                "Стоимость металла (руб)", "Стоимость операций (руб)", "Общая стоимость (руб)"
            };

            var data = new object[groups.Count + 1, header.Length];
            for (int c = 0; c < header.Length; c++)
                data[0, c] = header[c];

            for (int i = 0; i < groups.Count; i++)
            {
                var g = groups[i];
                int row = i + 1;
                data[row, 0] = g.Key.Name ?? "";
                data[row, 1] = g.Key.Marking ?? "";
                data[row, 2] = g.Key.Material ?? "";
                data[row, 3] = g.Count();
                data[row, 4] = Math.Round(g.First().Mass, 3);
                data[row, 5] = Math.Round(g.Sum(p => p.Mass), 3);
                data[row, 6] = Math.Round(g.Sum(p => p.MetalCost), 2);
                data[row, 7] = Math.Round(g.Sum(p => p.OperationsCost), 2);
                data[row, 8] = Math.Round(g.Sum(p => p.TotalCost), 2);
            }

            return data;
        }

        /// <summary>
        /// Данные листа материалов: наименование и одно числовое значение
        /// </summary>
        private static object[,] BuildMaterialsSheet(List<MaterialInfo> materials, string valueHeader, Func<MaterialInfo, double> value)
        {
            var data = new object[materials.Count + 1, 2];
            data[0, 0] = "Материал";
            data[0, 1] = valueHeader;

            for (int i = 0; i < materials.Count; i++)
            {
                data[i + 1, 0] = materials[i].Name ?? "";
                data[i + 1, 1] = Math.Round(value(materials[i]), 2);
            }

            return data;
        }

        /// <summary>
        /// Заполняет лист Excel: данные одним блоком, оформление заголовка, ширина столбцов.
        /// Текстовые столбцы получают формат "@", чтобы Excel не превращал значения в формулы, даты и числа.
        /// </summary>
        private static void WriteSheet(dynamic ws, string name, object[,] data, int textColumns, List<object> comObjects)
        {
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            ws.Name = name;

            dynamic cells = Track(comObjects, ws.Cells);
            dynamic topLeft = Track(comObjects, cells[1, 1]);
            dynamic bottomRight = Track(comObjects, cells[rows, cols]);
            dynamic textBottomRight = Track(comObjects, cells[rows, textColumns]);
            dynamic headerRight = Track(comObjects, cells[1, cols]);

            dynamic textRange = Track(comObjects, ws.Range[topLeft, textBottomRight]);
            textRange.NumberFormat = "@";

            dynamic range = Track(comObjects, ws.Range[topLeft, bottomRight]);
            range.Value2 = data;

            dynamic header = Track(comObjects, ws.Range[topLeft, headerRight]);
            dynamic font = Track(comObjects, header.Font);
            font.Bold = true;
            font.Color = 0xFFFFFF;
            dynamic interior = Track(comObjects, header.Interior);
            interior.Color = 0x505050;
            header.HorizontalAlignment = -4108; // xlCenter

            dynamic columns = Track(comObjects, range.Columns);
            columns.AutoFit();
        }

        private static dynamic Track(List<object> comObjects, object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
                comObjects.Add(comObject);
            return comObject;
        }

        private static void Release(object comObject)
        {
            try
            {
                if (comObject != null && Marshal.IsComObject(comObject))
                    Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
                // Объект уже освобождён
            }
        }

        /// <summary>
        /// Подготавливает значение для вставки в TSV: без табов и переводов строк,
        /// а строки, начинающиеся с символов формулы, защищены апострофом (защита от формульной инъекции в Excel)
        /// </summary>
        private static string Cell(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            string text = value.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');

            char first = text[0];
            if (first == '=' || first == '+' || first == '-' || first == '@')
                text = "'" + text;

            return text;
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
