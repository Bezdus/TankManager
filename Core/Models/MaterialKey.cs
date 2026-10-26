using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TankManager.Core.Models
{
    /// <summary>
    /// Материал, разобранный на вид проката, размер и марку: «Лист 3 ГОСТ 19904-90;AISI 304» и
    /// «Лист-3 AISI 304 ГОСТ 8568-77» — один материал «Лист 3 · AISI 304». По нему ищется цена
    /// </summary>
    public sealed class MaterialKey
    {
        /// <summary>
        /// Виды проката, которые распознаются в начале строки материала
        /// </summary>
        public static readonly string[] KnownKinds =
        {
            "Лист", "Рулон", "Лента", "Полоса", "Труба", "Круг", "Квадрат", "Шестигранник",
            "Уголок", "Швеллер", "Двутавр", "Профиль", "Пруток", "Проволока", "Стержень"
        };

        // Листовой прокат: размер — только толщина («Рулон 1,5х1250» → 1,5)
        private static readonly string[] SheetKinds = { "Лист", "Рулон", "Лента" };

        private static readonly string[] UnknownMaterials = { "без указания материала", "не определено", "не указан" };

        // Обозначение стандарта с номером: «ГОСТ 19904-90», «ТУ 38-1051082-86»
        private static readonly Regex StandardRegex = new Regex(
            @"(?<![\p{L}\d])(?:ГОСТ\s*Р|ГОСТ|ОСТ|ТУ|СТО|DIN|EN|ISO|ASTM)\s*\d[\w.\-–/]*",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // Размер сразу после вида проката: «38х2», «40х40х3», «1,5», «-3»
        private static readonly Regex SizeRegex = new Regex(
            @"^\s*[-–]?\s*(\d+(?:[.,]\d+)?(?:\s*[xXхХ×*]\s*\d+(?:[.,]\d+)?)*)",
            RegexOptions.CultureInvariant);

        private static readonly Regex SpacesRegex = new Regex(@"\s+");

        public static readonly MaterialKey Unknown = new MaterialKey(string.Empty, string.Empty, string.Empty);

        public MaterialKey(string kind, string size, string grade)
        {
            Kind = kind ?? string.Empty;
            Size = size ?? string.Empty;
            Grade = grade ?? string.Empty;
        }

        /// <summary>
        /// Вид проката («Лист», «Труба»; пусто — сортамент не указан, например «Сталь 10»)
        /// </summary>
        public string Kind { get; }

        /// <summary>
        /// Размер: толщина листа, сортамент трубы, диаметр круга («3», «38х2», «1,5»)
        /// </summary>
        public string Size { get; }

        /// <summary>
        /// Марка материала («AISI 304», «Сталь 10», «Ф-4»)
        /// </summary>
        public string Grade { get; }

        /// <summary>
        /// Материал не указан
        /// </summary>
        public bool IsUnknown => Kind.Length == 0 && Size.Length == 0 && Grade.Length == 0;

        /// <summary>
        /// «Лист 3 · AISI 304»
        /// </summary>
        public string Title => FormatTitle(Kind, Size, Grade, "Материал не указан");

        /// <summary>
        /// Ключ для сравнения: одинаков у по-разному записанных одинаковых материалов
        /// </summary>
        public string CompareKey => $"{TextKey(Kind)}|{SizeKey(Size)}|{TextKey(Grade)}";

        public static string FormatTitle(string kind, string size, string grade, string empty)
        {
            string sortament = string.Join(" ", new[] { kind, size }.Where(s => !string.IsNullOrEmpty(s)));
            string title = string.Join(" · ", new[] { sortament, grade }.Where(s => !string.IsNullOrEmpty(s)));
            return title.Length == 0 ? empty : title;
        }

        /// <summary>
        /// Разобрать строку материала (как в PartModel.Material)
        /// </summary>
        public static MaterialKey Parse(string material)
        {
            string text = Clean(material);
            if (text.Length == 0 || UnknownMaterials.Any(u => text.StartsWith(u, StringComparison.OrdinalIgnoreCase)))
                return Unknown;

            // КОМПАС пишет «Сортамент ГОСТ;Марка ГОСТ»
            string sortament = text;
            string grade = null;
            int semicolon = text.IndexOf(';');
            if (semicolon >= 0)
            {
                sortament = text.Substring(0, semicolon);
                grade = text.Substring(semicolon + 1);
            }

            string kind = KnownKinds.FirstOrDefault(k =>
                sortament.StartsWith(k, StringComparison.OrdinalIgnoreCase) &&
                (sortament.Length == k.Length || !char.IsLetter(sortament[k.Length])));

            if (kind == null)
                return new MaterialKey(string.Empty, string.Empty, Tidy(text.Replace(';', ' ')));

            string rest = sortament.Substring(kind.Length);
            string size = string.Empty;
            var match = SizeRegex.Match(rest);
            if (match.Success)
            {
                size = match.Groups[1].Value;
                rest = rest.Substring(match.Length);
            }

            if (SheetKinds.Contains(kind))
            {
                double thickness = PartModel.ParseSheetThickness(text);
                if (thickness > 0)
                    size = thickness.ToString("0.###", CultureInfo.InvariantCulture);
            }

            return new MaterialKey(kind, FormatSize(size), Tidy(grade ?? rest));
        }

        /// <summary>
        /// Размер в едином написании: кириллическая «х», десятичная запятая, без пробелов
        /// </summary>
        public static string FormatSize(string size)
        {
            if (string.IsNullOrWhiteSpace(size))
                return string.Empty;

            var sb = new StringBuilder(size.Length);
            foreach (char c in size)
            {
                if (char.IsWhiteSpace(c))
                    continue;
                if (c == 'x' || c == 'X' || c == 'Х' || c == '×' || c == '*')
                    sb.Append('х');
                else if (c == '.')
                    sb.Append(',');
                else
                    sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Ключ сравнения текста: без регистра, пробелов и дефисов, латинские буквы-двойники как кириллические
        /// («08X18H10» = «08Х18Н10», «Ф4» = «Ф-4»)
        /// </summary>
        public static string TextKey(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            const string latin = "aeopcxykmhtb";
            const string cyrillic = "аеорсхукмнтв";

            var sb = new StringBuilder(text.Length);
            foreach (char raw in text.ToLowerInvariant())
            {
                if (char.IsWhiteSpace(raw) || raw == '-' || raw == '–')
                    continue;
                int index = latin.IndexOf(raw);
                sb.Append(index >= 0 ? cyrillic[index] : raw);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Ключ сравнения размера: «88,9х2» = «88.9 x 2.0»
        /// </summary>
        public static string SizeKey(string size)
        {
            string formatted = FormatSize(size).ToLowerInvariant();
            if (formatted.Length == 0)
                return string.Empty;

            return string.Join("х", formatted.Split('х').Select(part =>
                double.TryParse(part.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                    ? value.ToString("0.###", CultureInfo.InvariantCulture)
                    : part));
        }

        private static string Clean(string material)
        {
            if (string.IsNullOrWhiteSpace(material))
                return string.Empty;

            string text = material.Replace("$d", " ").Replace("$", "");
            return SpacesRegex.Replace(text, " ").Trim();
        }

        // Марка без стандартов и лишних разделителей: « -AISI 304 ГОСТ 11068-81» → «AISI 304»
        private static string Tidy(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            text = StandardRegex.Replace(text, " ");
            text = SpacesRegex.Replace(text, " ");
            return text.Trim(' ', '-', '–', ',', ';', '.');
        }
    }
}
