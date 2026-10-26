using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace TankManager.Core.Models
{
    /// <summary>
    /// Строка списка материалов на вкладке «Цены»: материал изделия (по-разному записанные одинаковые
    /// материалы объединены) с расходом и суммой, или материал из расценок, которого в изделии нет
    /// </summary>
    public class MaterialPriceRow : INotifyPropertyChanged
    {
        private readonly Action<MaterialPriceRow, double?, MaterialPriceUnit> _priceEdited;

        private MaterialPriceEntry _rule;
        private bool _isExactRule;
        private int _partsCount;
        private double _totalMass;
        private double _totalLength;
        private double _cost;
        private IReadOnlyList<string> _materials = new string[0];
        private string _priceText;
        private MaterialPriceUnit _unit;

        public MaterialPriceRow(MaterialKey key, Action<MaterialPriceRow, double?, MaterialPriceUnit> priceEdited)
        {
            Key = key;
            _priceEdited = priceEdited;
        }

        public MaterialKey Key { get; }

        public string Title => Key.Title;

        /// <summary>
        /// Строки материала из КОМПАС, попавшие в эту строку (подсказка)
        /// </summary>
        public IReadOnlyList<string> Materials
        {
            get => _materials;
            private set { _materials = value; OnPropertyChanged(nameof(Materials)); OnPropertyChanged(nameof(MaterialsText)); }
        }

        public string MaterialsText => Materials.Count == 0 ? null : string.Join("\n", Materials);

        public int PartsCount
        {
            get => _partsCount;
            private set
            {
                _partsCount = value;
                OnPropertyChanged(nameof(PartsCount));
                OnPropertyChanged(nameof(AmountText));
                OnPropertyChanged(nameof(IsUsed));
                OnPropertyChanged(nameof(SourceText));
            }
        }

        /// <summary>
        /// Материал есть в открытом изделии (показываются расход и сумма)
        /// </summary>
        public bool IsUsed => PartsCount > 0;

        public double TotalMass
        {
            get => _totalMass;
            private set { _totalMass = value; OnPropertyChanged(nameof(TotalMass)); OnPropertyChanged(nameof(AmountText)); }
        }

        public double TotalLength
        {
            get => _totalLength;
            private set
            {
                _totalLength = value;
                OnPropertyChanged(nameof(TotalLength));
                OnPropertyChanged(nameof(AmountText));
                OnPropertyChanged(nameof(HasLength));
                OnPropertyChanged(nameof(ShowUnitChoice));
            }
        }

        /// <summary>
        /// Длина известна — можно задать цену за метр
        /// </summary>
        public bool HasLength => TotalLength > 0;

        /// <summary>
        /// Выбор руб/кг или руб/м (у материала без длины — только если цена уже за метр, чтобы её было видно)
        /// </summary>
        public bool ShowUnitChoice => HasLength || Unit == MaterialPriceUnit.Meter;

        public string UnitTitle => Unit == MaterialPriceUnit.Meter ? "руб/м" : "руб/кг";

        /// <summary>
        /// «125,3 кг · 12,4 м · 4 дет.»
        /// </summary>
        public string AmountText
        {
            get
            {
                var ru = CultureInfo.GetCultureInfo("ru-RU");
                var parts = new List<string> { TotalMass.ToString("N2", ru) + " кг" };
                if (HasLength)
                    parts.Add((TotalLength / 1000.0).ToString("N2", ru) + " м");
                parts.Add($"{PartsCount} дет.");
                return string.Join(" · ", parts);
            }
        }

        /// <summary>
        /// Стоимость материала всех деталей строки, руб
        /// </summary>
        public double Cost
        {
            get => _cost;
            private set { _cost = value; OnPropertyChanged(nameof(Cost)); }
        }

        /// <summary>
        /// Правило, по которому посчитана цена (null — цены нет)
        /// </summary>
        public MaterialPriceEntry Rule
        {
            get => _rule;
            private set
            {
                _rule = value;
                OnPropertyChanged(nameof(Rule));
                OnPropertyChanged(nameof(HasPrice));
                OnPropertyChanged(nameof(NeedsPrice));
                OnPropertyChanged(nameof(SourceText));
            }
        }

        public bool HasPrice => Rule != null;

        public bool CanHavePrice => !Key.IsUnknown;

        /// <summary>
        /// Цену нужно ввести (значок предупреждения)
        /// </summary>
        public bool NeedsPrice => CanHavePrice && !HasPrice;

        /// <summary>
        /// Цена задана общим правилом, а не для этого материала
        /// </summary>
        public bool IsExactRule
        {
            get => _isExactRule;
            private set { _isExactRule = value; OnPropertyChanged(nameof(IsExactRule)); OnPropertyChanged(nameof(SourceText)); }
        }

        /// <summary>
        /// Откуда цена: «Нет цены», «по правилу «Лист · AISI 304»»
        /// </summary>
        public string SourceText
        {
            get
            {
                if (Key.IsUnknown)
                    return "Материал не указан в КОМПАС — цену задать нельзя";
                if (Rule == null)
                    return "Нет цены — введите цену материала";
                if (Rule.Unit == MaterialPriceUnit.Meter && IsUsed && !HasLength)
                    return $"Цена за метр («{Rule.Title}»), но длина деталей неизвестна";
                return IsExactRule ? null : $"По правилу «{Rule.Title}»";
            }
        }

        /// <summary>
        /// Цена в поле ввода; пустая строка — удалить цену этого материала
        /// </summary>
        public string PriceText
        {
            get => _priceText;
            set
            {
                if (_priceText == value)
                    return;

                string text = (value ?? string.Empty).Replace(" ", string.Empty).Replace(" ", string.Empty).Replace(',', '.');
                double? price = null;
                if (text.Length > 0)
                {
                    if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
                        parsed < 0 || double.IsNaN(parsed) || double.IsInfinity(parsed))
                        throw new FormatException("Введите число, например 12,5");
                    price = parsed;
                }

                if (price == null ? IsExactRule : !SamePrice(price.Value, _unit))
                {
                    _priceText = value;
                    _priceEdited?.Invoke(this, price, _unit);
                }
                else
                {
                    // Цена не изменилась (или стёрта цена общего правила) — показываем действующую
                    _priceText = FormatPrice(Rule);
                }
                OnPropertyChanged(nameof(PriceText));
            }
        }

        /// <summary>
        /// Единица цены (руб/кг или руб/м); смена единицы при заданной цене сохраняет правило
        /// </summary>
        public MaterialPriceUnit Unit
        {
            get => _unit;
            set
            {
                if (_unit == value)
                    return;

                _unit = value;
                OnPropertyChanged(nameof(Unit));
                OnPropertyChanged(nameof(UnitTitle));
                OnPropertyChanged(nameof(ShowUnitChoice));

                if (Rule != null && !SamePrice(Rule.Price.Value, value))
                    _priceEdited?.Invoke(this, Rule.Price, value);
            }
        }

        private bool SamePrice(double price, MaterialPriceUnit unit) =>
            Rule != null && Rule.Unit == unit && Math.Abs(Rule.Price.Value - price) < 0.0001;

        /// <summary>
        /// Обновить расход, цену и сумму (строка переиспользуется, чтобы не терять фокус поля ввода)
        /// </summary>
        public void Update(IEnumerable<PartModel> parts, IReadOnlyList<string> materials, PricingSettings pricing)
        {
            int count = 0;
            double mass = 0, length = 0, cost = 0;
            foreach (var part in parts)
            {
                count++;
                mass += part.Mass;
                if (part.Length > 0)
                    length += part.Length;
                cost += part.MetalCost;
            }

            PartsCount = count;
            TotalMass = mass;
            TotalLength = length;
            Cost = cost;
            Materials = materials;

            var rule = pricing?.FindMaterialPrice(Key);
            Rule = rule;
            IsExactRule = rule != null && rule.IsExactFor(Key);

            _unit = rule?.Unit ?? (HasLength ? MaterialPriceUnit.Meter : MaterialPriceUnit.Kilogram);
            OnPropertyChanged(nameof(Unit));
            OnPropertyChanged(nameof(UnitTitle));
            OnPropertyChanged(nameof(ShowUnitChoice));
            _priceText = FormatPrice(rule);
            OnPropertyChanged(nameof(PriceText));
        }

        /// <summary>
        /// Подходит под строку поиска (без регистра, латиница и кириллица одинаково)
        /// </summary>
        public bool MatchesSearch(string search)
        {
            string key = MaterialKey.TextKey(search);
            if (key.Length == 0)
                return true;

            return MaterialKey.TextKey(Title).Contains(key) ||
                   Materials.Any(m => MaterialKey.TextKey(m).Contains(key));
        }

        private static string FormatPrice(MaterialPriceEntry rule) =>
            rule?.Price?.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
