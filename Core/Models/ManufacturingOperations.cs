using System;
using System.ComponentModel;

namespace TankManager.Core.Models
{
    /// <summary>
    /// Тип операции изготовления детали
    /// </summary>
    public enum ManufacturingOperationType
    {
        /// <summary>
        /// Резка на лазерном станке
        /// </summary>
        LaserCutting,

        /// <summary>
        /// Гибка
        /// </summary>
        Bending,

        /// <summary>
        /// Вальцовка
        /// </summary>
        Rolling,

        /// <summary>
        /// Отбортовка
        /// </summary>
        Flanging,

        /// <summary>
        /// Прочая операция, заданная пользователем (сварка, зачистка, покраска...)
        /// </summary>
        Custom
    }

    /// <summary>
    /// Откуда взялась операция
    /// </summary>
    public enum OperationOrigin
    {
        /// <summary>
        /// Определена автоматически по модели КОМПАС
        /// </summary>
        Kompas,

        /// <summary>
        /// Добавлена пользователем вручную
        /// </summary>
        Manual
    }

    /// <summary>
    /// Базовый класс операции изготовления детали
    /// </summary>
    public abstract class ManufacturingOperationBase : INotifyPropertyChanged
    {
        private string _name;
        private double _cost;
        private double _timeMinutes;
        private double _materialThickness;
        private OperationOrigin _origin;
        private int _quantity = 1;
        private bool _isExcluded;
        private bool _hasManualCost;
        private double _manualCost;
        private bool _isEdited;

        /// <summary>
        /// Тип операции
        /// </summary>
        public ManufacturingOperationType Type { get; }

        /// <summary>
        /// Название типа операции для интерфейса
        /// </summary>
        public string TypeTitle
        {
            get
            {
                switch (Type)
                {
                    case ManufacturingOperationType.LaserCutting: return "Резка на лазере";
                    case ManufacturingOperationType.Bending: return "Гибка";
                    case ManufacturingOperationType.Rolling: return "Вальцовка";
                    case ManufacturingOperationType.Flanging: return "Отбортовка";
                    default: return "Прочая операция";
                }
            }
        }

        /// <summary>
        /// Название операции
        /// </summary>
        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        /// <summary>
        /// Стоимость операции, руб
        /// </summary>
        public double Cost
        {
            get { return _cost; }
            set
            {
                if (Math.Abs(_cost - value) > 0.0001)
                {
                    _cost = value;
                    OnPropertyChanged(nameof(Cost));
                }
            }
        }

        /// <summary>
        /// Время выполнения, мин
        /// </summary>
        public double TimeMinutes
        {
            get { return _timeMinutes; }
            set
            {
                if (Math.Abs(_timeMinutes - value) > 0.0001)
                {
                    _timeMinutes = value;
                    OnPropertyChanged(nameof(TimeMinutes));
                }
            }
        }

        /// <summary>
        /// Толщина материала, мм
        /// </summary>
        public double MaterialThickness
        {
            get { return _materialThickness; }
            set
            {
                if (Math.Abs(_materialThickness - value) > 0.0001)
                {
                    _materialThickness = value;
                    OnPropertyChanged(nameof(MaterialThickness));
                }
            }
        }

        /// <summary>
        /// Определена по модели КОМПАС или добавлена вручную
        /// </summary>
        public OperationOrigin Origin
        {
            get { return _origin; }
            set
            {
                if (_origin != value)
                {
                    _origin = value;
                    OnPropertyChanged(nameof(Origin));
                    OnPropertyChanged(nameof(IsManual));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool IsManual => Origin == OperationOrigin.Manual;

        /// <summary>
        /// Количество (например, число гибов); стоимость единицы умножается на него
        /// </summary>
        public int Quantity
        {
            get { return _quantity; }
            set
            {
                if (value < 0) value = 0;
                if (_quantity != value)
                {
                    _quantity = value;
                    OnPropertyChanged(nameof(Quantity));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        /// <summary>
        /// Операция из КОМПАС, исключённая пользователем из расчёта. Не удаляется, чтобы правка
        /// пережила обновление данных из КОМПАС
        /// </summary>
        public bool IsExcluded
        {
            get { return _isExcluded; }
            set
            {
                if (_isExcluded != value)
                {
                    _isExcluded = value;
                    OnPropertyChanged(nameof(IsExcluded));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        /// <summary>
        /// Стоимость задана вручную и не пересчитывается по расценкам
        /// </summary>
        public bool HasManualCost
        {
            get { return _hasManualCost; }
            set
            {
                if (_hasManualCost != value)
                {
                    _hasManualCost = value;
                    OnPropertyChanged(nameof(HasManualCost));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        /// <summary>
        /// Стоимость, заданная вручную, руб (действует при HasManualCost)
        /// </summary>
        public double ManualCost
        {
            get { return _manualCost; }
            set
            {
                if (Math.Abs(_manualCost - value) > 0.0001)
                {
                    _manualCost = value;
                    OnPropertyChanged(nameof(ManualCost));
                }
            }
        }

        /// <summary>
        /// Пользователь изменил параметры операции, определённые по модели КОМПАС
        /// </summary>
        public bool IsEdited
        {
            get { return _isEdited; }
            set
            {
                if (_isEdited != value)
                {
                    _isEdited = value;
                    OnPropertyChanged(nameof(IsEdited));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        /// <summary>
        /// Операция отличается от определённой по КОМПАС (есть ручные правки)
        /// </summary>
        public bool IsModified => IsManual || IsExcluded || HasManualCost || IsEdited || Quantity != 1;

        protected ManufacturingOperationBase(ManufacturingOperationType type)
        {
            Type = type;
            _name = string.Empty;
        }

        /// <summary>
        /// Рассчитать стоимость операции на основе расценок (с учётом ручных правок)
        /// </summary>
        /// <param name="settings">Настройки расценок</param>
        public void CalculateCost(PricingSettings settings)
        {
            if (IsExcluded)
            {
                Cost = 0;
                return;
            }

            if (HasManualCost)
            {
                Cost = ManualCost;
                return;
            }

            if (settings == null) return;
            Cost = CalculateUnitCost(settings) * Quantity;
        }

        /// <summary>
        /// Стоимость одной операции по расценкам
        /// </summary>
        protected abstract double CalculateUnitCost(PricingSettings settings);

        /// <summary>
        /// Достаточно ли исходных данных для расчёта стоимости (иначе цена операции занижена)
        /// </summary>
        public bool IsCostReliable => IsExcluded || HasManualCost || IsUnitCostReliable;

        /// <summary>
        /// Достаточно ли исходных данных для расчёта по расценкам
        /// </summary>
        protected virtual bool IsUnitCostReliable => true;

        /// <summary>
        /// Копия операции без подписчиков PropertyChanged
        /// </summary>
        public ManufacturingOperationBase Clone()
        {
            var clone = (ManufacturingOperationBase)MemberwiseClone();
            clone.PropertyChanged = null;
            return clone;
        }

        /// <summary>
        /// Скопировать параметры операции (геометрию), не трогая признаки ручных правок
        /// </summary>
        public virtual void CopyParametersFrom(ManufacturingOperationBase source)
        {
            if (source == null) return;
            TimeMinutes = source.TimeMinutes;
            MaterialThickness = source.MaterialThickness;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Операция резки на лазерном станке
    /// </summary>
    public class LaserCuttingOperation : ManufacturingOperationBase
    {
        private double _cutLength;
        private double _engravingLength;

        public LaserCuttingOperation()
            : base(ManufacturingOperationType.LaserCutting)
        {
        }

        protected override double CalculateUnitCost(PricingSettings settings)
        {
            // Цена реза зависит от толщины листа (MaterialThickness выставляется при пересчёте стоимости)
            // Длина реза в мм, цена в руб/м
            return CutLength / 1000.0 * settings.GetLaserCuttingPricePerMeter(MaterialThickness)
                + EngravingLength * settings.EngravingPricePerMm;
        }

        public override void CopyParametersFrom(ManufacturingOperationBase source)
        {
            base.CopyParametersFrom(source);
            if (source is LaserCuttingOperation laser)
            {
                CutLength = laser.CutLength;
                EngravingLength = laser.EngravingLength;
            }
        }

        /// <summary>
        /// Длина реза, мм
        /// </summary>
        public double CutLength
        {
            get { return _cutLength; }
            set
            {
                if (Math.Abs(_cutLength - value) > 0.0001)
                {
                    _cutLength = value;
                    OnPropertyChanged(nameof(CutLength));
                }
            }
        }

        /// <summary>
        /// Длина гравировки, мм
        /// </summary>
        public double EngravingLength
        {
            get { return _engravingLength; }
            set
            {
                if (Math.Abs(_engravingLength - value) > 0.0001)
                {
                    _engravingLength = value;
                    OnPropertyChanged(nameof(EngravingLength));
                }
            }
        }
    }

    /// <summary>
    /// Операция гибки
    /// </summary>
    public class BendingOperation : ManufacturingOperationBase
    {
        private double _bendAngle;
        private double _bendLength;

        public BendingOperation()
            : base(ManufacturingOperationType.Bending)
        {
        }

        protected override double CalculateUnitCost(PricingSettings settings)
        {
            return settings.BendingPricePerOperation;
        }

        public override void CopyParametersFrom(ManufacturingOperationBase source)
        {
            base.CopyParametersFrom(source);
            if (source is BendingOperation bend)
            {
                BendAngle = bend.BendAngle;
                BendLength = bend.BendLength;
            }
        }

        /// <summary>
        /// Угол гиба, град
        /// </summary>
        public double BendAngle
        {
            get { return _bendAngle; }
            set
            {
                if (Math.Abs(_bendAngle - value) > 0.0001)
                {
                    _bendAngle = value;
                    OnPropertyChanged(nameof(BendAngle));
                }
            }
        }

        /// <summary>
        /// Длина гиба, мм
        /// </summary>
        public double BendLength
        {
            get { return _bendLength; }
            set
            {
                if (Math.Abs(_bendLength - value) > 0.0001)
                {
                    _bendLength = value;
                    OnPropertyChanged(nameof(BendLength));
                }
            }
        }
    }

    /// <summary>
    /// Операция вальцовки
    /// </summary>
    public class RollingOperation : ManufacturingOperationBase
    {
        private double _rollDiameter;
        private double _radius;
        private double _length;

        public RollingOperation()
            : base(ManufacturingOperationType.Rolling)
        {
        }

        protected override double CalculateUnitCost(PricingSettings settings)
        {
            return PartMass * settings.RollingPricePerKg;
        }

        public override void CopyParametersFrom(ManufacturingOperationBase source)
        {
            base.CopyParametersFrom(source);
            if (source is RollingOperation roll)
            {
                RollDiameter = roll.RollDiameter;
                Radius = roll.Radius;
                Length = roll.Length;
            }
        }

        /// <summary>
        /// Масса детали, кг. Цена вальцовки зависит от массы; значение выставляется при пересчёте стоимости
        /// </summary>
        public double PartMass { get; set; }

        protected override bool IsUnitCostReliable => PartMass > 0;

        /// <summary>
        /// Диаметр вальцовки, мм
        /// </summary>
        public double RollDiameter
        {
            get { return _rollDiameter; }
            set
            {
                if (Math.Abs(_rollDiameter - value) > 0.0001)
                {
                    _rollDiameter = value;
                    OnPropertyChanged(nameof(RollDiameter));
                }
            }
        }

        /// <summary>
        /// Радиус, мм
        /// </summary>
        public double Radius
        {
            get { return _radius; }
            set
            {
                if (Math.Abs(_radius - value) > 0.0001)
                {
                    _radius = value;
                    OnPropertyChanged(nameof(Radius));
                }
            }
        }

        /// <summary>
        /// Длина, мм
        /// </summary>
        public double Length
        {
            get { return _length; }
            set
            {
                if (Math.Abs(_length - value) > 0.0001)
                {
                    _length = value;
                    OnPropertyChanged(nameof(Length));
                }
            }
        }
    }

    /// <summary>
    /// Операция отбортовки
    /// </summary>
    public class FlangingOperation : ManufacturingOperationBase
    {
        private double _diameter;
        private double _radius;

        public FlangingOperation()
            : base(ManufacturingOperationType.Flanging)
        {
        }

        protected override double CalculateUnitCost(PricingSettings settings)
        {
            return settings.FlangingPricePerOperation;
        }

        public override void CopyParametersFrom(ManufacturingOperationBase source)
        {
            base.CopyParametersFrom(source);
            if (source is FlangingOperation flange)
            {
                Diameter = flange.Diameter;
                Radius = flange.Radius;
            }
        }

        /// <summary>
        /// Диаметр отбортовки, мм
        /// </summary>
        public double Diameter
        {
            get { return _diameter; }
            set
            {
                if (Math.Abs(_diameter - value) > 0.0001)
                {
                    _diameter = value;
                    OnPropertyChanged(nameof(Diameter));
                }
            }
        }

        /// <summary>
        /// Радиус отбортовки, мм
        /// </summary>
        public double Radius
        {
            get { return _radius; }
            set
            {
                if (Math.Abs(_radius - value) > 0.0001)
                {
                    _radius = value;
                    OnPropertyChanged(nameof(Radius));
                }
            }
        }
    }

    /// <summary>
    /// Прочая операция, заданная пользователем: название и цена за единицу
    /// </summary>
    public class CustomOperation : ManufacturingOperationBase
    {
        private double _unitPrice;

        public CustomOperation()
            : base(ManufacturingOperationType.Custom)
        {
            Origin = OperationOrigin.Manual;
        }

        protected override double CalculateUnitCost(PricingSettings settings)
        {
            return UnitPrice;
        }

        protected override bool IsUnitCostReliable => UnitPrice > 0;

        public override void CopyParametersFrom(ManufacturingOperationBase source)
        {
            base.CopyParametersFrom(source);
            if (source is CustomOperation custom)
            {
                Name = custom.Name;
                UnitPrice = custom.UnitPrice;
            }
        }

        /// <summary>
        /// Цена за единицу, руб
        /// </summary>
        public double UnitPrice
        {
            get { return _unitPrice; }
            set
            {
                if (Math.Abs(_unitPrice - value) > 0.0001)
                {
                    _unitPrice = value;
                    OnPropertyChanged(nameof(UnitPrice));
                }
            }
        }
    }
}
