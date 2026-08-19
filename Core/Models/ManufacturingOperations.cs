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
        Flanging
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

        /// <summary>
        /// Тип операции
        /// </summary>
        public ManufacturingOperationType Type { get; }

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

        protected ManufacturingOperationBase(ManufacturingOperationType type)
        {
            Type = type;
            _name = string.Empty;
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
}
