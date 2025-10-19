using System.ComponentModel;

namespace TankManager.Core.Models
{
    public class MaterialInfo : INotifyPropertyChanged
    {
        private string _name;
        private double _totalMass;

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

        public double TotalMass
        {
            get { return _totalMass; }
            set
            {
                if (System.Math.Abs(_totalMass - value) > 0.0001)
                {
                    _totalMass = value;
                    OnPropertyChanged(nameof(TotalMass));
                    OnPropertyChanged(nameof(DisplayText));
                }
            }
        }

        public string DisplayText => $"{Name} ({TotalMass:F2} кг)";

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}