using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ConstructionControl
{
    public class ArrivalItem : INotifyPropertyChanged
    {
        private DateTime date;
        private string materialName;
        private double quantity;
        private string unit;
        private string passport;
        private string stb;
        private string supplier;

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string prop = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        public DateTime Date
        {
            get => date;
            set { date = value; OnPropertyChanged(); }
        }

        public string MaterialName
        {
            get => materialName;
            set { materialName = value; OnPropertyChanged(); }
        }

        public double Quantity
        {
            get => quantity;
            set { quantity = value; OnPropertyChanged(); }
        }

        public string Unit
        {
            get => unit;
            set { unit = value; OnPropertyChanged(); }
        }

        public string Passport
        {
            get => passport;
            set { passport = value; OnPropertyChanged(); }
        }

        public string Stb
        {
            get => stb;
            set { stb = value; OnPropertyChanged(); }
        }

        public string Supplier
        {
            get => supplier;
            set { supplier = value; OnPropertyChanged(); }
        }

        public ObservableCollection<string> AvailableNames { get; set; }
            = new ObservableCollection<string>();

        public ObservableCollection<string> AvailableUnits { get; set; }
            = new ObservableCollection<string>();
    }
}
