using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ConstructionControl
{
    public class ProjectObject
    {
        public Dictionary<string, MaterialDemand> Demand = new();

        public ObjectArchive Archive { get; set; } = new();

        public string Name { get; set; }

        // ===== НОВЫЕ НАСТРОЙКИ ОБЪЕКТА =====

        // Количество блоков
        public int BlocksCount { get; set; }

        // Есть ли подвал
        public bool HasBasement { get; set; }

        // Одинаковое количество этажей во всех блоках
        public bool SameFloorsInBlocks { get; set; } = true;

        // Если этажи одинаковые
        public int FloorsPerBlock { get; set; }

        // Если этажи разные (ключ = номер блока, значение = этажи)
        public Dictionary<int, int> FloorsByBlock { get; set; } = new();

        // ===== СТАРОЕ (НЕ ТРОГАЕМ) 3443=====

        public Dictionary<string, List<string>> MaterialNamesByGroup { get; set; } = new();

        public Dictionary<string, string> StbByGroup { get; set; } = new();
        public Dictionary<string, string> SupplierByGroup { get; set; } = new();

        public List<MaterialGroup> MaterialGroups { get; set; } = new();
        public List<ArrivalItem> ArrivalHistory { get; set; } = new();
        public List<string> SummaryVisibleGroups { get; set; } = new();
        public List<OtJournalEntry> OtJournal { get; set; } = new();
    }

    public class OtJournalEntry : INotifyPropertyChanged
    {
        private DateTime instructionDate = DateTime.Today;
        private string fullName;
        private string specialty;
        private string instructionType = "Первичный на рабочем месте";
        private string instructionNumbers;
        private int repeatPeriodMonths = 3;
        private bool isBrigadier;
        private string brigadierName;

        public DateTime InstructionDate
        {
            get => instructionDate;
            set
            {
                instructionDate = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NextRepeatDate));
                OnPropertyChanged(nameof(IsRepeatRequired));
            }
        }

        public string FullName
        {
            get => fullName;
            set => SetField(ref fullName, value);
        }

        public string Specialty
        {
            get => specialty;
            set => SetField(ref specialty, value);
        }

        public string InstructionType
        {
            get => instructionType;
            set => SetField(ref instructionType, value);
        }

        public string InstructionNumbers
        {
            get => instructionNumbers;
            set => SetField(ref instructionNumbers, value);
        }

        public int RepeatPeriodMonths
        {
            get => repeatPeriodMonths;
            set
            {
                if (SetField(ref repeatPeriodMonths, value))
                {
                    OnPropertyChanged(nameof(NextRepeatDate));
                    OnPropertyChanged(nameof(IsRepeatRequired));
                }
            }
        }

        public bool IsBrigadier
        {
            get => isBrigadier;
            set
            {
                if (SetField(ref isBrigadier, value))
                {
                    if (isBrigadier)
                        BrigadierName = null;
                    OnPropertyChanged(nameof(IsWorker));
                }
            }
        }

        public bool IsWorker => !IsBrigadier;

        public string BrigadierName
        {
            get => brigadierName;
            set => SetField(ref brigadierName, value);
        }

        public DateTime NextRepeatDate => InstructionDate.AddMonths(Math.Max(1, RepeatPeriodMonths));

        public bool IsRepeatRequired => DateTime.Today >= NextRepeatDate;

        public event PropertyChangedEventHandler PropertyChanged;

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class MaterialDemand
    {
        public string Unit;
        public Dictionary<int, Dictionary<int, double>> Floors; // block → floor → qty
    }


    public class MaterialGroup
    {
        public string Name { get; set; }
        public List<string> Items { get; set; } = new();
    }

    public class Arrival
    {
        public string Category { get; set; }     // Основные / Допы
        public string SubCategory { get; set; }  // Внутренние / Малоценка

        public string MaterialGroup { get; set; }
        public string TtnNumber { get; set; }
        public List<ArrivalItem> Items { get; set; } = new();
    }




    public class JournalRecord
    {
        public string SheetName { get; set; }

        public DateTime Date { get; set; }
        public string ObjectName { get; set; }

        public string Category { get; set; }     // Основные / Допы
        public string SubCategory { get; set; }  // Внутренние / Малоценка

        public string MaterialGroup { get; set; }
        public string MaterialName { get; set; }

        public string Unit { get; set; }
        public double Quantity { get; set; }
        public string Passport { get; set; }
        public string Ttn { get; set; }
        public string Stb { get; set; }
        public string Supplier { get; set; }
        public string Position { get; set; }
        public string Volume { get; set; }

    }
    public class ObjectArchive
    {
        public List<string> Groups { get; set; } = new();
        public Dictionary<string, List<string>> Materials { get; set; } = new();

        public List<string> Units { get; set; } = new();
        public List<string> Suppliers { get; set; } = new();
        public List<string> Passports { get; set; } = new();
        public List<string> Stb { get; set; } = new();
    }

    public class SummaryRow
    {
        public string MaterialName { get; set; }
        public string Unit { get; set; }

        public Dictionary<int, double> ByBlocks { get; set; } = new();
        public double Total { get; set; }
    }

    public class JvkDay
    {
        public DateTime Date { get; set; }
        public List<JvkTtn> Ttns { get; set; } = new();
    }
    public class JvkTtn
    {
        public string Ttn { get; set; }
        public string Supplier { get; set; }
        public string Unit { get; set; }
        public string Stb { get; set; }
        public string MaterialGroup { get; set; }
        public List<JvkPosition> Positions { get; set; } = new();
    }


    public class JvkPosition
    {
        public string Name { get; set; }
        public double Quantity { get; set; }
        public string Passport { get; set; }
    }




}

