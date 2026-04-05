using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

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
        public List<MaterialCatalogItem> MaterialCatalog { get; set; } = new();
        public Dictionary<string, string> MaterialTreeSplitRules { get; set; } = new();
        public List<ArrivalItem> ArrivalHistory { get; set; } = new();
        public List<string> SummaryVisibleGroups { get; set; } = new();
        public Dictionary<string, List<string>> SummaryMarksByGroup { get; set; } = new();
        public List<OtJournalEntry> OtJournal { get; set; } = new();
        public List<TimesheetPersonEntry> TimesheetPeople { get; set; } = new();
        public List<ProductionJournalEntry> ProductionJournal { get; set; } = new();
    }

    public class TimesheetPersonEntry : INotifyPropertyChanged
    {
        private Guid personId = Guid.NewGuid();
        private string fullName;
        private string specialty;
        private string rank;
        private string brigadeName;
        private bool isBrigadier;

        public Guid PersonId
        {
            get => personId;
            set => SetField(ref personId, value);
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

        public string Rank
        {
            get => rank;
            set => SetField(ref rank, value);
        }

        public string BrigadeName
        {
            get => brigadeName;
            set => SetField(ref brigadeName, value);
        }

        public bool IsBrigadier
        {
            get => isBrigadier;
            set => SetField(ref isBrigadier, value);
        }

        public List<TimesheetMonthEntry> Months { get; set; } = new();

        public event PropertyChangedEventHandler PropertyChanged;

        public string GetDayValue(string monthKey, int day)
        {
            var entry = GetOrCreateDayEntry(monthKey, day, createIfMissing: false);
            return entry?.Value?.Trim() ?? string.Empty;
        }

        public string GetDayComment(string monthKey, int day)
          => GetOrCreateDayEntry(monthKey, day, createIfMissing: false)?.Comment?.Trim() ?? string.Empty;

        public bool HasDayComment(string monthKey, int day)
            => !string.IsNullOrWhiteSpace(GetDayComment(monthKey, day));
        public bool? GetDocumentAccepted(string monthKey, int day)
              => GetOrCreateDayEntry(monthKey, day, createIfMissing: false)?.DocumentAccepted;

        public void SetDayComment(string monthKey, int day, string comment)
        {
            var entry = GetOrCreateDayEntry(monthKey, day, createIfMissing: true);
            if (entry == null)
                return;
            entry.Comment = string.IsNullOrWhiteSpace(comment) ? string.Empty : comment.Trim();
            OnPropertyChanged(nameof(Months));
        }

        public void SetDocumentAccepted(string monthKey, int day, bool? accepted)
        {
            var entry = GetOrCreateDayEntry(monthKey, day, createIfMissing: true);
            if (entry == null)
                return;

            entry.DocumentAccepted = accepted;
            OnPropertyChanged(nameof(Months));
        }

        public string GetPresenceMark(string monthKey, int day)
          => GetOrCreateDayEntry(monthKey, day, createIfMissing: false)?.PresenceMark ?? string.Empty;

        public bool GetPresenceChecked(string monthKey, int day)
            => string.Equals(GetPresenceMark(monthKey, day), "✔", StringComparison.Ordinal);

        public void SetPresenceMark(string monthKey, int day, string mark)
        {
            var entry = GetOrCreateDayEntry(monthKey, day, createIfMissing: true);
            if (entry == null)
                return;
            entry.PresenceMark = mark == "✔" ? mark : string.Empty;
            OnPropertyChanged(nameof(Months));
        }

        public void SetPresenceChecked(string monthKey, int day, bool isChecked)
        {
            var entry = GetOrCreateDayEntry(monthKey, day, createIfMissing: true);
            if (entry == null)
                return;

            entry.PresenceMark = isChecked ? "✔" : string.Empty;
            OnPropertyChanged(nameof(Months));
        }

        public void SetDayValue(string monthKey, int day, string value)
        {
            var entry = GetOrCreateDayEntry(monthKey, day, createIfMissing: true);
            if (entry == null)
                return;

            entry.Value = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
            OnPropertyChanged(nameof(Months));
        }


        public bool TryApplyDayValue(string monthKey, int day, string value, out string errorMessage)
        {
            errorMessage = null;
            if (string.IsNullOrWhiteSpace(monthKey) || day < 1 || day > 31)
                return false;
            SetDayValue(monthKey, day, value);
            return true;
        }

        public bool IsNonHourCode(string monthKey, int day)
        {
            var value = GetDayValue(monthKey, day);
            if (string.IsNullOrWhiteSpace(value))
                return false;

            return !(double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out _)
                     || double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _));
        }

        private TimesheetDayEntry GetOrCreateDayEntry(string monthKey, int day, bool createIfMissing)
        {
            if (string.IsNullOrWhiteSpace(monthKey) || day < 1 || day > 31)
                return null;

            var month = Months.FirstOrDefault(x => x.MonthKey == monthKey);
            if (month == null && createIfMissing)
            {
                month = new TimesheetMonthEntry { MonthKey = monthKey };
                Months.Add(month);
            }

            if (month == null)
                return null;
            // миграция старых данных
            if (!month.DayEntries.TryGetValue(day, out var entry) || entry == null)
            {
                if (month.DayValues.TryGetValue(day, out var oldValue) && !string.IsNullOrWhiteSpace(oldValue))
                {
                    entry = new TimesheetDayEntry { Value = oldValue.Trim() };
                    month.DayEntries[day] = entry;
                }
                else if (createIfMissing)
                {
                    entry = new TimesheetDayEntry();
                    month.DayEntries[day] = entry;
                }

                if (entry != null)
                    month.DayValues[day] = entry.Value ?? string.Empty;

                return entry;
            }
            if (entry != null)
                month.DayValues[day] = entry.Value ?? string.Empty;

            return entry;
        }
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

    public class TimesheetMonthEntry
    {
        public string MonthKey { get; set; }
        public Dictionary<int, string> DayValues { get; set; } = new();
        public Dictionary<int, TimesheetDayEntry> DayEntries { get; set; } = new();
    }

    public class TimesheetDayEntry
    {
        public string Value { get; set; }
        public string PresenceMark { get; set; }
        public string Comment { get; set; }
        public bool? DocumentAccepted { get; set; }
        public bool ArrivalMarked { get; set; } // ← ДОБАВЬ ЭТО
    }

    public class OtJournalEntry : INotifyPropertyChanged
    {
        private Guid personId = Guid.NewGuid();
        private DateTime instructionDate = DateTime.Today;
        private string fullName;
        private string specialty;
        private string rank;
        private string profession;
        private string instructionType = "Первичный на рабочем месте";
        private string instructionNumbers;
        private int repeatPeriodMonths = 3;
        private bool isBrigadier;
        private string brigadierName;
        private bool isDismissed;
        private bool isPendingRepeat;
        private bool isRepeatCompleted;

        public Guid PersonId
        {
            get => personId;
            set => SetField(ref personId, value);
        }

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
            set
            {
                if (SetField(ref fullName, value))
                {
                    OnPropertyChanged(nameof(FullNameDisplay));
                    OnPropertyChanged(nameof(LastName));
                }
            }
        }
        public string FullNameDisplay => string.IsNullOrWhiteSpace(FullName)
    ? string.Empty
    : string.Join(Environment.NewLine,
        FullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

        public string LastName => string.IsNullOrWhiteSpace(FullName)
            ? string.Empty
            : FullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;

        public string Specialty
        {
            get => specialty;
            set => SetField(ref specialty, value);
        }
        public string Rank
        {
            get => rank;
            set => SetField(ref rank, value);
        }

        public string Profession
        {
            get => profession;
            set => SetField(ref profession, value);
        }
        public string InstructionType
        {
            get => instructionType;
            set
            {
                if (SetField(ref instructionType, value))
                {
                    OnPropertyChanged(nameof(IsPrimaryInstruction));
                    OnPropertyChanged(nameof(IsActionEnabled));
                    OnPropertyChanged(nameof(StatusLabel));
                }
            }
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
        public bool IsDismissed
        {
            get => isDismissed;
            set => SetField(ref isDismissed, value);
        }

        public bool IsPendingRepeat
        {
            get => isPendingRepeat;
            set
            {
                if (SetField(ref isPendingRepeat, value))
                {
                    OnPropertyChanged(nameof(IsActionEnabled));
                    OnPropertyChanged(nameof(StatusLabel));
                }
            }
        }

        public bool IsRepeatCompleted
        {
            get => isRepeatCompleted;
            set
            {
                if (SetField(ref isRepeatCompleted, value))
                {
                    OnPropertyChanged(nameof(StatusLabel));
                }
            }
        }

        public bool IsPrimaryInstruction =>
            !string.IsNullOrWhiteSpace(InstructionType)
            && InstructionType.Contains("первич", StringComparison.CurrentCultureIgnoreCase);

        public bool IsActionEnabled => !IsPrimaryInstruction && IsPendingRepeat;

        public string StatusLabel => IsPendingRepeat
            ? "Требуется повторный"
            : IsRepeatCompleted
                ? "Повторный пройден"
                : string.Empty;

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
        public string Unit { get; set; }

        public Dictionary<int, Dictionary<string, double>> Levels { get; set; }
            = new Dictionary<int, Dictionary<string, double>>();

        public Dictionary<int, Dictionary<string, double>> MountedLevels { get; set; }
            = new Dictionary<int, Dictionary<string, double>>();

        public Dictionary<int, Dictionary<int, double>> Floors { get; set; }
            = new Dictionary<int, Dictionary<int, double>>();

        public Dictionary<int, Dictionary<int, double>> MountedFloors { get; set; }
            = new Dictionary<int, Dictionary<int, double>>();
    }
    public class MaterialCatalogItem
    {
        public string CategoryName { get; set; }
        public string TypeName { get; set; }
        public string SubTypeName { get; set; }
        public List<string> ExtraLevels { get; set; } = new();
        public List<string> LevelMarks { get; set; } = new();
        public string MaterialName { get; set; }
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

    public class ProductionJournalEntry : INotifyPropertyChanged
    {
        private DateTime date = DateTime.Today;
        private string actionName;
        private string workName;
        private string elementsText;
        private string blocksText;
        private string marksText;
        private string brigadeName;
        private string weather;
        private string deviations;
        private bool requiresHiddenWorkAct;
        private string remainingInfo;

        public DateTime Date
        {
            get => date;
            set => SetField(ref date, value);
        }

        public string ActionName
        {
            get => actionName;
            set => SetField(ref actionName, value);
        }

        public string WorkName
        {
            get => workName;
            set => SetField(ref workName, value);
        }

        public string ElementsText
        {
            get => elementsText;
            set => SetField(ref elementsText, value);
        }

        public string BlocksText
        {
            get => blocksText;
            set => SetField(ref blocksText, value);
        }

        public string MarksText
        {
            get => marksText;
            set => SetField(ref marksText, value);
        }

        public string BrigadeName
        {
            get => brigadeName;
            set => SetField(ref brigadeName, value);
        }

        public string Weather
        {
            get => weather;
            set => SetField(ref weather, value);
        }

        public string Deviations
        {
            get => deviations;
            set => SetField(ref deviations, value);
        }

        public bool RequiresHiddenWorkAct
        {
            get => requiresHiddenWorkAct;
            set => SetField(ref requiresHiddenWorkAct, value);
        }

        public string RemainingInfo
        {
            get => remainingInfo;
            set => SetField(ref remainingInfo, value);
        }

        public string WorkKey => $"{ActionName?.Trim()}::{WorkName?.Trim()}";

        public event PropertyChangedEventHandler PropertyChanged;

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            return true;
        }
    }

    public static class LevelMarkHelper
    {
        public static string GetLegacyMarkLabel(int floor)
            => floor == 0 ? "Подвал" : floor.ToString();

        public static List<string> GetDefaultMarks(ProjectObject projectObject)
        {
            var result = new List<string>();
            if (projectObject == null)
                return result;

            if (projectObject.HasBasement)
                result.Add(GetLegacyMarkLabel(0));

            var maxFloors = projectObject.SameFloorsInBlocks
                ? projectObject.FloorsPerBlock
                : projectObject.FloorsByBlock.Values.DefaultIfEmpty(0).Max();

            for (var floor = 1; floor <= maxFloors; floor++)
                result.Add(GetLegacyMarkLabel(floor));

            return result;
        }

        public static List<string> GetMarksForGroup(ProjectObject projectObject, string group)
        {
            var marks = new List<string>();
            if (projectObject == null)
                return marks;

            if (!string.IsNullOrWhiteSpace(group)
                && projectObject.SummaryMarksByGroup != null
                && projectObject.SummaryMarksByGroup.TryGetValue(group, out var configured)
                && configured != null)
            {
                marks.AddRange(configured.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()));
            }

            if (marks.Count == 0 && projectObject.MaterialCatalog != null)
            {
                marks.AddRange(projectObject.MaterialCatalog
                    .Where(x => string.Equals(x.TypeName ?? string.Empty, group ?? string.Empty, StringComparison.CurrentCultureIgnoreCase))
                    .SelectMany(x => x.LevelMarks ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim()));
            }

            if (marks.Count == 0 && projectObject.Demand != null)
            {
                foreach (var pair in projectObject.Demand)
                {
                    var parts = pair.Key.Split(new[] { "::" }, StringSplitOptions.None);
                    if (parts.Length != 2 || !string.Equals(parts[0], group ?? string.Empty, StringComparison.CurrentCultureIgnoreCase))
                        continue;

                    if (pair.Value?.Levels == null)
                        continue;

                    foreach (var block in pair.Value.Levels.Values)
                        marks.AddRange(block.Keys.Where(x => !string.IsNullOrWhiteSpace(x)));
                }
            }

            if (marks.Count == 0)
                marks.AddRange(GetDefaultMarks(projectObject));

            return marks
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public static List<int> ParseBlocks(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<int>();

            return Regex.Matches(text, @"\d+")
                .Cast<Match>()
                .Select(x => int.TryParse(x.Value, out var value) ? value : 0)
                .Where(x => x > 0)
                .Distinct()
                .ToList();
        }

        public static List<string> ParseMarks(string text) => SplitText(text);

        public static List<string> SplitText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<string>();

            return text
                .Split(new[] { ';', ',', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }
    }




}

