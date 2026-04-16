using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Text.Json.Serialization;

namespace ConstructionControl
{
    public class ProjectObject
    {
        public Dictionary<string, MaterialDemand> Demand { get; set; } = new();

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
        public Dictionary<int, string> BlockAxesByNumber { get; set; } = new();
        public string FullObjectName { get; set; } = string.Empty;
        public string GeneralContractorRepresentative { get; set; } = string.Empty;
        public string TechnicalSupervisorRepresentative { get; set; } = string.Empty;
        public string ProjectOrganizationRepresentative { get; set; } = string.Empty;
        public string ProjectDocumentationName { get; set; } = string.Empty;
        public List<string> MasterNames { get; set; } = new();
        public List<string> ForemanNames { get; set; } = new();
        public string ResponsibleForeman { get; set; } = string.Empty;
        public string SiteManagerName { get; set; } = string.Empty;

        // ===== СТАРОЕ (НЕ ТРОГАЕМ) 3443=====

        public Dictionary<string, List<string>> MaterialNamesByGroup { get; set; } = new();

        public Dictionary<string, string> StbByGroup { get; set; } = new();
        public Dictionary<string, string> SupplierByGroup { get; set; } = new();

        public List<MaterialGroup> MaterialGroups { get; set; } = new();
        public List<MaterialCatalogItem> MaterialCatalog { get; set; } = new();
        public Dictionary<string, string> MaterialTreeSplitRules { get; set; } = new();
        public List<string> AutoSplitMaterialNames { get; set; } = new();
        public List<ArrivalItem> ArrivalHistory { get; set; } = new();
        public List<string> SummaryVisibleGroups { get; set; } = new();
        public Dictionary<string, List<string>> SummaryMarksByGroup { get; set; } = new();
        public Dictionary<string, string> OtInstructionNumbersByProfession { get; set; } = new();
        public Dictionary<string, List<string>> ProductionDeviationsByType { get; set; } = new();
        public Dictionary<string, List<string>> ProductionWorksByAction { get; set; } = new();
        public Dictionary<string, List<string>> ProductionMaterialsByType { get; set; } = new();
        public Dictionary<string, string> HiddenWorkTitlePrefixReplacements { get; set; } = new();
        public List<OtJournalEntry> OtJournal { get; set; } = new();
        public List<TimesheetPersonEntry> TimesheetPeople { get; set; } = new();
        public List<ProductionJournalEntry> ProductionJournal { get; set; } = new();
        public ProductionAutoFillSettings ProductionAutoFillSettings { get; set; } = new();
        public List<ProductionAutoFillProfile> ProductionAutoFillProfiles { get; set; } = new();
        public string SelectedProductionAutoFillProfileName { get; set; } = string.Empty;
        public List<ProductionJournalTemplate> ProductionTemplates { get; set; } = new();
        public List<ProductionWorkRule> ProductionWorkRules { get; set; } = new();
        public List<InspectionJournalEntry> InspectionJournal { get; set; } = new();
        public List<InspectionJournalTemplate> InspectionTemplates { get; set; } = new();
        public List<ProjectNoteEntry> Notes { get; set; } = new();
        public List<DocumentTreeNode> PdfDocuments { get; set; } = new();
        public List<DocumentTreeNode> EstimateDocuments { get; set; } = new();
        public List<ArrivalFilterTemplate> ArrivalFilterTemplates { get; set; } = new();
        public List<SummaryBalanceHistoryEntry> SummaryBalanceHistory { get; set; } = new();
        public ProjectUiSettings UiSettings { get; set; } = new();
        public List<ProjectChangeLogEntry> ChangeLog { get; set; } = new();
        public HiddenWorkActDefaults HiddenWorkDefaults { get; set; } = new();
        public List<HiddenWorkActRecord> HiddenWorkActs { get; set; } = new();
        public List<HiddenWorkMaterialPreset> HiddenWorkMaterialPresets { get; set; } = new();
    }

    public class ProjectChangeLogEntry
    {
        public DateTime TimestampUtc { get; set; } = DateTime.UtcNow;
        public string UserName { get; set; } = string.Empty;
        public string Action { get; set; } = string.Empty;
        public string Details { get; set; } = string.Empty;
    }

    public class ProductionAutoFillSettings
    {
        public int MinQuantityPerRow { get; set; } = 4;
        public int MaxQuantityPerRow { get; set; } = 8;
        public int MinRowsPerRun { get; set; } = 4;
        public int TargetRowsPerRun { get; set; } = 5;
        public int MaxRowsPerRun { get; set; } = 6;
        public int MaxItemsPerRow { get; set; } = 2;
        public bool PreferSelectedTypeOnly { get; set; } = true;
        public bool UseBalancedDistribution { get; set; } = true;
        public bool PreferDemandDeficit { get; set; } = true;
        public bool RespectSelectedBlocksAndMarks { get; set; } = true;
        public bool AllowMixedMaterialsInRow { get; set; } = true;
    }

    public class ProductionAutoFillProfile
    {
        public string Name { get; set; } = string.Empty;
        public ProductionAutoFillSettings Settings { get; set; } = new();
    }

    public class ProductionJournalTemplate
    {
        public string Name { get; set; } = string.Empty;
        public string ActionName { get; set; } = string.Empty;
        public string WorkName { get; set; } = string.Empty;
        public string ElementsText { get; set; } = string.Empty;
        public string BlocksText { get; set; } = string.Empty;
        public string MarksText { get; set; } = string.Empty;
        public string BrigadeName { get; set; } = string.Empty;
        public string Weather { get; set; } = string.Empty;
        public string Deviations { get; set; } = string.Empty;
        public bool RequiresHiddenWorkAct { get; set; }
        public bool AllowCustomElements { get; set; }
        public bool IgnorePhotoRule { get; set; }
    }

    public class ProductionWorkRule
    {
        public string WorkName { get; set; } = string.Empty;
        public bool AllowCustomElements { get; set; }
        public bool IgnorePhotoRule { get; set; }
    }

    public class InspectionJournalTemplate
    {
        public string Name { get; set; } = string.Empty;
        public string JournalName { get; set; } = string.Empty;
        public string InspectionName { get; set; } = string.Empty;
        public int ReminderPeriodDays { get; set; } = 7;
        public string Notes { get; set; } = string.Empty;
    }

    public class HiddenWorkMaterialPreset
    {
        public string WorkTemplateKey { get; set; } = string.Empty;
        public List<string> MaterialNames { get; set; } = new();
        public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;
    }

    public class HiddenWorkActDefaults
    {
        public string FullObjectName { get; set; } = string.Empty;
        public string GeneralContractorInfo { get; set; } = string.Empty;
        public string SubcontractorInfo { get; set; } = string.Empty;
        public string TechnicalSupervisorInfo { get; set; } = string.Empty;
        public string ProjectOrganizationInfo { get; set; } = string.Empty;
        public string WorkExecutorInfo { get; set; } = string.Empty;
        public string ProjectDocumentation { get; set; } = string.Empty;
        public string Deviations { get; set; } = string.Empty;
        public string ContractorSignerName { get; set; } = string.Empty;
        public string TechnicalSupervisorSignerName { get; set; } = string.Empty;
        public string ProjectOrganizationSignerName { get; set; } = string.Empty;
    }

    public class HiddenWorkActMaterialEntry : INotifyPropertyChanged
    {
        private bool isSelected = true;
        private string materialName = string.Empty;
        private string passport = string.Empty;
        private DateTime? arrivalDate;

        public bool IsSelected
        {
            get => isSelected;
            set => SetField(ref isSelected, value);
        }

        public string MaterialName
        {
            get => materialName;
            set => SetField(ref materialName, value ?? string.Empty);
        }

        public string Passport
        {
            get => passport;
            set => SetField(ref passport, value ?? string.Empty);
        }

        public DateTime? ArrivalDate
        {
            get => arrivalDate;
            set
            {
                if (SetField(ref arrivalDate, value))
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrivalDateText)));
            }
        }

        [JsonIgnore]
        public string ArrivalDateText => ArrivalDate?.ToString("dd.MM.yyyy") ?? string.Empty;

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

    public class HiddenWorkActRecord : INotifyPropertyChanged
    {
        private Guid id = Guid.NewGuid();
        private string groupKey = string.Empty;
        private string workTemplateKey = string.Empty;
        private string actionName = string.Empty;
        private string workName = string.Empty;
        private string blocksText = string.Empty;
        private string marksText = string.Empty;
        private string workTitle = string.Empty;
        private string fullObjectName = string.Empty;
        private string generalContractorInfo = string.Empty;
        private string subcontractorInfo = string.Empty;
        private string technicalSupervisorInfo = string.Empty;
        private string projectOrganizationInfo = string.Empty;
        private string workExecutorInfo = string.Empty;
        private string projectDocumentation = string.Empty;
        private string deviations = string.Empty;
        private string contractorSignerName = string.Empty;
        private string technicalSupervisorSignerName = string.Empty;
        private string projectOrganizationSignerName = string.Empty;
        private DateTime startDate = DateTime.Today;
        private DateTime endDate = DateTime.Today;
        private bool isFixed;
        private bool isPrinted;
        private ObservableCollection<HiddenWorkActMaterialEntry> materials = new();

        public Guid Id
        {
            get => id;
            set => SetField(ref id, value == Guid.Empty ? Guid.NewGuid() : value);
        }

        public string GroupKey
        {
            get => groupKey;
            set => SetField(ref groupKey, value ?? string.Empty);
        }

        public string WorkTemplateKey
        {
            get => workTemplateKey;
            set => SetField(ref workTemplateKey, value ?? string.Empty);
        }

        public string ActionName
        {
            get => actionName;
            set => SetField(ref actionName, value ?? string.Empty);
        }

        public string WorkName
        {
            get => workName;
            set => SetField(ref workName, value ?? string.Empty);
        }

        public string BlocksText
        {
            get => blocksText;
            set => SetField(ref blocksText, value ?? string.Empty);
        }

        public string MarksText
        {
            get => marksText;
            set => SetField(ref marksText, value ?? string.Empty);
        }

        public string WorkTitle
        {
            get => workTitle;
            set
            {
                if (SetField(ref workTitle, value ?? string.Empty))
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TitleDisplay)));
            }
        }

        public string FullObjectName
        {
            get => fullObjectName;
            set => SetField(ref fullObjectName, value ?? string.Empty);
        }

        public string GeneralContractorInfo
        {
            get => generalContractorInfo;
            set => SetField(ref generalContractorInfo, value ?? string.Empty);
        }

        public string SubcontractorInfo
        {
            get => subcontractorInfo;
            set => SetField(ref subcontractorInfo, value ?? string.Empty);
        }

        public string TechnicalSupervisorInfo
        {
            get => technicalSupervisorInfo;
            set => SetField(ref technicalSupervisorInfo, value ?? string.Empty);
        }

        public string ProjectOrganizationInfo
        {
            get => projectOrganizationInfo;
            set => SetField(ref projectOrganizationInfo, value ?? string.Empty);
        }

        public string WorkExecutorInfo
        {
            get => workExecutorInfo;
            set => SetField(ref workExecutorInfo, value ?? string.Empty);
        }

        public string ProjectDocumentation
        {
            get => projectDocumentation;
            set => SetField(ref projectDocumentation, value ?? string.Empty);
        }

        public string Deviations
        {
            get => deviations;
            set => SetField(ref deviations, value ?? string.Empty);
        }

        public string ContractorSignerName
        {
            get => contractorSignerName;
            set => SetField(ref contractorSignerName, value ?? string.Empty);
        }

        public string TechnicalSupervisorSignerName
        {
            get => technicalSupervisorSignerName;
            set => SetField(ref technicalSupervisorSignerName, value ?? string.Empty);
        }

        public string ProjectOrganizationSignerName
        {
            get => projectOrganizationSignerName;
            set => SetField(ref projectOrganizationSignerName, value ?? string.Empty);
        }

        public DateTime StartDate
        {
            get => startDate;
            set
            {
                if (SetField(ref startDate, value.Date))
                {
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StartDateText)));
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PeriodDisplay)));
                }
            }
        }

        public DateTime EndDate
        {
            get => endDate;
            set
            {
                if (SetField(ref endDate, value.Date))
                {
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EndDateText)));
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PeriodDisplay)));
                }
            }
        }

        public bool IsFixed
        {
            get => isFixed;
            set
            {
                if (SetField(ref isFixed, value))
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StateDisplay)));
            }
        }

        public bool IsPrinted
        {
            get => isPrinted;
            set
            {
                if (SetField(ref isPrinted, value))
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StateDisplay)));
            }
        }

        public ObservableCollection<HiddenWorkActMaterialEntry> Materials
        {
            get => materials;
            set
            {
                if (value == null)
                    value = new ObservableCollection<HiddenWorkActMaterialEntry>();

                if (ReferenceEquals(materials, value))
                    return;

                materials = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Materials)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MaterialsSummary)));
            }
        }

        [JsonIgnore]
        public string StartDateText => StartDate.ToString("dd.MM.yyyy");

        [JsonIgnore]
        public string EndDateText => EndDate.ToString("dd.MM.yyyy");

        [JsonIgnore]
        public string PeriodDisplay => $"{StartDate:dd.MM.yyyy} - {EndDate:dd.MM.yyyy}";

        [JsonIgnore]
        public string TitleDisplay => LevelMarkHelper.PreventSingleLetterWrap(WorkTitle ?? string.Empty);

        [JsonIgnore]
        public string StateDisplay
            => IsPrinted
                ? "Распечатан"
                : IsFixed
                    ? "Зафиксирован"
                    : "Черновик";

        [JsonIgnore]
        public string MaterialsSummary => string.Join(", ",
            (Materials ?? new ObservableCollection<HiddenWorkActMaterialEntry>())
                .Where(x => x != null && x.IsSelected && !string.IsNullOrWhiteSpace(x.MaterialName))
                .Select(x => x.MaterialName.Trim()));

        public void NotifyMaterialsChanged()
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MaterialsSummary)));
        }

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

    public class ProjectUiSettings
    {
        public bool DisableTree { get; set; }
        public bool PinTreeByDefault { get; set; }
        public bool PinJournalPanelsByDefault { get; set; }
        public bool ShowReminderPopup { get; set; } = true;
        public string ReminderPresentationMode { get; set; } = ReminderPresentationModes.Overlay;
        public int ReminderSnoozeMinutes { get; set; } = 15;
        public int AutoSaveIntervalMinutes { get; set; } = 5;
        public bool HideReminderDetails { get; set; }
        public bool SafeStartupMode { get; set; }
        public string UiDensityMode { get; set; } = "Стандартный";
        public string UiThemeMode { get; set; } = UiThemeModes.Light;
        public string AccessRole { get; set; } = ProjectAccessRoles.Critical;
        public bool RequireCodeForCriticalOperations { get; set; } = true;
        public bool SummaryReminderOnOverage { get; set; } = true;
        public bool SummaryReminderOnDeficit { get; set; }
        public bool SummaryReminderOnlyMain { get; set; } = true;
        public bool AutoFitCurrentTabColumns { get; set; } = true;
        public string DataRootDirectory { get; set; } = string.Empty;
        public string PreferredPdfEditorPath { get; set; } = string.Empty;
        public string PreferredSpreadsheetEditorPath { get; set; } = string.Empty;
        public bool CheckUpdatesOnStart { get; set; }
        public string UpdateFeedUrl { get; set; } = string.Empty;
        public string OtStatusFilter { get; set; } = "Все";
        public string OtSpecialtyFilter { get; set; } = "Все";
        public string OtBrigadeFilter { get; set; } = "Все";
        public Dictionary<string, string> CommandPaletteShortcuts { get; set; } = new(StringComparer.CurrentCultureIgnoreCase);
        public Dictionary<string, string> TabDisplayModes { get; set; } = new(StringComparer.CurrentCultureIgnoreCase);
        public Dictionary<string, List<GridColumnPreference>> GridColumnPreferences { get; set; } = new(StringComparer.CurrentCultureIgnoreCase);
        public Dictionary<string, Dictionary<string, List<GridColumnPreference>>> GridColumnPresets { get; set; }
            = new(StringComparer.CurrentCultureIgnoreCase);
    }

    public static class ProjectAccessRoles
    {
        public const string View = "view";
        public const string Edit = "edit";
        public const string Critical = "critical";

        public static readonly string[] All =
        {
            View,
            Edit,
            Critical
        };

        public static string ToDisplay(string role)
        {
            return (role ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                View => "Просмотр",
                Edit => "Редактирование",
                _ => "Критические операции"
            };
        }
    }

    public static class UiThemeModes
    {
        public const string Light = "light";
        public const string Dark = "dark";
        public const string System = "system";
    }

    public static class ReminderPresentationModes
    {
        public const string Overlay = "overlay";
        public const string Tabs = "tabs";
        public const string Combined = "combined";
    }

    public class GridColumnPreference
    {
        public string Header { get; set; } = string.Empty;
        public bool IsVisible { get; set; } = true;
        public int DisplayIndex { get; set; }
        public double Width { get; set; } = double.NaN;
    }

    public class ReminderSectionViewModel
    {
        public string Header { get; set; }
        public List<string> Items { get; set; } = new();
    }

    public class DocumentTreeNode : INotifyPropertyChanged
    {
        private string name;
        private string filePath;
        private string storedRelativePath;
        private string contentHash;
        private long? fileSizeBytes;
        private DateTime? hashVerifiedAtUtc;
        private int previewPage = 1;
        private double previewZoom = 100;
        private int previewScrollX;
        private int previewScrollY;
        private string previewSheetName;
        private bool isFolder;

        public string Name
        {
            get => name;
            set => SetField(ref name, value);
        }

        public string FilePath
        {
            get => filePath;
            set => SetField(ref filePath, value);
        }

        public string StoredRelativePath
        {
            get => storedRelativePath;
            set => SetField(ref storedRelativePath, value);
        }

        public string ContentHash
        {
            get => contentHash;
            set => SetField(ref contentHash, value);
        }

        public long? FileSizeBytes
        {
            get => fileSizeBytes;
            set => SetField(ref fileSizeBytes, value);
        }

        public DateTime? HashVerifiedAtUtc
        {
            get => hashVerifiedAtUtc;
            set => SetField(ref hashVerifiedAtUtc, value);
        }

        public int PreviewPage
        {
            get => previewPage <= 0 ? 1 : previewPage;
            set => SetField(ref previewPage, value <= 0 ? 1 : value);
        }

        public double PreviewZoom
        {
            get => previewZoom <= 0 ? 100 : previewZoom;
            set => SetField(ref previewZoom, value <= 0 ? 100 : value);
        }

        public int PreviewScrollX
        {
            get => previewScrollX;
            set => SetField(ref previewScrollX, Math.Max(0, value));
        }

        public int PreviewScrollY
        {
            get => previewScrollY;
            set => SetField(ref previewScrollY, Math.Max(0, value));
        }

        public string PreviewSheetName
        {
            get => previewSheetName;
            set => SetField(ref previewSheetName, value);
        }

        public bool IsFolder
        {
            get => isFolder;
            set => SetField(ref isFolder, value);
        }

        public List<DocumentTreeNode> Children { get; set; } = new();

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

    public class TimesheetPersonEntry : INotifyPropertyChanged
    {
        private Guid personId = Guid.NewGuid();
        private string fullName;
        private string specialty;
        private string rank;
        private string brigadeName;
        private bool isBrigadier;
        private int dailyWorkHours = 8;

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

        public int DailyWorkHours
        {
            get => dailyWorkHours;
            set
            {
                var normalized = Math.Clamp(value, 1, 24);
                SetField(ref dailyWorkHours, normalized);
            }
        }

        public List<TimesheetMonthEntry> Months { get; set; } = new();
        public List<TimesheetMonthEntry> ArchivedMonths { get; set; } = new();

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
        private bool isScheduledRepeat;

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
            set
            {
                if (SetField(ref isDismissed, value))
                {
                    OnPropertyChanged(nameof(IsActionEnabled));
                    OnPropertyChanged(nameof(StatusLabel));
                }
            }
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

        public bool IsScheduledRepeat
        {
            get => isScheduledRepeat;
            set
            {
                if (SetField(ref isScheduledRepeat, value))
                    OnPropertyChanged(nameof(StatusLabel));
            }
        }

        public bool IsPrimaryInstruction =>
            !string.IsNullOrWhiteSpace(InstructionType)
            && InstructionType.Contains("первич", StringComparison.CurrentCultureIgnoreCase);

        public bool IsActionEnabled => IsPendingRepeat && !IsDismissed;

        public string StatusLabel => IsDismissed
            ? "Снят с объекта"
            : IsPrimaryInstruction && IsPendingRepeat
                ? "Требуется первичный"
                : IsPendingRepeat
                    ? "Требуется повторный"
                    : IsPrimaryInstruction && IsRepeatCompleted
                        ? "Первичный пройден"
                        : IsRepeatCompleted
                            ? "Повторный пройден"
                            : IsScheduledRepeat
                                ? "Запланирован"
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
        public bool IsAnomaly { get; set; }
        public string AnomalyText { get; set; }

    }

    public class ArrivalFilterTemplate
    {
        public string Name { get; set; } = string.Empty;
        public List<string> SelectedTypes { get; set; } = new();
        public List<string> SelectedNames { get; set; } = new();
        public bool ShowMain { get; set; } = true;
        public bool ShowExtra { get; set; } = true;
        public bool ShowLowCost { get; set; } = true;
        public bool ShowInternal { get; set; } = true;
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public string SearchText { get; set; } = string.Empty;
    }

    public class SummaryBalanceHistoryEntry
    {
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public string Group { get; set; } = string.Empty;
        public string Material { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public double Quantity { get; set; }
        public bool IsOverage { get; set; }
        public string Reason { get; set; } = string.Empty;
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
        private string blocksDisplayText;
        private bool suppressDateDisplay;
        private bool suppressWeatherDisplay;
        private bool isAutoCorrectedQuantity;
        private bool isGeneratedCompanion;
        private bool armoringCompanionRequested;
        private bool armoringPromptHandled;
        private bool allowCustomElements;
        private bool ignorePhotoRule;

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

        [JsonIgnore]
        public string BlocksDisplayText
        {
            get => string.IsNullOrWhiteSpace(blocksDisplayText) ? (BlocksText ?? string.Empty) : blocksDisplayText;
            set => SetField(ref blocksDisplayText, value);
        }

        public bool SuppressDateDisplay
        {
            get => suppressDateDisplay;
            set => SetField(ref suppressDateDisplay, value);
        }

        public bool SuppressWeatherDisplay
        {
            get => suppressWeatherDisplay;
            set => SetField(ref suppressWeatherDisplay, value);
        }

        public bool IsAutoCorrectedQuantity
        {
            get => isAutoCorrectedQuantity;
            set => SetField(ref isAutoCorrectedQuantity, value);
        }

        public bool IsGeneratedCompanion
        {
            get => isGeneratedCompanion;
            set => SetField(ref isGeneratedCompanion, value);
        }

        public bool ArmoringCompanionRequested
        {
            get => armoringCompanionRequested;
            set => SetField(ref armoringCompanionRequested, value);
        }

        public bool ArmoringPromptHandled
        {
            get => armoringPromptHandled;
            set => SetField(ref armoringPromptHandled, value);
        }

        public bool AllowCustomElements
        {
            get => allowCustomElements;
            set
            {
                if (SetField(ref allowCustomElements, value))
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StatusDisplay)));
            }
        }

        public bool IgnorePhotoRule
        {
            get => ignorePhotoRule;
            set
            {
                if (SetField(ref ignorePhotoRule, value))
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StatusDisplay)));
            }
        }

        public string DateDisplay => SuppressDateDisplay ? string.Empty : Date.ToString("dd.MM.yyyy");

        public string WeatherDisplay => SuppressWeatherDisplay ? string.Empty : LevelMarkHelper.PreventSingleLetterWrap(Weather ?? string.Empty);

        public string ElementsDisplay => string.Join(Environment.NewLine, LevelMarkHelper.SplitText(ElementsText));

        public string DeviationsDisplay => string.Join(Environment.NewLine, LevelMarkHelper.SplitText(Deviations));

        public string StatusDisplay
        {
            get
            {
                var parts = new List<string>
                {
                    IsGeneratedCompanion
                        ? "Автосоздано (компаньон)"
                        : (IsAutoCorrectedQuantity ? "Скорректировано по остатку" : "Норма")
                };

                if (AllowCustomElements)
                    parts.Add("свои элементы");

                return string.Join(" | ", parts);
            }
        }

        public string WorkKey => $"{ActionName?.Trim()}::{WorkName?.Trim()}";

        public event PropertyChangedEventHandler PropertyChanged;

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

            if (propertyName == nameof(Date) || propertyName == nameof(SuppressDateDisplay))
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DateDisplay)));

            if (propertyName == nameof(Weather) || propertyName == nameof(SuppressWeatherDisplay))
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(WeatherDisplay)));

            if (propertyName == nameof(ElementsText))
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ElementsDisplay)));

            if (propertyName == nameof(Deviations))
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DeviationsDisplay)));

            if (propertyName == nameof(IsAutoCorrectedQuantity)
                || propertyName == nameof(IsGeneratedCompanion)
                || propertyName == nameof(AllowCustomElements))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StatusDisplay)));
            }

            return true;
        }
    }

    public class InspectionJournalEntry : INotifyPropertyChanged
    {
        private string journalName;
        private string inspectionName;
        private DateTime reminderStartDate = DateTime.Today;
        private int reminderPeriodDays = 7;
        private DateTime? lastCompletedDate;
        private string notes;
        private bool isCompletionHistory;

        public string JournalName
        {
            get => journalName;
            set => SetField(ref journalName, value);
        }

        public string InspectionName
        {
            get => inspectionName;
            set => SetField(ref inspectionName, value);
        }

        public string JournalDisplay => LevelMarkHelper.PreventSingleLetterWrap(JournalName ?? string.Empty);

        public string InspectionDisplay => LevelMarkHelper.PreventSingleLetterWrap(InspectionName ?? string.Empty);

        public string NotesDisplay => LevelMarkHelper.PreventSingleLetterWrap(Notes ?? string.Empty);

        public DateTime ReminderStartDate
        {
            get => reminderStartDate;
            set => SetField(ref reminderStartDate, value);
        }

        public int ReminderPeriodDays
        {
            get => reminderPeriodDays;
            set => SetField(ref reminderPeriodDays, value <= 0 ? 1 : value);
        }

        public DateTime? LastCompletedDate
        {
            get => lastCompletedDate;
            set => SetField(ref lastCompletedDate, value);
        }

        public string Notes
        {
            get => notes;
            set => SetField(ref notes, value);
        }

        public bool IsCompletionHistory
        {
            get => isCompletionHistory;
            set => SetField(ref isCompletionHistory, value);
        }

        public DateTime NextReminderDate
        {
            get
            {
                var period = ReminderPeriodDays <= 0 ? 1 : ReminderPeriodDays;
                if (LastCompletedDate.HasValue)
                {
                    var fromLast = LastCompletedDate.Value.Date.AddDays(period);
                    return fromLast > ReminderStartDate.Date ? fromLast : ReminderStartDate.Date;
                }

                return ReminderStartDate.Date;
            }
        }

        public bool IsDue => !IsCompletionHistory && DateTime.Today >= NextReminderDate.Date;

        public bool IsUpcoming => !IsCompletionHistory && !IsDue;

        public string ReminderStatus => IsCompletionHistory
            ? (LastCompletedDate.HasValue
                ? $"Проведен: {LastCompletedDate.Value:dd.MM.yyyy}"
                : "Проведен")
            : (IsDue
                ? $"Нужно провести с {NextReminderDate:dd.MM.yyyy}"
                : $"Следующий: {NextReminderDate:dd.MM.yyyy}");

        public string ReminderStatusDisplay => LevelMarkHelper.PreventSingleLetterWrap(ReminderStatus);

        public event PropertyChangedEventHandler PropertyChanged;

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

            if (propertyName == nameof(JournalName))
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(JournalDisplay)));

            if (propertyName == nameof(InspectionName))
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(InspectionDisplay)));

            if (propertyName == nameof(Notes))
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NotesDisplay)));

            if (propertyName != nameof(ReminderStatus)
                && propertyName != nameof(ReminderStatusDisplay)
                && propertyName != nameof(NextReminderDate)
                && propertyName != nameof(IsDue)
                && propertyName != nameof(IsUpcoming))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NextReminderDate)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsDue)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsUpcoming)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReminderStatus)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReminderStatusDisplay)));
            }
            return true;
        }
    }

    public static class NoteReminderModes
    {
        public const string None = "none";
        public const string Once = "once";
        public const string Periodic = "periodic";
        public const string Persistent = "persistent";

        public static readonly string[] All =
        {
            None,
            Once,
            Periodic,
            Persistent
        };

        public static string Normalize(string mode)
        {
            return (mode ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                Once => Once,
                Periodic => Periodic,
                Persistent => Persistent,
                _ => None
            };
        }

        public static string ToDisplay(string mode)
        {
            return Normalize(mode) switch
            {
                Once => "Разово",
                Periodic => "Периодически",
                Persistent => "Постоянно",
                _ => "Без напоминания"
            };
        }
    }

    public static class NoteReminderIntervalUnits
    {
        public const string Minutes = "minutes";
        public const string Hours = "hours";
        public const string Days = "days";

        public static readonly string[] All =
        {
            Minutes,
            Hours,
            Days
        };

        public static string Normalize(string unit)
        {
            return (unit ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                Minutes => Minutes,
                Hours => Hours,
                _ => Days
            };
        }

        public static string ToDisplay(string unit)
        {
            return Normalize(unit) switch
            {
                Minutes => "Минуты",
                Hours => "Часы",
                _ => "Дни"
            };
        }

        public static TimeSpan ToTimeSpan(string unit, int value)
        {
            var normalizedValue = Math.Max(1, value);
            return Normalize(unit) switch
            {
                Minutes => TimeSpan.FromMinutes(normalizedValue),
                Hours => TimeSpan.FromHours(normalizedValue),
                _ => TimeSpan.FromDays(normalizedValue)
            };
        }
    }

    public class ProjectNoteEntry : INotifyPropertyChanged
    {
        private Guid id = Guid.NewGuid();
        private string title = string.Empty;
        private string body = string.Empty;
        private DateTime createdAtUtc = DateTime.UtcNow;
        private DateTime updatedAtUtc = DateTime.UtcNow;
        private bool isDone;
        private string reminderMode = NoteReminderModes.None;
        private DateTime? reminderStart;
        private int reminderIntervalValue = 1;
        private string reminderIntervalUnit = NoteReminderIntervalUnits.Days;
        private DateTime? reminderLastAcknowledgedAt;

        public Guid Id
        {
            get => id;
            set => SetField(ref id, value == Guid.Empty ? Guid.NewGuid() : value);
        }

        public string Title
        {
            get => title;
            set => SetField(ref title, value ?? string.Empty);
        }

        public string Body
        {
            get => body;
            set => SetField(ref body, value ?? string.Empty);
        }

        public DateTime CreatedAtUtc
        {
            get => createdAtUtc;
            set => SetField(ref createdAtUtc, value == default ? DateTime.UtcNow : DateTime.SpecifyKind(value, DateTimeKind.Utc));
        }

        public DateTime UpdatedAtUtc
        {
            get => updatedAtUtc;
            set => SetField(ref updatedAtUtc, value == default ? DateTime.UtcNow : DateTime.SpecifyKind(value, DateTimeKind.Utc));
        }

        public bool IsDone
        {
            get => isDone;
            set => SetField(ref isDone, value);
        }

        public string ReminderMode
        {
            get => reminderMode;
            set => SetField(ref reminderMode, NoteReminderModes.Normalize(value));
        }

        public DateTime? ReminderStart
        {
            get => reminderStart;
            set => SetField(ref reminderStart, value?.Kind == DateTimeKind.Unspecified ? DateTime.SpecifyKind(value.Value, DateTimeKind.Local) : value);
        }

        public int ReminderIntervalValue
        {
            get => reminderIntervalValue;
            set => SetField(ref reminderIntervalValue, Math.Max(1, value));
        }

        public string ReminderIntervalUnit
        {
            get => reminderIntervalUnit;
            set => SetField(ref reminderIntervalUnit, NoteReminderIntervalUnits.Normalize(value));
        }

        public DateTime? ReminderLastAcknowledgedAt
        {
            get => reminderLastAcknowledgedAt;
            set => SetField(ref reminderLastAcknowledgedAt, value);
        }

        [JsonIgnore]
        public string TitleDisplay => LevelMarkHelper.PreventSingleLetterWrap(Title ?? string.Empty);

        [JsonIgnore]
        public string BodyDisplay => LevelMarkHelper.PreventSingleLetterWrap(Body ?? string.Empty);

        [JsonIgnore]
        public DateTime CreatedAtLocal => DateTime.SpecifyKind(CreatedAtUtc, DateTimeKind.Utc).ToLocalTime();

        [JsonIgnore]
        public DateTime UpdatedAtLocal => DateTime.SpecifyKind(UpdatedAtUtc, DateTimeKind.Utc).ToLocalTime();

        [JsonIgnore]
        public string ReminderModeDisplay => NoteReminderModes.ToDisplay(ReminderMode);

        [JsonIgnore]
        public string ReminderIntervalDisplay
        {
            get
            {
                var mode = NoteReminderModes.Normalize(ReminderMode);
                if (mode != NoteReminderModes.Periodic && mode != NoteReminderModes.Persistent)
                    return string.Empty;

                var value = Math.Max(1, ReminderIntervalValue);
                return $"{value} {NoteReminderIntervalUnits.ToDisplay(ReminderIntervalUnit).ToLowerInvariant()}";
            }
        }

        [JsonIgnore]
        public DateTime? NextReminderAt => ComputeNextReminderAt(DateTime.Now);

        [JsonIgnore]
        public string NextReminderText => NextReminderAt.HasValue ? NextReminderAt.Value.ToString("dd.MM.yyyy HH:mm") : "—";

        [JsonIgnore]
        public bool IsReminderDue => IsReminderDueAt(DateTime.Now);

        [JsonIgnore]
        public string ReminderStatus
        {
            get
            {
                if (IsDone)
                    return "Завершено";

                var mode = NoteReminderModes.Normalize(ReminderMode);
                if (mode == NoteReminderModes.None)
                    return "Без напоминания";

                if (IsReminderDue)
                    return "Требуется";

                return "Запланировано";
            }
        }

        [JsonIgnore]
        public string ReminderStatusDisplay => LevelMarkHelper.PreventSingleLetterWrap(ReminderStatus);

        public DateTime? ComputeNextReminderAt(DateTime now)
        {
            if (IsDone)
                return null;

            var mode = NoteReminderModes.Normalize(ReminderMode);
            if (mode == NoteReminderModes.None)
                return null;

            var start = ReminderStart ?? CreatedAtLocal;
            if (mode == NoteReminderModes.Once)
                return ReminderLastAcknowledgedAt.HasValue ? null : start;

            if (mode == NoteReminderModes.Periodic)
            {
                if (!ReminderLastAcknowledgedAt.HasValue)
                    return start;

                var interval = NoteReminderIntervalUnits.ToTimeSpan(ReminderIntervalUnit, ReminderIntervalValue);
                return ReminderLastAcknowledgedAt.Value + interval;
            }

            // Постоянное напоминание не "гасится", оно активно с даты старта.
            return start;
        }

        public bool IsReminderDueAt(DateTime now)
        {
            if (IsDone)
                return false;

            var next = ComputeNextReminderAt(now);
            if (!next.HasValue)
                return false;

            return now >= next.Value;
        }

        public void AcknowledgeReminder(DateTime? timestamp = null)
        {
            var now = timestamp ?? DateTime.Now;
            if (NoteReminderModes.Normalize(ReminderMode) == NoteReminderModes.None)
                return;

            ReminderLastAcknowledgedAt = now;
            UpdatedAtUtc = DateTime.UtcNow;
            NotifyComputedProperties();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            NotifyComputedProperties();
            return true;
        }

        private void NotifyComputedProperties()
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TitleDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(BodyDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CreatedAtLocal)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(UpdatedAtLocal)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReminderModeDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReminderIntervalDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NextReminderAt)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NextReminderText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsReminderDue)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReminderStatus)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReminderStatusDisplay)));
        }
    }

    public static class LevelMarkHelper
    {
        private static readonly Regex SingleLetterWordPattern = new(
            @"(?<=^|\s)([A-Za-zА-Яа-я])\s+(?=\S)",
            RegexOptions.Compiled);

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
                .Select(x => PreventSingleLetterWrap(x.Trim()))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public static string PreventSingleLetterWrap(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var normalized = text.Replace('\u00A0', ' ');
            return SingleLetterWordPattern.Replace(normalized, match => $"{match.Groups[1].Value}\u00A0");
        }
    }




}

