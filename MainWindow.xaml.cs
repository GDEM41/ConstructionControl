using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Documents;
using System.Windows.Threading;
using WinForms = System.Windows.Forms;
using WpfPath = System.Windows.Shapes.Path;
using System.Text.RegularExpressions;

public enum ExportMode
{
    Merged,
    Detailed
}



namespace ConstructionControl
{
    public partial class MainWindow : Window
    {
        private bool arrivalPanelVisible = false;
        private readonly Dictionary<string, double> columnWidths = new();

        private readonly List<string> colorPalette = new()
{
    "#EAF2FF", "#EEF4FF", "#F5F7FA", "#F9FAFB", "#E6F0FF",
    "#F0F6FF", "#E8F1FF", "#EDF3FF", "#F3F7FF", "#EAF4FF"
};

        private readonly Dictionary<string, Brush> colorMap = new();
        private int colorIndex = 0;

        private Brush GetColor(string group)
        {
            if (!colorMap.ContainsKey(group))
            {
                var color = (Color)ColorConverter.ConvertFromString(colorPalette[colorIndex % colorPalette.Count]);
                colorMap[group] = new SolidColorBrush(color);
                colorIndex++;
            }
            return colorMap[group];
        }

        private const string DefaultSaveFileName = "data.json";
        private string currentSaveFileName = DefaultSaveFileName;
        // ===== ИСТОРИЯ ДЛЯ НАЗАД / ВПЕРЁД =====
        private readonly Stack<AppState> undoStack = new();
        private readonly Stack<AppState> redoStack = new();

        private ProjectObject currentObject;
        private List<JournalRecord> journal = new();
        private List<JournalRecord> filteredJournal = new();
      

        private bool isLocked;
        private bool mergeEnabled = false;

        private Grid summaryGrid;
        private int summaryRowIndex;
        private List<SummaryColumnInfo> summaryColumns;
        private List<SummaryBlockInfo> summaryBlocks;
        private int summaryTotalColumn;
        private int summaryNotArrivedColumn;
        private int summaryArrivedColumn;
        private bool summaryFilterUpdating;
        private List<string> summaryFilterGroups = new();
        private string summarySelectedSubType = string.Empty;
        private bool summaryMountedMode;
        private readonly ObservableCollection<string> brigadierNames = new();
        private readonly ObservableCollection<string> specialties = new();
        private readonly ObservableCollection<string> professions = new();
        private string otSearchText = string.Empty;
        private bool isTreePinned;
        private Point treeDragStart;
        private readonly ObservableCollection<TimesheetRowViewModel> timesheetRows = new();
        private readonly ObservableCollection<string> timesheetBrigades = new();
        private readonly ObservableCollection<string> timesheetAssignableBrigades = new();
        private readonly List<TimesheetPersonEntry> subscribedTimesheetPeople = new();
        private DateTime timesheetMonth = new(DateTime.Today.Year, DateTime.Today.Month, 1);
        private string selectedTimesheetBrigade = "Все бригады";
        private TimesheetRowViewModel selectedTimesheetRow;
        private int selectedTimesheetDay = -1;
        private readonly ObservableCollection<ProductionJournalEntry> productionJournalRows = new();
        private readonly ObservableCollection<string> productionActions = new();
        private readonly ObservableCollection<string> productionTargets = new();
        private readonly ObservableCollection<string> productionElements = new();
        private readonly ObservableCollection<string> productionBlockOptions = new();
        private readonly ObservableCollection<string> productionMarkOptions = new();
        private readonly ObservableCollection<string> productionWeatherOptions = new();
        private readonly ObservableCollection<string> productionDeviationOptions = new();
        private readonly ObservableCollection<InspectionJournalEntry> inspectionJournalRows = new();
        private readonly ObservableCollection<string> inspectionJournalNames = new();
        private readonly ObservableCollection<string> inspectionNames = new();
        private string productionRowSnapshotJson;
        private ProductionJournalEntry selectedProductionRow;
        private string inspectionRowSnapshotJson;
        private InspectionJournalEntry selectedInspectionRow;
        private bool initialUiPrepared;
        private bool arrivalMatrixMode;
        private bool arrivalLegacyRefreshPending;
        private DocumentTreeNode selectedPdfNode;
        private DocumentTreeNode selectedEstimateNode;
        private Point pdfTreeDragStart;
        private Point estimateTreeDragStart;
        private DocumentTreeNode pdfDragNode;
        private DocumentTreeNode estimateDragNode;
        private bool isPdfTreePinned;
        private bool isEstimateTreePinned;
        private readonly Random productionAutoRandom = new();
        private readonly ObservableCollection<ReminderSectionViewModel> reminderSections = new();
        private readonly DispatcherTimer reminderRefreshTimer;
        private readonly DispatcherTimer reminderRefreshDebounceTimer;
        private readonly DispatcherTimer timesheetRebuildDebounceTimer;
        private DateTime? reminderSnoozedUntil;
        private bool timesheetNeedsRebuild;
        private bool reminderRefreshRequested;
        private bool timesheetRebuildRequested;
        private bool timesheetRebuildForceRequested;
        private bool timesheetOtSyncDirty = true;
        private bool isSyncingTimesheetToOt;
        private const int MaxTimesheetMissingDocs = 3;
        private ReminderOverlayWindow reminderOverlayWindow;
        private WinForms.Panel estimateExcelPanel;
        private object estimateExcelApplication;
        private object estimateExcelWorkbook;
        private IntPtr estimateExcelWindowHandle = IntPtr.Zero;
        private string estimateEmbeddedFilePath = string.Empty;
        private bool previewWarmupStarted;
        private string lastSavedStateSnapshot = string.Empty;
        private bool closeConfirmed;

        private const int GWL_STYLE = -16;
        private const int GWL_EXSTYLE = -20;
        private const long WS_CHILD = 0x40000000L;
        private const long WS_CAPTION = 0x00C00000L;
        private const long WS_THICKFRAME = 0x00040000L;
        private const long WS_MINIMIZEBOX = 0x00020000L;
        private const long WS_MAXIMIZEBOX = 0x00010000L;
        private const long WS_SYSMENU = 0x00080000L;
        private const long WS_POPUP = unchecked((int)0x80000000);
        private const long WS_EX_APPWINDOW = 0x00040000L;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const int SW_SHOW = 5;
        private const int EmbeddedExcelTopTrim = 40;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(
            IntPtr hWnd,
            IntPtr hWndInsertAfter,
            int x,
            int y,
            int cx,
            int cy,
            uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        private sealed class TimesheetRowViewModel : INotifyPropertyChanged
        {
            public sealed class PresenceAccessor
            {
                private readonly TimesheetRowViewModel owner;
                public PresenceAccessor(TimesheetRowViewModel owner) => this.owner = owner;
                public string this[int day]
                {
                    get => owner.GetPresenceMark(day);
                    set
                    {
                        owner.SetPresenceMark(day, value);
                        owner.OnPropertyChanged($"Presence[{day}]");
                    }
                }
            }

            public sealed class PresenceCheckedAccessor
            {
                private readonly TimesheetRowViewModel owner;
                public PresenceCheckedAccessor(TimesheetRowViewModel owner) => this.owner = owner;
                public bool this[int day]
                {
                    get => owner.GetPresenceChecked(day);
                    set
                    {
                        owner.SetPresenceChecked(day, value);
                        owner.OnPropertyChanged($"PresenceChecked[{day}]");
                        owner.OnPropertyChanged($"Presence[{day}]");
                    }
                }
            }

            private readonly TimesheetPersonEntry source;
            private readonly string monthKey;
            private int number;
            private bool isCrewStart;
            private bool isCrewEnd;
            private double monthTotalHours;

            public TimesheetRowViewModel(TimesheetPersonEntry source, string monthKey)
            {
                this.source = source;
                this.monthKey = monthKey;
                FullName = source.FullName;
                Specialty = source.Specialty;
                Rank = source.Rank;
                BrigadeName = source.BrigadeName;
                IsBrigadier = source.IsBrigadier;
                PersonId = source.PersonId;
                Presence = new PresenceAccessor(this);
                PresenceChecked = new PresenceCheckedAccessor(this);
            }
            public Guid PersonId { get; }
            public TimesheetPersonEntry Source => source;
            public PresenceAccessor Presence { get; }
            public PresenceCheckedAccessor PresenceChecked { get; }

            public string FullName
            {
                get => source.FullName;
                set
                {
                    var trimmed = value?.Trim();
                    if (string.Equals(source.FullName, trimmed, StringComparison.CurrentCulture))
                        return;
                    source.FullName = trimmed;
                    if (source.IsBrigadier)
                        source.BrigadeName = trimmed;
                    OnPropertyChanged(nameof(FullName));
                    OnPropertyChanged(nameof(BrigadeName));
                }
            }

            public int DailyWorkHours
            {
                get => source.DailyWorkHours <= 0 ? 8 : source.DailyWorkHours;
                set
                {
                    var normalized = Math.Clamp(value, 1, 24);
                    if (source.DailyWorkHours == normalized)
                        return;

                    source.DailyWorkHours = normalized;
                    OnPropertyChanged(nameof(DailyWorkHours));
                }
            }

            public string Specialty
            {
                get => source.Specialty;
                set
                {
                    var trimmed = value?.Trim();
                    if (string.Equals(source.Specialty, trimmed, StringComparison.CurrentCulture))
                        return;
                    source.Specialty = trimmed;
                    OnPropertyChanged(nameof(Specialty));
                }
            }

            public string Rank
            {
                get => source.Rank;
                set
                {
                    var trimmed = value?.Trim();
                    if (string.Equals(source.Rank, trimmed, StringComparison.CurrentCulture))
                        return;
                    source.Rank = trimmed;
                    OnPropertyChanged(nameof(Rank));
                }
            }

            public string BrigadeName
            {
                get => source.BrigadeName;
                set
                {
                    var trimmed = value?.Trim();
                    if (source.IsBrigadier)
                        trimmed = source.FullName?.Trim();
                    if (string.Equals(source.BrigadeName, trimmed, StringComparison.CurrentCulture))
                        return;
                    source.BrigadeName = trimmed;
                    OnPropertyChanged(nameof(BrigadeName));
                }
            }

            public bool IsBrigadier
            {
                get => source.IsBrigadier;
                set
                {
                    if (source.IsBrigadier == value)
                        return;
                    source.IsBrigadier = value;
                    source.BrigadeName = value ? source.FullName?.Trim() : source.BrigadeName?.Trim();
                    OnPropertyChanged(nameof(IsBrigadier));
                    OnPropertyChanged(nameof(BrigadeName));
                }
            }

            public int Number
            {
                get => number;
                set { number = value; OnPropertyChanged(nameof(Number)); }
            }

            public bool IsCrewStart
            {
                get => isCrewStart;
                set { isCrewStart = value; OnPropertyChanged(nameof(IsCrewStart)); }
            }

            public bool IsCrewEnd
            {
                get => isCrewEnd;
                set { isCrewEnd = value; OnPropertyChanged(nameof(IsCrewEnd)); }
            }

            public double MonthTotalHours
            {
                get => monthTotalHours;
                set { monthTotalHours = value; OnPropertyChanged(nameof(MonthTotalHours)); }
            }

            public string GetDayValue(int day) => source.GetDayValue(monthKey, day);
            public string GetDayComment(int day) => source.GetDayComment(monthKey, day);
            public bool HasDayComment(int day) => source.HasDayComment(monthKey, day);
            public bool IsNonHourCode(int day) => source.IsNonHourCode(monthKey, day);
            public bool? IsDocumentAccepted(int day) => source.GetDocumentAccepted(monthKey, day);
            public void SetComment(int day, string comment) => source.SetDayComment(monthKey, day, comment);
            public void SetDocumentAccepted(int day, bool? accepted) => source.SetDocumentAccepted(monthKey, day, accepted);
            public string GetPresenceMark(int day) => source.GetPresenceMark(monthKey, day);
            public void SetPresenceMark(int day, string mark) => source.SetPresenceMark(monthKey, day, mark);
            public bool GetPresenceChecked(int day) => source.GetPresenceChecked(monthKey, day);
            public void SetPresenceChecked(int day, bool isChecked) => source.SetPresenceChecked(monthKey, day, isChecked);

            public void SetDayValue(int day, string value)
            {
                source.SetDayValue(monthKey, day, value);
                RecalculateTotal();
            }
            public string this[int day]
            {
                get => GetDayValue(day);
                set
                {
                    SetDayValue(day, value);

                    // уведомляем WPF
                    OnPropertyChanged($"Item[{day}]");
                    OnPropertyChanged(nameof(MonthTotalHours));
                }
            }

            public void RecalculateTotal()
            {
                double total = 0;
                for (var day = 1; day <= 31; day++)
                {
                    var raw = source.GetDayValue(monthKey, day);
                    if (double.TryParse(raw, NumberStyles.Any, CultureInfo.CurrentCulture, out var hours)
                        || double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out hours))
                    {
                        total += hours;
                    }
                }
                MonthTotalHours = total;
            }
            public Brush GetDayBackground(int day)
            {
                if (!IsNonHourCode(day))
                    return Brushes.Transparent;

                var docAccepted = IsDocumentAccepted(day);
                if (docAccepted == true)
                    return new SolidColorBrush(Color.FromRgb(254, 240, 138));

                return new SolidColorBrush(Color.FromRgb(254, 226, 226));
            }


            public event PropertyChangedEventHandler PropertyChanged;
            private void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        private sealed class TreeNodeMeta
        {
            public string Kind { get; set; }
            public string MaterialName { get; set; }
            public string GroupName { get; set; }
            public string SubCategory { get; set; }
            public string Category { get; set; }
            public List<string> PrefixSegments { get; set; }
        }

        private sealed class SelectableOption : INotifyPropertyChanged
        {
            private bool isSelected;

            public string Value { get; set; } = string.Empty;

            public bool IsSelected
            {
                get => isSelected;
                set
                {
                    if (isSelected == value)
                        return;

                    isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private sealed class ProductionItemEditorRow : INotifyPropertyChanged
        {
            private string materialName = string.Empty;
            private double quantity;

            public string MaterialName
            {
                get => materialName;
                set
                {
                    if (materialName == value)
                        return;

                    materialName = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MaterialName)));
                }
            }

            public double Quantity
            {
                get => quantity;
                set
                {
                    if (Math.Abs(quantity - value) < 0.0001)
                        return;

                    quantity = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Quantity)));
                }
            }

            public ObservableCollection<string> AvailableNames { get; set; } = new();

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private sealed class SummaryReorderPreviewRow
        {
            public string Group { get; set; }
            public string Material { get; set; }
            public int Block { get; set; }
            public string Mark { get; set; }
            public double Quantity { get; set; }
            public string Unit { get; set; }
        }

        private sealed class SummaryBalanceReminderItem
        {
            public string Category { get; set; }
            public string Group { get; set; }
            public string Material { get; set; }
            public string Unit { get; set; }
            public double Quantity { get; set; }
            public bool IsOverage { get; set; }
        }

        private sealed class AutoProductionCandidate
        {
            public string Group { get; set; }
            public string Material { get; set; }
            public string Unit { get; set; }
            public double Available { get; set; }
            public double Deficit { get; set; }
            public double RemainingToPlan { get; set; }
        }

        public ObservableCollection<string> BrigadierNames => brigadierNames;
        public ObservableCollection<string> Specialties => specialties;
        public ObservableCollection<string> Professions => professions;
        public ObservableCollection<string> TimesheetAssignableBrigades => timesheetAssignableBrigades;
        public ObservableCollection<string> ProductionActions => productionActions;
        public ObservableCollection<string> ProductionTargets => productionTargets;
        public ObservableCollection<string> ProductionElements => productionElements;
        public ObservableCollection<string> ProductionBlockOptions => productionBlockOptions;
        public ObservableCollection<string> ProductionMarkOptions => productionMarkOptions;
        public ObservableCollection<string> ProductionWeatherOptions => productionWeatherOptions;
        public ObservableCollection<string> ProductionDeviationOptions => productionDeviationOptions;
        public ObservableCollection<string> InspectionJournalNames => inspectionJournalNames;
        public ObservableCollection<string> InspectionNames => inspectionNames;
        public MainWindow()
        {
            InitializeComponent();
            currentSaveFileName = ResolveDefaultSavePath();
            EnsureReminderOverlayWindow();
            InitializeEstimatePreviewHost();
            SizeChanged += (_, _) => UpdateReminderOverlayPlacement();
            LocationChanged += (_, _) => UpdateReminderOverlayPlacement();
            StateChanged += MainWindow_StateChanged;
            reminderRefreshTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(30)
            };
            reminderRefreshTimer.Tick += ReminderRefreshTimer_Tick;
            reminderRefreshTimer.Start();
            reminderRefreshDebounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(150)
            };
            reminderRefreshDebounceTimer.Tick += ReminderRefreshDebounceTimer_Tick;

            timesheetRebuildDebounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(120)
            };
            timesheetRebuildDebounceTimer.Tick += TimesheetRebuildDebounceTimer_Tick;
            // ===== БЛОКИРОВКА ВКЛЮЧЕНА ПО УМОЛЧАНИЮ =====
            isLocked = true;

            LoadState();
            InitializeOtJournal();
            InitializeTimesheet();
            InitializeProductionJournal();
            InitializeInspectionJournal();
            filteredJournal = journal.ToList();

            ArrivalPanel.ArrivalAdded += OnArrivalAdded;

            PushUndo();
            UpdateUndoRedoButtons();

            if (currentObject != null)
                ArrivalPanel.SetObject(currentObject, journal);
            RefreshArrivalTypes();
            RefreshArrivalNames();
            RefreshDocumentLibraries();
            UpdateArrivalViewMode();

            RefreshTreePreserveState();
            ApplyProjectUiSettings();
            StartPreviewWarmupAsync();
            lastSavedStateSnapshot = BuildCurrentStateJson();

        }

        private static string ResolveDefaultSavePath()
        {
            static void AddCandidate(List<string> list, string? path)
            {
                if (string.IsNullOrWhiteSpace(path))
                    return;

                var full = System.IO.Path.GetFullPath(path);
                if (!list.Contains(full, StringComparer.OrdinalIgnoreCase))
                    list.Add(full);
            }

            var candidates = new List<string>();
            AddCandidate(candidates, System.IO.Path.Combine(Environment.CurrentDirectory, DefaultSaveFileName));
            AddCandidate(candidates, System.IO.Path.Combine(AppContext.BaseDirectory, DefaultSaveFileName));

            var scanDir = new DirectoryInfo(AppContext.BaseDirectory);
            for (var i = 0; i < 6 && scanDir != null; i++, scanDir = scanDir.Parent)
            {
                AddCandidate(candidates, System.IO.Path.Combine(scanDir.FullName, DefaultSaveFileName));
            }

            var existing = candidates.FirstOrDefault(File.Exists);
            return existing ?? candidates[0];
        }

        private void StartPreviewWarmupAsync()
        {
            if (previewWarmupStarted)
                return;

            previewWarmupStarted = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                WarmupExcelInterop();
            }), DispatcherPriority.ApplicationIdle);
        }

        private void WarmupExcelInterop()
        {
            try
            {
                _ = EnsureEstimateExcelApplication();
            }
            catch
            {
                // Ignore Excel warmup errors to keep app startup resilient.
            }
        }

        private dynamic EnsureEstimateExcelApplication()
        {
            if (estimateExcelApplication != null)
                return estimateExcelApplication;

            var excelType = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Microsoft Excel не найден в системе.");

            dynamic excelApp = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("Не удалось создать экземпляр Excel.");

            try
            {
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                try { excelApp.AskToUpdateLinks = false; } catch { }
                try { excelApp.EnableEvents = false; } catch { }
                try { excelApp.ScreenUpdating = true; } catch { }
                try { excelApp.DisplayStatusBar = true; } catch { }
                try { excelApp.DisplayFormulaBar = true; } catch { }
            }
            catch
            {
                // Ignore non-critical UI tuning errors for compatibility with different Excel versions.
            }

            estimateExcelApplication = excelApp;
            return excelApp;
        }

        private void ConfigureEstimateExcelLiteUi(dynamic excelApp)
        {
            if (excelApp == null)
                return;

            try
            {
                excelApp.Visible = true;
                excelApp.DisplayAlerts = false;
                try { excelApp.WindowState = -4143; } catch { } // xlNormal
                try { excelApp.DisplayFullScreen = false; } catch { }
                try { excelApp.DisplayStatusBar = true; } catch { }
                try { excelApp.DisplayFormulaBar = true; } catch { }
                try { excelApp.Interactive = true; } catch { }
                try { excelApp.UserControl = true; } catch { }
                try { excelApp.CommandBars["Ribbon"].Visible = true; } catch { }
                try { excelApp.ExecuteExcel4Macro("SHOW.TOOLBAR(\"Ribbon\",True)"); } catch { }
                try { excelApp.ExecuteExcel4Macro("SHOW.TOOLBAR(\"Standard\",True)"); } catch { }
                try { excelApp.ExecuteExcel4Macro("SHOW.TOOLBAR(\"Formatting\",True)"); } catch { }
                try { excelApp.CommandBars["Worksheet Menu Bar"].Enabled = true; } catch { }
                try
                {
                    try { excelApp.CommandBars.ExecuteMso("MinimizeRibbon"); } catch { }
                    var ribbonMinimized = excelApp.CommandBars.GetPressedMso("MinimizeRibbon");
                    var isMinimized = false;
                    if (ribbonMinimized is bool boolValue)
                    {
                        isMinimized = boolValue;
                    }
                    else
                    {
                        var minimizedText = ribbonMinimized?.ToString();
                        var parsedBool = false;
                        if (!string.IsNullOrWhiteSpace(minimizedText)
                            && bool.TryParse(minimizedText, out parsedBool))
                        {
                            isMinimized = parsedBool;
                        }
                    }
                    if (isMinimized)
                        excelApp.CommandBars.ExecuteMso("MinimizeRibbon");
                }
                catch { }
                try { excelApp.CommandBars.ExecuteMso("TabHome"); } catch { }
                try { excelApp.ActiveWindow.DisplayHeadings = true; } catch { }
                try { excelApp.ActiveWindow.DisplayGridlines = true; } catch { }
                try { excelApp.ActiveWindow.DisplayWorkbookTabs = true; } catch { }
                try { excelApp.ActiveWindow.DisplayHorizontalScrollBar = true; } catch { }
                try { excelApp.ActiveWindow.DisplayVerticalScrollBar = true; } catch { }
                try { excelApp.ActiveWindow.View = 1; } catch { } // xlNormalView
                try { excelApp.ActiveWindow.Zoom = 100; } catch { }
            }
            catch
            {
                // Keep preview alive even if a UI tweak is unsupported on this Excel build.
            }

        }

        private void CloseEstimateWorkbook(bool saveChanges)
        {
            if (estimateExcelWorkbook == null)
                return;

            try
            {
                ((dynamic)estimateExcelWorkbook).Close(saveChanges);
            }
            catch
            {
                // Ignore workbook close errors during preview cleanup.
            }
            finally
            {
                Marshal.FinalReleaseComObject(estimateExcelWorkbook);
                estimateExcelWorkbook = null;
                estimateExcelWindowHandle = IntPtr.Zero;
                estimateEmbeddedFilePath = string.Empty;
            }
        }

        private void DisposeEstimateExcelApplication()
        {
            CloseEstimateWorkbook(saveChanges: true);

            if (estimateExcelApplication != null)
            {
                try
                {
                    ((dynamic)estimateExcelApplication).Quit();
                }
                catch
                {
                    // Ignore Excel shutdown errors on app close.
                }
                finally
                {
                    Marshal.FinalReleaseComObject(estimateExcelApplication);
                    estimateExcelApplication = null;
                }
            }
        }

        private ReminderOverlayWindow EnsureReminderOverlayWindow()
        {
            if (reminderOverlayWindow != null)
                return reminderOverlayWindow;

            reminderOverlayWindow = new ReminderOverlayWindow();
            reminderOverlayWindow.SectionsHostElement.ItemsSource = reminderSections;
            reminderOverlayWindow.SnoozeRequested += SnoozeRemindersButton_Click;
            reminderOverlayWindow.ToggleDetailsRequested += ToggleReminderDetailsButton_Click;
            return reminderOverlayWindow;
        }

        private void InitializeEstimatePreviewHost()
        {
            if (EstimateExcelHost == null || estimateExcelPanel != null)
                return;

            estimateExcelPanel = new WinForms.Panel
            {
                BackColor = System.Drawing.Color.White
            };
            estimateExcelPanel.Resize += (_, _) => LayoutEmbeddedEstimateWindow();
            EstimateExcelHost.Child = estimateExcelPanel;
        }

        private void MainWindow_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                reminderOverlayWindow?.Hide();
                return;
            }

            UpdateReminderOverlayPlacement();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            reminderRefreshTimer?.Stop();
            reminderRefreshDebounceTimer?.Stop();
            timesheetRebuildDebounceTimer?.Stop();
            HideReminderOverlayWindow();
            StopEstimateEmbeddedPreview();
            DisposeEstimateExcelApplication();
            if (reminderOverlayWindow != null)
            {
                reminderOverlayWindow.Close();
                reminderOverlayWindow = null;
            }
        }

        private void ReminderRefreshTimer_Tick(object sender, EventArgs e)
        {
            RequestReminderRefresh();
        }

        private void ReminderRefreshDebounceTimer_Tick(object sender, EventArgs e)
        {
            reminderRefreshDebounceTimer.Stop();
            if (!reminderRefreshRequested)
                return;

            reminderRefreshRequested = false;
            UpdateOtReminders();
        }

        private void TimesheetRebuildDebounceTimer_Tick(object sender, EventArgs e)
        {
            timesheetRebuildDebounceTimer.Stop();
            if (!timesheetRebuildRequested)
                return;

            var force = timesheetRebuildForceRequested;
            timesheetRebuildRequested = false;
            timesheetRebuildForceRequested = false;
            RebuildTimesheetView(force);
        }

        private void RequestReminderRefresh(bool immediate = false)
        {
            if (immediate || reminderRefreshDebounceTimer == null)
            {
                reminderRefreshRequested = false;
                reminderRefreshDebounceTimer?.Stop();
                UpdateOtReminders();
                return;
            }

            reminderRefreshRequested = true;
            reminderRefreshDebounceTimer.Stop();
            reminderRefreshDebounceTimer.Start();
        }

        private void RequestTimesheetRebuild(bool force = false)
        {
            if (force)
                timesheetRebuildForceRequested = true;

            if (!force && !ReferenceEquals(MainTabs?.SelectedItem, TimesheetTab))
            {
                timesheetNeedsRebuild = true;
                timesheetRebuildRequested = false;
                timesheetRebuildDebounceTimer?.Stop();
                return;
            }

            if (timesheetRebuildDebounceTimer == null)
            {
                RebuildTimesheetView(force: force || timesheetRebuildForceRequested);
                timesheetRebuildForceRequested = false;
                return;
            }

            timesheetRebuildRequested = true;
            timesheetRebuildDebounceTimer.Stop();
            timesheetRebuildDebounceTimer.Start();
        }

        private void MarkTimesheetOtSyncDirty()
        {
            timesheetOtSyncDirty = true;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (initialUiPrepared)
                return;

            initialUiPrepared = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ApplyAllFilters();
                Activate();
                RequestReminderRefresh(immediate: true);
                UpdateReminderOverlayPlacement();
            }), DispatcherPriority.Background);
        }

        private void EnsureDocumentLibraries()
        {
            if (currentObject == null)
                return;

            currentObject.PdfDocuments ??= new List<DocumentTreeNode>();
            currentObject.EstimateDocuments ??= new List<DocumentTreeNode>();
            NormalizeDocumentPaths(currentObject.PdfDocuments, isPdfLibrary: true);
            NormalizeDocumentPaths(currentObject.EstimateDocuments, isPdfLibrary: false);
        }

        private void NormalizeDocumentPaths(IEnumerable<DocumentTreeNode> nodes, bool isPdfLibrary)
        {
            if (nodes == null)
                return;

            foreach (var node in nodes)
            {
                if (node == null)
                    continue;

                if (node.IsFolder)
                {
                    NormalizeDocumentPaths(node.Children, isPdfLibrary);
                    continue;
                }

                if (string.IsNullOrWhiteSpace(node.StoredRelativePath) && !string.IsNullOrWhiteSpace(node.FilePath))
                {
                    var relative = TryBuildStoredRelativePath(node.FilePath, isPdfLibrary);
                    if (!string.IsNullOrWhiteSpace(relative))
                        node.StoredRelativePath = relative;
                }

                var resolved = ResolveDocumentPath(node);
                if (!string.IsNullOrWhiteSpace(resolved))
                    node.FilePath = resolved;

                NormalizeDocumentPaths(node.Children, isPdfLibrary);
            }
        }

        private string TryBuildStoredRelativePath(string filePath, bool isPdfLibrary)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return string.Empty;

            try
            {
                var folder = GetDocumentStorageFolder(isPdfLibrary, createIfMissing: true);
                if (string.IsNullOrWhiteSpace(folder))
                    return string.Empty;

                var fullFilePath = System.IO.Path.GetFullPath(filePath);
                var fullFolderPath = System.IO.Path.GetFullPath(folder);
                if (!fullFilePath.StartsWith(fullFolderPath, StringComparison.OrdinalIgnoreCase))
                    return string.Empty;

                var fileName = System.IO.Path.GetFileName(fullFilePath);
                return System.IO.Path.Combine(isPdfLibrary ? "pdf" : "estimate", fileName);
            }
            catch
            {
                return string.Empty;
            }
        }

        private string ResolveDocumentPath(DocumentTreeNode node)
        {
            if (node == null || node.IsFolder)
                return string.Empty;

            if (!string.IsNullOrWhiteSpace(node.FilePath) && File.Exists(node.FilePath))
                return node.FilePath;

            if (!string.IsNullOrWhiteSpace(node.StoredRelativePath))
            {
                var baseFolder = GetProjectStorageRoot(createIfMissing: false);
                if (!string.IsNullOrWhiteSpace(baseFolder))
                {
                    var candidate = System.IO.Path.Combine(baseFolder, node.StoredRelativePath);
                    if (File.Exists(candidate))
                        return candidate;
                }
            }

            return node.FilePath ?? string.Empty;
        }

        private string GetProjectStorageRoot(bool createIfMissing)
        {
            if (string.IsNullOrWhiteSpace(currentSaveFileName))
                return string.Empty;

            var savePath = System.IO.Path.GetFullPath(currentSaveFileName);
            var saveFolder = System.IO.Path.GetDirectoryName(savePath);
            var saveName = System.IO.Path.GetFileNameWithoutExtension(savePath);
            if (string.IsNullOrWhiteSpace(saveFolder) || string.IsNullOrWhiteSpace(saveName))
                return string.Empty;

            var storageRoot = System.IO.Path.Combine(saveFolder, $"{saveName}_files");
            if (createIfMissing)
                Directory.CreateDirectory(storageRoot);

            return storageRoot;
        }

        private string GetDocumentStorageFolder(bool isPdfLibrary, bool createIfMissing)
        {
            var root = GetProjectStorageRoot(createIfMissing);
            if (string.IsNullOrWhiteSpace(root))
                return string.Empty;

            var folder = System.IO.Path.Combine(root, isPdfLibrary ? "pdf" : "estimate");
            if (createIfMissing)
                Directory.CreateDirectory(folder);

            return folder;
        }

        private bool TryCopyDocumentToStorage(string sourceFilePath, bool isPdfLibrary, out string storedAbsolutePath, out string storedRelativePath)
        {
            storedAbsolutePath = sourceFilePath;
            storedRelativePath = string.Empty;

            if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
                return false;

            var folder = GetDocumentStorageFolder(isPdfLibrary, createIfMissing: true);
            if (string.IsNullOrWhiteSpace(folder))
                return false;

            try
            {
                var extension = System.IO.Path.GetExtension(sourceFilePath);
                var baseName = System.IO.Path.GetFileNameWithoutExtension(sourceFilePath);
                foreach (var invalid in System.IO.Path.GetInvalidFileNameChars())
                    baseName = baseName.Replace(invalid, '_');
                if (string.IsNullOrWhiteSpace(baseName))
                    baseName = "document";

                var uniqueName = $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmssfff}_{Guid.NewGuid().ToString("N")[..6]}{extension}";
                var targetPath = System.IO.Path.Combine(folder, uniqueName);
                File.Copy(sourceFilePath, targetPath, overwrite: false);

                storedAbsolutePath = targetPath;
                storedRelativePath = System.IO.Path.Combine(isPdfLibrary ? "pdf" : "estimate", uniqueName);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void RefreshDocumentLibraries()
        {
            EnsureDocumentLibraries();

            if (PdfTreeView != null)
            {
                PdfTreeView.ItemsSource = null;
                PdfTreeView.ItemsSource = currentObject?.PdfDocuments;
            }

            if (EstimateTreeView != null)
            {
                EstimateTreeView.ItemsSource = null;
                EstimateTreeView.ItemsSource = currentObject?.EstimateDocuments;
            }

            if (!ContainsDocumentNode(currentObject?.PdfDocuments, selectedPdfNode))
                selectedPdfNode = null;

            if (!ContainsDocumentNode(currentObject?.EstimateDocuments, selectedEstimateNode))
                selectedEstimateNode = null;

            UpdatePdfSelectionInfo();
            UpdateEstimateSelectionInfo();
            UpdateTabButtons();
            UpdatePdfTreePanelState(forceVisible: isPdfTreePinned);
            UpdateEstimateTreePanelState(forceVisible: isEstimateTreePinned);
        }

        private static bool ContainsDocumentNode(IEnumerable<DocumentTreeNode> nodes, DocumentTreeNode target)
        {
            if (nodes == null || target == null)
                return false;

            foreach (var node in nodes)
            {
                if (ReferenceEquals(node, target))
                    return true;

                if (ContainsDocumentNode(node.Children, target))
                    return true;
            }

            return false;
        }

        private static DocumentTreeNode FindDocumentParent(IEnumerable<DocumentTreeNode> nodes, DocumentTreeNode target)
        {
            if (nodes == null || target == null)
                return null;

            foreach (var node in nodes)
            {
                if (node.Children?.Contains(target) == true)
                    return node;

                var nested = FindDocumentParent(node.Children, target);
                if (nested != null)
                    return nested;
            }

            return null;
        }

        private static bool IsDocumentDescendant(DocumentTreeNode parent, DocumentTreeNode candidate)
        {
            if (parent?.Children == null || candidate == null)
                return false;

            foreach (var child in parent.Children)
            {
                if (ReferenceEquals(child, candidate))
                    return true;

                if (IsDocumentDescendant(child, candidate))
                    return true;
            }

            return false;
        }

        private static List<DocumentTreeNode> GetOwningDocumentCollection(List<DocumentTreeNode> root, DocumentTreeNode node)
        {
            if (root == null || node == null)
                return null;

            if (root.Contains(node))
                return root;

            var parent = FindDocumentParent(root, node);
            return parent?.Children;
        }

        private static List<DocumentTreeNode> GetDocumentInsertCollection(List<DocumentTreeNode> root, DocumentTreeNode anchor)
        {
            if (root == null)
                return null;

            if (anchor == null)
                return root;

            if (anchor.IsFolder)
            {
                anchor.Children ??= new List<DocumentTreeNode>();
                return anchor.Children;
            }

            return GetOwningDocumentCollection(root, anchor) ?? root;
        }

        private static DocumentTreeNode GetDocumentNodeFromOriginalSource(object originalSource)
        {
            var current = originalSource as DependencyObject;
            while (current != null)
            {
                if (current is FrameworkElement fe && fe.DataContext is DocumentTreeNode node)
                    return node;

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        private void UpdateTabButtons()
        {
            SetTabButtonState(SummaryTabButton, ReferenceEquals(MainTabs?.SelectedItem, SummaryTab));
            SetTabButtonState(JvkTabButton, ReferenceEquals(MainTabs?.SelectedItem, JvkTab));
            SetTabButtonState(ArrivalTabButton, ReferenceEquals(MainTabs?.SelectedItem, ArrivalTab));
            SetTabButtonState(OtTabButton, ReferenceEquals(MainTabs?.SelectedItem, OtTab));
            SetTabButtonState(TimesheetTabButton, ReferenceEquals(MainTabs?.SelectedItem, TimesheetTab));
            SetTabButtonState(ProductionTabButton, ReferenceEquals(MainTabs?.SelectedItem, ProductionTab));
            SetTabButtonState(InspectionTabButton, ReferenceEquals(MainTabs?.SelectedItem, InspectionTab));
            SetTabButtonState(PdfPinnedTabButton, ReferenceEquals(MainTabs?.SelectedItem, PdfTab));
            SetTabButtonState(EstimatePinnedTabButton, ReferenceEquals(MainTabs?.SelectedItem, EstimateTab));
        }

        private static void SetTabButtonState(Button button, bool isActive)
        {
            if (button == null)
                return;

            button.Background = (Brush)new BrushConverter().ConvertFromString(isActive ? "#DBEAFE" : "#F9FAFB");
            button.BorderBrush = (Brush)new BrushConverter().ConvertFromString(isActive ? "#3B82F6" : "#E5E7EB");
            button.Foreground = (Brush)new BrushConverter().ConvertFromString(isActive ? "#111827" : "#6B7280");
        }

        private void SelectMainTab(TabItem tab)
        {
            if (MainTabs == null || tab == null)
                return;

            MainTabs.SelectedItem = tab;
            UpdateTabButtons();
        }

        private void SummaryTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(SummaryTab);
        private void JvkTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(JvkTab);
        private void ArrivalTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(ArrivalTab);
        private void OtTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(OtTab);
        private void TimesheetTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(TimesheetTab);
        private void ProductionTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(ProductionTab);
        private void InspectionTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(InspectionTab);
        private void PdfPinnedTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(PdfTab);
        private void EstimatePinnedTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(EstimateTab);

        private void UpdatePdfSelectionInfo()
        {
            UpdateDocumentSelectionInfo(selectedPdfNode, PdfSelectedNameText, PdfSelectedPathText, PdfSelectedTypeText);
            UpdateDocumentPreview(selectedPdfNode, PdfInfoPanel, PdfPreviewContainer, PdfPreviewBrowser, PdfPreviewPlaceholder, PdfPreviewStatusText);
        }

        private void UpdateEstimateSelectionInfo()
        {
            UpdateDocumentSelectionInfo(selectedEstimateNode, EstimateSelectedNameText, EstimateSelectedPathText, EstimateSelectedTypeText);
            UpdateEstimatePreview(selectedEstimateNode);
        }

        private void UpdateDocumentSelectionInfo(DocumentTreeNode node, TextBlock nameText, TextBlock pathText, TextBlock typeText)
        {
            if (nameText == null || pathText == null || typeText == null)
                return;

            if (node == null)
            {
                nameText.Text = "Ничего не выбрано";
                pathText.Text = "—";
                typeText.Text = "—";
                return;
            }

            nameText.Text = string.IsNullOrWhiteSpace(node.Name) ? "Без названия" : node.Name;
            var resolvedPath = ResolveDocumentPath(node);
            pathText.Text = string.IsNullOrWhiteSpace(resolvedPath) ? "—" : resolvedPath;
            typeText.Text = node.IsFolder ? "Папка" : "Файл";
        }

        private void UpdateEstimatePreview(DocumentTreeNode node)
        {
            if (EstimateInfoPanel == null || EstimatePreviewContainer == null || EstimatePreviewBrowser == null || EstimatePreviewPlaceholder == null || EstimatePreviewStatusText == null)
                return;

            InitializeEstimatePreviewHost();

            if (node == null)
            {
                StopEstimateEmbeddedPreview();
                EstimateInfoPanel.Visibility = Visibility.Visible;
                EstimatePreviewContainer.Visibility = Visibility.Collapsed;
                ShowEstimatePreviewPlaceholder("Выберите файл сметы в дереве слева, и он откроется здесь.");
                return;
            }

            if (node.IsFolder)
            {
                StopEstimateEmbeddedPreview();
                EstimateInfoPanel.Visibility = Visibility.Visible;
                EstimatePreviewContainer.Visibility = Visibility.Collapsed;
                ShowEstimatePreviewPlaceholder("Для папки предпросмотр не показывается. Выберите конкретный файл.");
                return;
            }

            var resolvedPath = ResolveDocumentPath(node);
            if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath))
            {
                StopEstimateEmbeddedPreview();
                EstimateInfoPanel.Visibility = Visibility.Visible;
                EstimatePreviewContainer.Visibility = Visibility.Collapsed;
                ShowEstimatePreviewPlaceholder("Файл не найден по сохраненному пути.");
                return;
            }

            var extension = System.IO.Path.GetExtension(resolvedPath)?.ToLowerInvariant() ?? string.Empty;
            if (!IsEstimateExcelExtension(extension))
            {
                StopEstimateEmbeddedPreview();
                UpdateDocumentPreview(node, EstimateInfoPanel, EstimatePreviewContainer, EstimatePreviewBrowser, EstimatePreviewPlaceholder, EstimatePreviewStatusText);
                return;
            }

            EstimateInfoPanel.Visibility = Visibility.Collapsed;
            EstimatePreviewContainer.Visibility = Visibility.Visible;

            try
            {
                ShowEmbeddedEstimateWorkbook(resolvedPath);
            }
            catch (Exception ex)
            {
                StopEstimateEmbeddedPreview();
                EstimateInfoPanel.Visibility = Visibility.Visible;
                EstimatePreviewContainer.Visibility = Visibility.Collapsed;
                ShowEstimatePreviewPlaceholder($"Не удалось открыть Excel-предпросмотр: {ex.Message}");
            }
        }

        private static bool IsEstimateExcelExtension(string extension)
            => extension is ".xlsx" or ".xlsm" or ".xls";

        private void ShowEstimatePreviewPlaceholder(string message)
        {
            if (EstimateExcelHost != null)
                EstimateExcelHost.Visibility = Visibility.Collapsed;

            ShowDocumentPreviewPlaceholder(EstimatePreviewBrowser, EstimatePreviewPlaceholder, EstimatePreviewStatusText, message);
        }

        private void UpdateDocumentPreview(DocumentTreeNode node, FrameworkElement infoPanel, Border previewContainer, WebBrowser browser, Border placeholder, TextBlock statusText)
        {
            if (infoPanel == null || previewContainer == null || browser == null || placeholder == null || statusText == null)
                return;

            if (node == null)
            {
                infoPanel.Visibility = Visibility.Visible;
                previewContainer.Visibility = Visibility.Collapsed;
                ShowDocumentPreviewPlaceholder(browser, placeholder, statusText, "Выберите файл в дереве слева, и он откроется здесь.");
                return;
            }

            if (node.IsFolder)
            {
                infoPanel.Visibility = Visibility.Visible;
                previewContainer.Visibility = Visibility.Collapsed;
                ShowDocumentPreviewPlaceholder(browser, placeholder, statusText, "Для папки предпросмотр не показывается. Выберите конкретный файл.");
                return;
            }

            var resolvedPath = ResolveDocumentPath(node);
            if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath))
            {
                infoPanel.Visibility = Visibility.Visible;
                previewContainer.Visibility = Visibility.Collapsed;
                ShowDocumentPreviewPlaceholder(browser, placeholder, statusText, "Файл не найден по сохраненному пути.");
                return;
            }

            infoPanel.Visibility = Visibility.Collapsed;
            previewContainer.Visibility = Visibility.Visible;

            var extension = System.IO.Path.GetExtension(resolvedPath)?.ToLowerInvariant() ?? string.Empty;

            try
            {
                switch (extension)
                {
                    case ".pdf":
                    case ".htm":
                    case ".html":
                    case ".png":
                    case ".jpg":
                    case ".jpeg":
                    case ".bmp":
                    case ".gif":
                    case ".tif":
                    case ".tiff":
                    case ".doc":
                        ShowDocumentPreviewUri(browser, placeholder, resolvedPath);
                        break;
                    case ".xlsx":
                    case ".xlsm":
                        ShowDocumentPreviewHtml(browser, placeholder, BuildWorkbookPreviewHtml(resolvedPath));
                        break;
                    case ".docx":
                        ShowDocumentPreviewHtml(browser, placeholder, BuildDocxPreviewHtml(resolvedPath));
                        break;
                    case ".xls":
                        ShowDocumentPreviewPlaceholder(
                            browser,
                            placeholder,
                            statusText,
                            "Для формата .xls встроенный предпросмотр ограничен. Используйте кнопку \"Открыть\".");
                        break;
                    case ".txt":
                    case ".log":
                    case ".json":
                    case ".xml":
                    case ".csv":
                        ShowDocumentPreviewHtml(browser, placeholder, BuildTextPreviewHtml(resolvedPath));
                        break;
                    default:
                        ShowDocumentPreviewPlaceholder(
                            browser,
                            placeholder,
                            statusText,
                            $"Для формата {extension} встроенный предпросмотр пока не сделан. Используйте кнопку \"Открыть\".");
                        break;
                }
            }
            catch (Exception ex)
            {
                ShowDocumentPreviewPlaceholder(browser, placeholder, statusText, $"Не удалось открыть предпросмотр: {ex.Message}");
            }
        }

        private static void ShowDocumentPreviewPlaceholder(WebBrowser browser, Border placeholder, TextBlock statusText, string message)
        {
            statusText.Text = message;
            placeholder.Visibility = Visibility.Visible;
            browser.Visibility = Visibility.Collapsed;

            try
            {
            browser.NavigateToString("<html><body style='background:#F9FAFB;'></body></html>");
            }
            catch
            {
                // Ignore browser cleanup errors for unsupported embedded engines.
            }
        }

        private static void ShowDocumentPreviewHtml(WebBrowser browser, Border placeholder, string html)
        {
            placeholder.Visibility = Visibility.Collapsed;
            browser.Visibility = Visibility.Visible;
            browser.NavigateToString(html);
        }

        private static void ShowDocumentPreviewUri(WebBrowser browser, Border placeholder, string filePath)
        {
            placeholder.Visibility = Visibility.Collapsed;
            browser.Visibility = Visibility.Visible;
            browser.Navigate(new Uri(filePath, UriKind.Absolute));
        }

        private void ShowEmbeddedEstimateWorkbook(string filePath)
        {
            InitializeEstimatePreviewHost();
            if (EstimateExcelHost == null || estimateExcelPanel == null)
                throw new InvalidOperationException("Область предпросмотра Excel не готова.");

            if (string.Equals(estimateEmbeddedFilePath, filePath, StringComparison.CurrentCultureIgnoreCase)
                && estimateExcelWindowHandle != IntPtr.Zero)
            {
                EstimateExcelHost.Visibility = Visibility.Visible;
                EstimatePreviewBrowser.Visibility = Visibility.Collapsed;
                EstimatePreviewPlaceholder.Visibility = Visibility.Collapsed;
                LayoutEmbeddedEstimateWindow();
                return;
            }

            CloseEstimateWorkbook(saveChanges: true);

            dynamic excelApp = EnsureEstimateExcelApplication();
            object workbooks = null;
            dynamic workbook = null;

            try
            {
                excelApp.Visible = true;
                excelApp.DisplayAlerts = false;
                workbooks = excelApp.Workbooks;
                workbook = workbooks.GetType().InvokeMember(
                    "Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    workbooks,
                    new object[] { filePath, Type.Missing, false });

                estimateExcelWorkbook = workbook;
                estimateEmbeddedFilePath = filePath;
                estimateExcelWindowHandle = new IntPtr((int)excelApp.Hwnd);
                if (estimateExcelWindowHandle == IntPtr.Zero)
                    throw new InvalidOperationException("Excel не предоставил окно для встраивания.");

                ConfigureEstimateExcelLiteUi(excelApp);

                ConfigureEmbeddedWindow(estimateExcelWindowHandle, estimateExcelPanel.Handle);
                EstimateExcelHost.Visibility = Visibility.Visible;
                EstimatePreviewBrowser.Visibility = Visibility.Collapsed;
                EstimatePreviewPlaceholder.Visibility = Visibility.Collapsed;
                LayoutEmbeddedEstimateWindow();
            }
            catch
            {
                CloseEstimateWorkbook(saveChanges: false);
                throw;
            }
            finally
            {
                if (workbooks != null)
                    Marshal.FinalReleaseComObject(workbooks);
            }
        }

        private void StopEstimateEmbeddedPreview()
        {
            if (EstimateExcelHost != null)
                EstimateExcelHost.Visibility = Visibility.Collapsed;

            CloseEstimateWorkbook(saveChanges: true);
        }

        private void LayoutEmbeddedEstimateWindow()
        {
            if (estimateExcelWindowHandle == IntPtr.Zero || estimateExcelPanel == null || estimateExcelPanel.IsDisposed)
                return;

            var width = Math.Max(0, estimateExcelPanel.ClientSize.Width);
            var hostHeight = Math.Max(0, estimateExcelPanel.ClientSize.Height);
            var topTrim = hostHeight > 120 ? EmbeddedExcelTopTrim : 0;
            var height = Math.Max(0, hostHeight + topTrim);
            SetWindowPos(
                estimateExcelWindowHandle,
                IntPtr.Zero,
                0,
                -topTrim,
                width,
                height,
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED | SWP_SHOWWINDOW);
            ShowWindow(estimateExcelWindowHandle, SW_SHOW);
        }

        private static void ConfigureEmbeddedWindow(IntPtr windowHandle, IntPtr parentHandle)
        {
            if (windowHandle == IntPtr.Zero || parentHandle == IntPtr.Zero)
                return;

            SetParent(windowHandle, parentHandle);

            var style = GetWindowLongPtr(windowHandle, GWL_STYLE).ToInt64();
            style &= ~(WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU | WS_POPUP);
            style |= WS_CHILD;
            SetWindowLongPtr(windowHandle, GWL_STYLE, new IntPtr(style));

            var exStyle = GetWindowLongPtr(windowHandle, GWL_EXSTYLE).ToInt64();
            exStyle &= ~WS_EX_APPWINDOW;
            SetWindowLongPtr(windowHandle, GWL_EXSTYLE, new IntPtr(exStyle));
        }

        private static IntPtr GetWindowLongPtr(IntPtr handle, int index)
            => IntPtr.Size == 8
                ? GetWindowLongPtr64(handle, index)
                : new IntPtr(GetWindowLong32(handle, index));

        private static IntPtr SetWindowLongPtr(IntPtr handle, int index, IntPtr value)
            => IntPtr.Size == 8
                ? SetWindowLongPtr64(handle, index, value)
                : new IntPtr(SetWindowLong32(handle, index, value.ToInt32()));

        private static string BuildTextPreviewHtml(string filePath)
        {
            const int maxChars = 16000;
            var text = File.ReadAllText(filePath);
            if (text.Length > maxChars)
                text = text[..maxChars] + Environment.NewLine + Environment.NewLine + "... предпросмотр обрезан ...";

            var encoded = System.Net.WebUtility.HtmlEncode(text);
            return WrapPreviewHtml(
                System.IO.Path.GetFileName(filePath),
                $"<pre style=\"white-space:pre-wrap;font-family:'Consolas','Segoe UI',monospace;font-size:13px;line-height:1.5;margin:0;\">{encoded}</pre>");
        }

        private static string BuildWorkbookPreviewHtml(string filePath)
        {
            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            if (worksheet == null)
                return WrapPreviewHtml(System.IO.Path.GetFileName(filePath), "<p>В книге нет листов для предпросмотра.</p>");

            var range = worksheet.RangeUsed();
            if (range == null)
                return WrapPreviewHtml(System.IO.Path.GetFileName(filePath), "<p>Лист пуст.</p>");

            var startRow = range.RangeAddress.FirstAddress.RowNumber;
            var endRow = Math.Min(range.RangeAddress.LastAddress.RowNumber, startRow + 39);
            var startColumn = range.RangeAddress.FirstAddress.ColumnNumber;
            var endColumn = Math.Min(range.RangeAddress.LastAddress.ColumnNumber, startColumn + 11);

            var html = new System.Text.StringBuilder();
            html.Append("<table style=\"border-collapse:collapse;width:100%;font-size:13px;\">");

            for (var row = startRow; row <= endRow; row++)
            {
                html.Append("<tr>");
                for (var column = startColumn; column <= endColumn; column++)
                {
                    var value = worksheet.Cell(row, column).GetFormattedString();
            html.Append("<td style=\"border:1px solid #E5E7EB;padding:6px 8px;vertical-align:top;\">");
                    html.Append(System.Net.WebUtility.HtmlEncode(value));
                    html.Append("</td>");
                }
                html.Append("</tr>");
            }

            html.Append("</table>");

            if (range.RowCount() > 40 || range.ColumnCount() > 12)
            html.Append("<p style=\"margin-top:12px;color:#6B7280;\">Показана только часть листа для быстрого предпросмотра.</p>");

            return WrapPreviewHtml(System.IO.Path.GetFileName(filePath), html.ToString());
        }

        private static string BuildDocxPreviewHtml(string filePath)
        {
            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null)
                return WrapPreviewHtml(System.IO.Path.GetFileName(filePath), "<p>Документ пуст.</p>");

            var paragraphs = body
                .Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .Select(p => p.InnerText?.Trim())
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Take(80)
                .ToList();

            if (paragraphs.Count == 0)
                return WrapPreviewHtml(System.IO.Path.GetFileName(filePath), "<p>В документе нет текста для предпросмотра.</p>");

            var html = new System.Text.StringBuilder();
            foreach (var paragraph in paragraphs)
            {
                html.Append("<p style=\"margin:0 0 10px 0;line-height:1.55;\">");
                html.Append(System.Net.WebUtility.HtmlEncode(paragraph));
                html.Append("</p>");
            }

            html.Append("<p style=\"margin-top:12px;color:#6B7280;\">Показан текстовый предпросмотр документа.</p>");
            return WrapPreviewHtml(System.IO.Path.GetFileName(filePath), html.ToString());
        }

        private static string WrapPreviewHtml(string title, string body)
        {
            return $@"
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />
    <meta charset=""utf-8"" />
    <title>{System.Net.WebUtility.HtmlEncode(title)}</title>
</head>
<body style=""margin:0;padding:16px;background:#FFFFFF;color:#111827;font-family:'Segoe UI',sans-serif;"">
    {body}
</body>
</html>";
        }

        private void PdfTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            selectedPdfNode = e.NewValue as DocumentTreeNode;
            UpdatePdfSelectionInfo();
        }

        private void EstimateTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            selectedEstimateNode = e.NewValue as DocumentTreeNode;
            UpdateEstimateSelectionInfo();
        }

        private void AddPdfFiles_Click(object sender, RoutedEventArgs e)
        {
            EnsureDocumentLibraries();
            AddDocumentFiles(currentObject?.PdfDocuments, selectedPdfNode, "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*", true);
        }

        private void AddEstimateFiles_Click(object sender, RoutedEventArgs e)
        {
            EnsureDocumentLibraries();
            AddDocumentFiles(currentObject?.EstimateDocuments, selectedEstimateNode, "Estimate files (*.pdf;*.xls;*.xlsx;*.doc;*.docx)|*.pdf;*.xls;*.xlsx;*.doc;*.docx|All files (*.*)|*.*", false);
        }

        private void AddDocumentFiles(List<DocumentTreeNode> root, DocumentTreeNode selectedNode, string filter, bool refreshPdf)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            EnsureDocumentLibraries();
            var dialog = new OpenFileDialog
            {
                Filter = filter,
                Multiselect = true
            };

            if (dialog.ShowDialog() != true || dialog.FileNames.Length == 0)
                return;

            var targetCollection = GetDocumentInsertCollection(root, selectedNode);
            var copyFailedFiles = new List<string>();
            foreach (var file in dialog.FileNames.Distinct(StringComparer.CurrentCultureIgnoreCase))
            {
                TryCopyDocumentToStorage(file, refreshPdf, out var storedPath, out var storedRelativePath);
                if (string.IsNullOrWhiteSpace(storedRelativePath))
                    copyFailedFiles.Add(file);

                targetCollection.Add(new DocumentTreeNode
                {
                    Name = System.IO.Path.GetFileNameWithoutExtension(file),
                    FilePath = storedPath,
                    StoredRelativePath = storedRelativePath,
                    IsFolder = false
                });
            }

            if (refreshPdf)
                selectedPdfNode = targetCollection.LastOrDefault();
            else
                selectedEstimateNode = targetCollection.LastOrDefault();

            SaveState();
            RefreshDocumentLibraries();

            if (copyFailedFiles.Count > 0)
            {
                MessageBox.Show(
                    "Часть файлов не удалось скопировать во внутреннее хранилище проекта. Они добавлены по внешнему пути и могут быть недоступны после переноса проекта.",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void AddPdfFolder_Click(object sender, RoutedEventArgs e)
        {
            EnsureDocumentLibraries();
            AddDocumentFolder(currentObject?.PdfDocuments, selectedPdfNode);
        }

        private void AddEstimateFolder_Click(object sender, RoutedEventArgs e)
        {
            EnsureDocumentLibraries();
            AddDocumentFolder(currentObject?.EstimateDocuments, selectedEstimateNode);
        }

        private void AddDocumentFolder(List<DocumentTreeNode> root, DocumentTreeNode selectedNode)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            var folderName = Microsoft.VisualBasic.Interaction.InputBox("Название папки:", "Новая папка", "Новая папка");
            if (string.IsNullOrWhiteSpace(folderName))
                return;

            var collection = GetDocumentInsertCollection(root, selectedNode);
            collection.Add(new DocumentTreeNode
            {
                Name = folderName.Trim(),
                IsFolder = true
            });

            SaveState();
            RefreshDocumentLibraries();
        }

        private void RenamePdfNode_Click(object sender, RoutedEventArgs e)
            => RenameDocumentNode(currentObject?.PdfDocuments, selectedPdfNode);

        private void RenameEstimateNode_Click(object sender, RoutedEventArgs e)
            => RenameDocumentNode(currentObject?.EstimateDocuments, selectedEstimateNode);

        private void RenameDocumentNode(List<DocumentTreeNode> root, DocumentTreeNode selectedNode)
        {
            if (root == null || selectedNode == null)
            {
                MessageBox.Show("Выберите узел в дереве.");
                return;
            }

            var input = Microsoft.VisualBasic.Interaction.InputBox("Новое название:", "Переименование", selectedNode.Name ?? string.Empty);
            if (string.IsNullOrWhiteSpace(input))
                return;

            selectedNode.Name = input.Trim();
            SaveState();
            RefreshDocumentLibraries();
        }

        private void DeletePdfNode_Click(object sender, RoutedEventArgs e)
            => DeleteDocumentNode(currentObject?.PdfDocuments, selectedPdfNode, isPdf: true);

        private void DeleteEstimateNode_Click(object sender, RoutedEventArgs e)
            => DeleteDocumentNode(currentObject?.EstimateDocuments, selectedEstimateNode, isPdf: false);

        private void DeleteDocumentNode(List<DocumentTreeNode> root, DocumentTreeNode selectedNode, bool isPdf)
        {
            if (root == null || selectedNode == null)
            {
                MessageBox.Show("Выберите узел в дереве.");
                return;
            }

            if (MessageBox.Show($"Удалить \"{selectedNode.Name}\"?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            var collection = GetOwningDocumentCollection(root, selectedNode);
            if (collection == null)
                return;

            collection.Remove(selectedNode);
            if (isPdf)
                selectedPdfNode = null;
            else
                selectedEstimateNode = null;

            SaveState();
            RefreshDocumentLibraries();
        }

        private void MovePdfNodeUp_Click(object sender, RoutedEventArgs e)
            => MoveDocumentNodeInSiblings(currentObject?.PdfDocuments, selectedPdfNode, -1);

        private void MovePdfNodeDown_Click(object sender, RoutedEventArgs e)
            => MoveDocumentNodeInSiblings(currentObject?.PdfDocuments, selectedPdfNode, 1);

        private void MoveEstimateNodeUp_Click(object sender, RoutedEventArgs e)
            => MoveDocumentNodeInSiblings(currentObject?.EstimateDocuments, selectedEstimateNode, -1);

        private void MoveEstimateNodeDown_Click(object sender, RoutedEventArgs e)
            => MoveDocumentNodeInSiblings(currentObject?.EstimateDocuments, selectedEstimateNode, 1);

        private void MoveDocumentNodeInSiblings(List<DocumentTreeNode> root, DocumentTreeNode selectedNode, int delta)
        {
            if (root == null || selectedNode == null)
            {
                MessageBox.Show("Выберите узел в дереве.");
                return;
            }

            var collection = GetOwningDocumentCollection(root, selectedNode);
            if (collection == null)
                return;

            var index = collection.IndexOf(selectedNode);
            var newIndex = index + delta;
            if (index < 0 || newIndex < 0 || newIndex >= collection.Count)
                return;

            collection.RemoveAt(index);
            collection.Insert(newIndex, selectedNode);
            SaveState();
            RefreshDocumentLibraries();
        }

        private void OpenPdfNode_Click(object sender, RoutedEventArgs e)
            => OpenDocumentNode(selectedPdfNode);

        private void OpenEstimateNode_Click(object sender, RoutedEventArgs e)
            => OpenDocumentNode(selectedEstimateNode);

        private void OpenDocumentNode(DocumentTreeNode node)
        {
            if (node == null || node.IsFolder)
            {
                MessageBox.Show("Выберите файл, а не папку.");
                return;
            }

            var resolvedPath = ResolveDocumentPath(node);
            if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath))
            {
                MessageBox.Show("Файл не найден по сохраненному пути.");
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = resolvedPath,
                UseShellExecute = true
            });
        }

        private void EstimatePreviewContainer_SizeChanged(object sender, SizeChangedEventArgs e)
            => LayoutEmbeddedEstimateWindow();

        private void PdfTreeView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            pdfTreeDragStart = e.GetPosition(null);
            pdfDragNode = GetDocumentNodeFromOriginalSource(e.OriginalSource);
        }

        private void PdfTreeView_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || pdfDragNode == null)
                return;

            var position = e.GetPosition(null);
            if (Math.Abs(position.X - pdfTreeDragStart.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(position.Y - pdfTreeDragStart.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            DragDrop.DoDragDrop(PdfTreeView, pdfDragNode, DragDropEffects.Move);
            pdfDragNode = null;
        }

        private void PdfTreeView_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(DocumentTreeNode)) is not DocumentTreeNode sourceNode)
                return;

            MoveDocumentNodeByDrop(currentObject?.PdfDocuments, sourceNode, GetDocumentNodeFromOriginalSource(e.OriginalSource));
        }

        private void EstimateTreeView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            estimateTreeDragStart = e.GetPosition(null);
            estimateDragNode = GetDocumentNodeFromOriginalSource(e.OriginalSource);
        }

        private void EstimateTreeView_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || estimateDragNode == null)
                return;

            var position = e.GetPosition(null);
            if (Math.Abs(position.X - estimateTreeDragStart.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(position.Y - estimateTreeDragStart.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            DragDrop.DoDragDrop(EstimateTreeView, estimateDragNode, DragDropEffects.Move);
            estimateDragNode = null;
        }

        private void EstimateTreeView_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(DocumentTreeNode)) is not DocumentTreeNode sourceNode)
                return;

            MoveDocumentNodeByDrop(currentObject?.EstimateDocuments, sourceNode, GetDocumentNodeFromOriginalSource(e.OriginalSource));
        }

        private void MoveDocumentNodeByDrop(List<DocumentTreeNode> root, DocumentTreeNode sourceNode, DocumentTreeNode targetNode)
        {
            if (root == null || sourceNode == null || targetNode == null)
                return;

            if (!ContainsDocumentNode(root, sourceNode))
                return;

            if (ReferenceEquals(sourceNode, targetNode) || IsDocumentDescendant(sourceNode, targetNode))
                return;

            var sourceCollection = GetOwningDocumentCollection(root, sourceNode);
            var targetCollection = GetDocumentInsertCollection(root, targetNode);
            if (sourceCollection == null || targetCollection == null)
                return;

            sourceCollection.Remove(sourceNode);
            targetCollection.Add(sourceNode);
            SaveState();
            RefreshDocumentLibraries();
        }

        private void UpdateArrivalViewMode()
        {
            if (ArrivalLegacyGrid == null || ArrivalMatrixScrollViewer == null)
                return;

            ArrivalLegacyGrid.Visibility = arrivalMatrixMode ? Visibility.Collapsed : Visibility.Visible;
            ArrivalMatrixScrollViewer.Visibility = arrivalMatrixMode ? Visibility.Visible : Visibility.Collapsed;

            if (ArrivalLegacyViewButton != null)
                ArrivalLegacyViewButton.Opacity = arrivalMatrixMode ? 0.72 : 1.0;

            if (ArrivalMatrixViewButton != null)
                ArrivalMatrixViewButton.Opacity = arrivalMatrixMode ? 1.0 : 0.72;

            if (ArrivalExtraSubtypeFiltersPanel != null)
                ArrivalExtraSubtypeFiltersPanel.Visibility = arrivalMatrixMode ? Visibility.Collapsed : Visibility.Visible;
        }

        private void ArrivalLegacyViewButton_Click(object sender, RoutedEventArgs e)
        {
            arrivalMatrixMode = false;
            UpdateArrivalViewMode();
            ApplyAllFilters();
        }

        private void ArrivalMatrixViewButton_Click(object sender, RoutedEventArgs e)
        {
            arrivalMatrixMode = true;
            SyncArrivalMatrixSelectionWithTree();
            if (selectedArrivalTypes.Count > 1)
            {
                var selected = selectedArrivalTypes.OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase).First();
                selectedArrivalTypes.Clear();
                selectedArrivalTypes.Add(selected);
                RefreshArrivalTypes();
                RefreshArrivalNames();
            }

            UpdateArrivalViewMode();
            ApplyAllFilters();
        }
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source is not TabControl tab || tab.SelectedItem is not TabItem item)
                return;

            UpdateTabButtons();
            UpdateTreePanelState(forceVisible: isTreePinned);
            UpdatePdfTreePanelState(forceVisible: isPdfTreePinned);
            UpdateEstimateTreePanelState(forceVisible: isEstimateTreePinned);

            if (!ReferenceEquals(item, EstimateTab))
                StopEstimateEmbeddedPreview();

            if (item.Header?.ToString() == "Приход")
            {
                ArrivalPopup.Visibility = arrivalPanelVisible
                    ? Visibility.Visible
                    : Visibility.Collapsed;

                ShowArrivalButton.Visibility = arrivalPanelVisible
                    ? Visibility.Collapsed
                    : Visibility.Visible;

                UpdateArrivalViewMode();
                if (initialUiPrepared)
                {
                    if (arrivalMatrixMode)
                        RenderArrivalMatrix();
                    else
                        ArrivalLegacyGrid.ItemsSource = filteredJournal;
                }

            }
            else
            {
                ArrivalPopup.Visibility = Visibility.Collapsed;
                ShowArrivalButton.Visibility = Visibility.Visible;
            }

            RequestReminderRefresh();
            var view = CollectionViewSource.GetDefaultView(OtJournalGrid.ItemsSource);

            if (view is IEditableCollectionView editable)
            {
                if (editable.IsAddingNew)
                    editable.CommitNew();

                if (editable.IsEditingItem)
                    editable.CommitEdit();
            }

            view?.Refresh();
            if (item.Header?.ToString() == "Табель")
            {
                if (timesheetNeedsRebuild || TimesheetGrid?.ItemsSource == null)
                    RebuildTimesheetView(force: true);
            }
            if (item.Header?.ToString() == "ПР")
                RefreshProductionJournalState();
            if (item.Header?.ToString() == "Осмотры")
                RefreshInspectionJournalState();
            if (ReferenceEquals(item, EstimateTab))
                UpdateEstimateSelectionInfo();
        }

        private void ShowArrivalButton_Click(object sender, RoutedEventArgs e)
        {
            arrivalPanelVisible = true;
            ArrivalPopup.Visibility = Visibility.Visible;
            ShowArrivalButton.Visibility = Visibility.Collapsed;
        }

        private void HideArrivalButton_Click(object sender, RoutedEventArgs e)
        {
            arrivalPanelVisible = false;
            ArrivalPopup.Visibility = Visibility.Collapsed;
            ShowArrivalButton.Visibility = Visibility.Visible;
        }

        private void InitializeOtJournal()
        {
            EnsureOtJournalStorage();
            BindOtJournal();
        }

        private void EnsureOtJournalStorage()
        {
            if (currentObject != null && currentObject.OtJournal == null)
                currentObject.OtJournal = new List<OtJournalEntry>();
        }

        private void BindOtJournal()
        {
            OtJournalGrid.ItemsSource = currentObject?.OtJournal;
            if (currentObject?.OtJournal == null)
            {
                OtJournalGrid.ItemsSource = null;
                return;
            }

            var view = CollectionViewSource.GetDefaultView(currentObject.OtJournal);
            view.Filter = OtJournalFilter;
            OtJournalGrid.ItemsSource = view;
            SubscribeOtJournalEntryEvents();
            RefreshBrigadierNames();
            RefreshSpecialties();
            RefreshProfessions();
            NormalizeOtRows();
            SortOtJournal();
            RequestReminderRefresh();
            MarkTimesheetOtSyncDirty();
            RequestTimesheetRebuild();
        }
        private bool OtJournalFilter(object item)
        {
            if (item is not OtJournalEntry row)
                return false;

            if (string.IsNullOrWhiteSpace(otSearchText))
                return true;

            return (row.FullName ?? string.Empty).Contains(otSearchText, StringComparison.CurrentCultureIgnoreCase);
        }
        private void SubscribeOtJournalEntryEvents()
        {
            if (currentObject?.OtJournal == null)
                return;

            foreach (var item in currentObject.OtJournal)
            {
                item.PropertyChanged -= OtJournalEntry_PropertyChanged;
                item.PropertyChanged += OtJournalEntry_PropertyChanged;
            }
        }

        private void OtJournalEntry_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is not OtJournalEntry row)
                return;

            if (e.PropertyName == nameof(OtJournalEntry.FullName))
                WarnAboutDuplicatePerson(row);

            if (e.PropertyName == nameof(OtJournalEntry.Specialty))
                RefreshSpecialties();
            if (e.PropertyName == nameof(OtJournalEntry.Profession))
                RefreshProfessions();
            if (e.PropertyName == nameof(OtJournalEntry.Specialty)
                || e.PropertyName == nameof(OtJournalEntry.Profession))
            {
                SyncProfessionAndSpecialty(row, e.PropertyName);
                FillInstructionNumbersFromTemplate(row);
            }

            if (e.PropertyName == nameof(OtJournalEntry.InstructionDate)
                || e.PropertyName == nameof(OtJournalEntry.RepeatPeriodMonths)
                || e.PropertyName == nameof(OtJournalEntry.IsScheduledRepeat)
                || e.PropertyName == nameof(OtJournalEntry.IsPendingRepeat)
                || e.PropertyName == nameof(OtJournalEntry.IsDismissed))
            {
                NormalizeOtRows();
                SortOtJournal();
                RequestReminderRefresh();
            }

            if (!isSyncingTimesheetToOt && IsOtPropertyAffectingTimesheet(e.PropertyName))
            {
                MarkTimesheetOtSyncDirty();
                RequestTimesheetRebuild();
            }
        }

        private static bool IsOtPropertyAffectingTimesheet(string propertyName)
        {
            return propertyName == nameof(OtJournalEntry.FullName)
                || propertyName == nameof(OtJournalEntry.Specialty)
                || propertyName == nameof(OtJournalEntry.Rank)
                || propertyName == nameof(OtJournalEntry.IsBrigadier)
                || propertyName == nameof(OtJournalEntry.BrigadierName)
                || propertyName == nameof(OtJournalEntry.IsDismissed)
                || propertyName == nameof(OtJournalEntry.PersonId);
        }
        private void RefreshSpecialties()
        {
            specialties.Clear();

            if (currentObject?.OtJournal == null)
                return;

            foreach (var item in currentObject.OtJournal
                         .Where(x => !string.IsNullOrWhiteSpace(x.Specialty))
                         .Select(x => x.Specialty.Trim())
                         .Distinct(StringComparer.CurrentCultureIgnoreCase)
                         .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
            {
                specialties.Add(item);
            }
        }
        private void RefreshProfessions()
        {
            professions.Clear();

            if (currentObject?.OtJournal == null)
                return;

            foreach (var item in currentObject.OtJournal
                         .Where(x => !string.IsNullOrWhiteSpace(x.Profession))
                         .Select(x => x.Profession.Trim())
                         .Distinct(StringComparer.CurrentCultureIgnoreCase)
                         .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
            {
                professions.Add(item);
            }
        }

        private void SyncProfessionAndSpecialty(OtJournalEntry row, string changedPropertyName)
        {
            if (row == null)
                return;

            if (changedPropertyName == nameof(OtJournalEntry.Specialty)
                && string.IsNullOrWhiteSpace(row.Profession)
                && !string.IsNullOrWhiteSpace(row.Specialty))
            {
                row.Profession = row.Specialty.Trim();
                return;
            }

            if (changedPropertyName == nameof(OtJournalEntry.Profession)
                && string.IsNullOrWhiteSpace(row.Specialty)
                && !string.IsNullOrWhiteSpace(row.Profession))
            {
                row.Specialty = row.Profession.Trim();
            }
        }

        private void FillInstructionNumbersFromTemplate(OtJournalEntry row)
        {
            if (currentObject?.OtJournal == null || row == null)
                return;

            if (!string.IsNullOrWhiteSpace(row.InstructionNumbers))
                return;

            var key = string.IsNullOrWhiteSpace(row.Profession)
                ? row.Specialty?.Trim()
                : row.Profession.Trim();

            if (string.IsNullOrWhiteSpace(key))
                return;

            var template = currentObject.OtJournal
                .Where(x => !ReferenceEquals(x, row))
                .FirstOrDefault(x =>
                    !string.IsNullOrWhiteSpace(x.InstructionNumbers)
                    && (string.Equals(x.Profession?.Trim(), key, StringComparison.CurrentCultureIgnoreCase)
                        || string.Equals(x.Specialty?.Trim(), key, StringComparison.CurrentCultureIgnoreCase)));

            if (template != null)
                row.InstructionNumbers = template.InstructionNumbers;
        }
        private void RefreshBrigadierNames()
        {
            brigadierNames.Clear();

            if (currentObject?.OtJournal == null)
                return;

            foreach (var name in currentObject.OtJournal
                         .Where(x => x.IsBrigadier && !string.IsNullOrWhiteSpace(x.FullName))
                         .Select(x => x.FullName.Trim())
                         .Distinct(StringComparer.CurrentCultureIgnoreCase)
                         .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
            {
                brigadierNames.Add(name);
            }
        }
        private void NormalizeOtRows()
        {
            if (currentObject?.OtJournal == null)
                return;

            foreach (var scheduledRow in currentObject.OtJournal
                         .Where(x => !x.IsDismissed && x.IsScheduledRepeat && DateTime.Today >= x.InstructionDate)
                         .ToList())
            {
                scheduledRow.IsScheduledRepeat = false;
                scheduledRow.IsPendingRepeat = true;
                scheduledRow.IsRepeatCompleted = false;
            }

            var toAdd = new List<OtJournalEntry>();

            foreach (var group in currentObject.OtJournal
                         .Where(x => !string.IsNullOrWhiteSpace(x.FullName))
                         .GroupBy(x => x.FullName.Trim(), StringComparer.CurrentCultureIgnoreCase))
            {
                var activeRows = group.Where(x => !x.IsDismissed).ToList();
                if (!activeRows.Any())
                    continue;

                var pendingExists = activeRows.Any(x => x.IsPendingRepeat);
                if (pendingExists)
                    continue;

                var lastCompleted = activeRows
                    .Where(x => !x.IsPendingRepeat && !x.IsScheduledRepeat)
                    .OrderByDescending(x => x.InstructionDate)
                    .FirstOrDefault();

                if (lastCompleted == null)
                    continue;

                var repeatDate = lastCompleted.NextRepeatDate;
                if (DateTime.Today < repeatDate)
                    continue;

                var clone = new OtJournalEntry
                {
                    PersonId = lastCompleted.PersonId,
                    InstructionDate = repeatDate,
                    FullName = lastCompleted.FullName,
                    Specialty = lastCompleted.Specialty,
                    Rank = lastCompleted.Rank,
                    Profession = lastCompleted.Profession,
                    InstructionType = BuildRepeatInstructionType(group.Count(IsRepeatInstruction) + 1),
                    InstructionNumbers = lastCompleted.InstructionNumbers,
                    RepeatPeriodMonths = Math.Max(1, lastCompleted.RepeatPeriodMonths),
                    IsBrigadier = lastCompleted.IsBrigadier,
                    BrigadierName = lastCompleted.BrigadierName,
                    IsPendingRepeat = true,
                    IsScheduledRepeat = false,
                    IsRepeatCompleted = false,
                    IsDismissed = false,
                };
                clone.PropertyChanged += OtJournalEntry_PropertyChanged;
                toAdd.Add(clone);
            }

            if (toAdd.Count > 0)
                currentObject.OtJournal.AddRange(toAdd);
        }

        private void SortOtJournal()
        {
            if (currentObject?.OtJournal == null)
                return;

            currentObject.OtJournal = currentObject.OtJournal
                .OrderByDescending(x => x.InstructionDate)
                .ThenBy(x => x.FullName ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var view = CollectionViewSource.GetDefaultView(currentObject.OtJournal);
            view.Filter = OtJournalFilter;
            OtJournalGrid.ItemsSource = view;
            view.Refresh();
        }

        private void WarnAboutDuplicatePerson(OtJournalEntry source)
        {
            if (currentObject?.OtJournal == null || string.IsNullOrWhiteSpace(source.FullName))
                return;

            var samePeople = currentObject.OtJournal
                .Where(x => !ReferenceEquals(x, source)
                            && !string.IsNullOrWhiteSpace(x.FullName)
                            && string.Equals(x.FullName.Trim(), source.FullName.Trim(), StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            if (!samePeople.Any())
                return;

            var wasDismissed = samePeople.Any(x => x.IsDismissed);
            var message = wasDismissed
                ? "Сотрудник был ранее отмечен как отсутствующий на объекте. При возвращении ему требуется повторный инструктаж."
                : "Сотрудник с таким ФИО уже есть в журнале ОТ.";

            MessageBox.Show(message, "Уведомление", MessageBoxButton.OK, MessageBoxImage.Warning);

            if (wasDismissed)
            {
                source.InstructionType = BuildRepeatInstructionType(GetNextRepeatIndexForPerson(source.FullName));
                source.IsPendingRepeat = true;
                source.IsRepeatCompleted = false;
            }
        }
        private static bool IsRepeatInstruction(OtJournalEntry entry)
    => !string.IsNullOrWhiteSpace(entry?.InstructionType)
        && entry.InstructionType.Contains("повторн", StringComparison.CurrentCultureIgnoreCase);

        private static string BuildRepeatInstructionType(int index)
            => index <= 1 ? "Повторный" : $"Повторный ({index})";

        private int GetNextRepeatIndexForPerson(string fullName)
        {
            if (currentObject?.OtJournal == null || string.IsNullOrWhiteSpace(fullName))
                return 1;

            return currentObject.OtJournal
                .Where(x => !string.IsNullOrWhiteSpace(x.FullName)
                            && string.Equals(x.FullName.Trim(), fullName.Trim(), StringComparison.CurrentCultureIgnoreCase)
                            && IsRepeatInstruction(x))
                .Count() + 1;
        }

        private int GetRepeatIndexForRow(OtJournalEntry row)
        {
            if (currentObject?.OtJournal == null || row == null || string.IsNullOrWhiteSpace(row.FullName))
                return 1;

            var samePersonRepeats = currentObject.OtJournal
                .Where(x => !string.IsNullOrWhiteSpace(x.FullName)
                            && string.Equals(x.FullName.Trim(), row.FullName.Trim(), StringComparison.CurrentCultureIgnoreCase)
                            && IsRepeatInstruction(x))
                .OrderBy(x => x.InstructionDate)
                .ToList();

            var idx = samePersonRepeats.IndexOf(row);
            return idx >= 0 ? idx + 1 : samePersonRepeats.Count + 1;
        }
        private int GetReminderSnoozeMinutes()
        {
            EnsureProjectUiSettings();

            var minutes = currentObject?.UiSettings?.ReminderSnoozeMinutes ?? 15;
            if (minutes < 1)
                return 1;
            if (minutes > 240)
                return 240;
            return minutes;
        }

        private static bool IsDiscreteUnit(string unit)
        {
            if (string.IsNullOrWhiteSpace(unit))
                return false;

            var normalized = unit.Trim().ToLowerInvariant();
            return normalized is "шт" or "шт." or "piece" or "pieces" or "pc" or "pcs";
        }

        private static double NormalizeQuantityByUnit(double value, string unit)
        {
            if (!IsDiscreteUnit(unit))
                return Math.Max(0, value);

            return Math.Max(0, Math.Round(value, 0, MidpointRounding.AwayFromZero));
        }

        private string FormatNumberByUnit(double value, string unit)
            => IsDiscreteUnit(unit)
                ? NormalizeQuantityByUnit(value, unit).ToString("0", CultureInfo.CurrentCulture)
                : FormatNumber(value);

        private List<string> CollectTimesheetMissingDocsPreview(out int count)
        {
            count = 0;
            var preview = new List<string>();

            if (currentObject?.TimesheetPeople == null)
                return preview;

            var monthKey = timesheetMonth.ToString("yyyy-MM");
            var daysInMonth = DateTime.DaysInMonth(timesheetMonth.Year, timesheetMonth.Month);

            foreach (var person in currentObject.TimesheetPeople)
            {
                for (var day = 1; day <= daysInMonth; day++)
                {
                    if (!person.IsNonHourCode(monthKey, day))
                        continue;

                    if (person.GetDocumentAccepted(monthKey, day) == true)
                        continue;

                    count++;
                    preview.Add($"{person.FullName} — {day:00}.{timesheetMonth.Month:00}.{timesheetMonth.Year}");

                    // Ограничиваем объём по запросу пользователя: максимум 3 отсутствующих с требованием документа.
                    if (count >= MaxTimesheetMissingDocs)
                        return preview;
                }
            }

            return preview;
        }

        private void SetReminderPopupVisible(bool visible)
        {
            if (!visible)
            {
                HideReminderOverlayWindow();
                return;
            }

            if (!IsLoaded || !IsVisible)
                return;

            var overlay = EnsureReminderOverlayWindow();
            if (WindowState == WindowState.Minimized)
                return;

            if (!overlay.IsLoaded && overlay.Owner == null)
                overlay.Owner = this;

            if (!overlay.IsVisible)
                overlay.Show();

            UpdateReminderOverlayPlacement();
            Dispatcher.BeginInvoke(new Action(UpdateReminderOverlayPlacement), DispatcherPriority.Loaded);
        }

        private void HideReminderOverlayWindow()
        {
            if (reminderOverlayWindow?.IsVisible == true)
                reminderOverlayWindow.Hide();
        }

        private void UpdateReminderOverlayPlacement()
        {
            if (reminderOverlayWindow == null || !reminderOverlayWindow.IsVisible || MainRootGrid == null || WindowState == WindowState.Minimized)
                return;

            var source = PresentationSource.FromVisual(this);
            if (source?.CompositionTarget == null)
                return;

            const double margin = 12.0;
            var fromDevice = source.CompositionTarget.TransformFromDevice;
            var topLeft = fromDevice.Transform(MainRootGrid.PointToScreen(new Point(0, 0)));
            var bottomRight = fromDevice.Transform(MainRootGrid.PointToScreen(new Point(MainRootGrid.ActualWidth, MainRootGrid.ActualHeight)));

            var root = reminderOverlayWindow.RootElement;
            root.Measure(new Size(reminderOverlayWindow.Width, double.PositiveInfinity));
            var height = root.DesiredSize.Height;
            if (height <= 0)
            {
                reminderOverlayWindow.UpdateLayout();
                height = reminderOverlayWindow.ActualHeight;
            }

            var width = reminderOverlayWindow.Width > 0 ? reminderOverlayWindow.Width : reminderOverlayWindow.ActualWidth;
            reminderOverlayWindow.Left = Math.Max(topLeft.X + margin, bottomRight.X - width - margin);
            reminderOverlayWindow.Top = Math.Max(topLeft.Y + margin, bottomRight.Y - height - margin);
        }

        private void UpdateOtReminders()
        {
            if (currentObject == null)
            {
                reminderSections.Clear();
                SetReminderPopupVisible(false);
                return;
            }

            EnsureProjectUiSettings();
            if (currentObject.UiSettings?.ShowReminderPopup == false)
            {
                reminderSections.Clear();
                SetReminderPopupVisible(false);
                return;
            }

            if (reminderSnoozedUntil.HasValue && reminderSnoozedUntil.Value <= DateTime.Now)
                reminderSnoozedUntil = null;

            var sections = new List<ReminderSectionViewModel>();
            var totalCount = 0;

            var dueRows = currentObject.OtJournal?
                .Where(x => x.IsPendingRepeat && !x.IsDismissed)
                .ToList() ?? new List<OtJournalEntry>();
            if (dueRows.Count > 0)
            {
                totalCount += dueRows.Count;
                sections.Add(new ReminderSectionViewModel
                {
                    Header = "Вкладка ОТ",
                    Items = dueRows
                        .Take(4)
                        .Select(x =>
                        {
                            var person = !string.IsNullOrWhiteSpace(x.LastName) ? x.LastName : x.FullName;
                            return $"Нужен повторный инструктаж: {person}";
                        })
                        .ToList()
                });
            }

            var missingDocsPreview = CollectTimesheetMissingDocsPreview(out var missingDocsCount);
            if (missingDocsCount > 0)
            {
                totalCount += missingDocsCount;
                var tabItems = missingDocsPreview
                    .Take(MaxTimesheetMissingDocs)
                    .Select(x => $"Нет документа: {x}")
                    .ToList();
                if (missingDocsCount >= MaxTimesheetMissingDocs)
                    tabItems.Add("Порог контроля: максимум 3 отсутствующих с документами.");

                sections.Add(new ReminderSectionViewModel
                {
                    Header = "Вкладка Табель",
                    Items = tabItems
                });
            }

            var dueInspections = currentObject.InspectionJournal?
                .Where(x => x.IsDue)
                .OrderBy(x => x.NextReminderDate)
                .ToList() ?? new List<InspectionJournalEntry>();
            if (dueInspections.Count > 0)
            {
                totalCount += dueInspections.Count;
                sections.Add(new ReminderSectionViewModel
                {
                    Header = "Вкладка Осмотры",
                    Items = dueInspections
                        .Take(4)
                        .Select(x => $"{x.JournalName}: {x.InspectionName} (с {x.NextReminderDate:dd.MM.yyyy})")
                        .ToList()
                });
            }

            var summaryBalanceItems = CollectSummaryBalanceReminderItems();
            if (summaryBalanceItems.Count > 0)
            {
                totalCount += summaryBalanceItems.Count;
                var summaryPreview = summaryBalanceItems
                    .Take(4)
                    .Select(x =>
                    {
                        var deltaText = x.IsOverage ? "Переход" : "Недоход";
                        return $"Тип: {x.Group}; Наименование: {x.Material}; {deltaText}: {FormatNumberByUnit(x.Quantity, x.Unit)}";
                    })
                    .ToList();

                if (summaryBalanceItems.Count > 4)
                    summaryPreview.Add($"И еще позиций: {summaryBalanceItems.Count - 4}.");

                sections.Add(new ReminderSectionViewModel
                {
                    Header = "Вкладка Сводка",
                    Items = summaryPreview
                });
            }

            if (totalCount == 0)
            {
                reminderSections.Clear();
                SetReminderPopupVisible(false);
                return;
            }

            reminderSections.Clear();
            foreach (var section in sections)
                reminderSections.Add(section);

            var overlay = EnsureReminderOverlayWindow();
            overlay.SnoozeButtonElement.Content = $"Отложить ({GetReminderSnoozeMinutes()} мин)";

            var isSnoozed = reminderSnoozedUntil.HasValue && reminderSnoozedUntil.Value > DateTime.Now;
            var detailsHidden = currentObject.UiSettings?.HideReminderDetails == true;
            overlay.ToggleDetailsButtonElement.Content = detailsHidden ? "Показать" : "Скрыть";

            if (isSnoozed)
            {
                // По запросу: при отложении уведомления полностью скрываются до окончания срока.
                SetReminderPopupVisible(false);
                return;
            }
            else if (detailsHidden)
            {
                overlay.StateTextElement.Text = "Уведомления скрыты. Нажмите «Показать», чтобы увидеть детали.";
                overlay.StateTextElement.Visibility = Visibility.Visible;
                overlay.SectionsHostElement.Visibility = Visibility.Collapsed;
            }
            else
            {
                overlay.StateTextElement.Visibility = Visibility.Collapsed;
                overlay.SectionsHostElement.Visibility = Visibility.Visible;
            }

            SetReminderPopupVisible(true);
        }

        private List<SummaryBalanceReminderItem> CollectSummaryBalanceReminderItems()
        {
            var result = new List<SummaryBalanceReminderItem>();
            if (currentObject == null)
                return result;

            var keyComparer = StringComparer.CurrentCultureIgnoreCase;
            var candidates = new HashSet<(string Group, string Material)>();

            var mainCatalogPairs = (currentObject.MaterialCatalog ?? new List<MaterialCatalogItem>())
                .Where(x => x != null
                         && string.Equals((x.CategoryName ?? string.Empty).Trim(), "Основные", StringComparison.CurrentCultureIgnoreCase)
                         && !string.IsNullOrWhiteSpace(x.TypeName)
                         && !string.IsNullOrWhiteSpace(x.MaterialName))
                .Select(x => $"{x.TypeName.Trim()}||{x.MaterialName.Trim()}")
                .ToHashSet(keyComparer);

            var nonMainCatalogPairs = (currentObject.MaterialCatalog ?? new List<MaterialCatalogItem>())
                .Where(x => x != null
                         && !string.Equals((x.CategoryName ?? string.Empty).Trim(), "Основные", StringComparison.CurrentCultureIgnoreCase)
                         && !string.IsNullOrWhiteSpace(x.TypeName)
                         && !string.IsNullOrWhiteSpace(x.MaterialName))
                .Select(x => $"{x.TypeName.Trim()}||{x.MaterialName.Trim()}")
                .ToHashSet(keyComparer);

            bool IsMainCandidate(string group, string material)
            {
                var key = $"{group}||{material}";
                if (mainCatalogPairs.Contains(key))
                    return true;

                if (nonMainCatalogPairs.Contains(key))
                    return false;

                var hasMainRecord = journal.Any(x =>
                    string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals((x.MaterialGroup ?? string.Empty).Trim(), group, StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase));
                if (hasMainRecord)
                    return true;

                var hasNonMainRecord = journal.Any(x =>
                    !string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals((x.MaterialGroup ?? string.Empty).Trim(), group, StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase));
                if (hasNonMainRecord)
                    return false;

                return false;
            }

            if (currentObject.Demand != null)
            {
                foreach (var key in currentObject.Demand.Keys.Where(x => !string.IsNullOrWhiteSpace(x)))
                {
                    var parts = key.Split(new[] { "::" }, 2, StringSplitOptions.None);
                    if (parts.Length == 2
                        && !string.IsNullOrWhiteSpace(parts[0])
                        && !string.IsNullOrWhiteSpace(parts[1]))
                    {
                        var group = parts[0].Trim();
                        var material = parts[1].Trim();
                        if (IsMainCandidate(group, material))
                            candidates.Add((group, material));
                    }
                }
            }

            foreach (var row in journal.Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)))
            {
                var group = row.MaterialGroup?.Trim();
                var material = row.MaterialName?.Trim();
                if (!string.IsNullOrWhiteSpace(group) && !string.IsNullOrWhiteSpace(material))
                {
                    if (!IsMainCandidate(group, material))
                        continue;
                    candidates.Add((group, material));
                }
            }

            if (currentObject.MaterialCatalog != null)
            {
                foreach (var item in currentObject.MaterialCatalog.Where(x =>
                             string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)))
                {
                    var group = item.TypeName?.Trim();
                    var material = item.MaterialName?.Trim();
                    if (!string.IsNullOrWhiteSpace(group) && !string.IsNullOrWhiteSpace(material))
                    {
                        if (!IsMainCandidate(group, material))
                            continue;
                        candidates.Add((group, material));
                    }
                }
            }

            foreach (var (group, material) in candidates
                         .OrderBy(x => x.Group, keyComparer)
                         .ThenBy(x => x.Material, keyComparer))
            {
                var records = journal
                    .Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals((x.MaterialGroup ?? string.Empty).Trim(), group, StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase))
                    .ToList();

                var unit = records
                    .Select(x => x.Unit)
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))
                    ?? GetUnitForMaterial(group, material);

                var totalArrived = NormalizeQuantityByUnit(records.Sum(x => x.Quantity), unit);
                var demand = GetOrCreateDemand(BuildDemandKey(group, material), unit);
                var totalNeed = NormalizeQuantityByUnit(
                    BuildSummaryBlocks(group)
                        .SelectMany(x => x.Levels.Select(level => GetDemandValue(demand, x.Block, level)))
                        .Sum(),
                    unit);

                var delta = totalArrived - totalNeed;
                if (delta <= 0.0001)
                    continue;

                result.Add(new SummaryBalanceReminderItem
                {
                    Category = "Основные",
                    Group = group,
                    Material = material,
                    Unit = unit,
                    Quantity = delta,
                    IsOverage = true
                });
            }

            return result;
        }

        private void SnoozeRemindersButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
                return;

            reminderSnoozedUntil = DateTime.Now.AddMinutes(GetReminderSnoozeMinutes());
            RequestReminderRefresh(immediate: true);
        }

        private void ToggleReminderDetailsButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
                return;

            EnsureProjectUiSettings();
            currentObject.UiSettings.HideReminderDetails = !(currentObject.UiSettings.HideReminderDetails);
            SaveState();
            RequestReminderRefresh(immediate: true);
        }

        private void AddOtRow_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            EnsureOtJournalStorage();

            var row = new OtJournalEntry();
            row.PropertyChanged += OtJournalEntry_PropertyChanged;
            currentObject.OtJournal.Add(row);
            SortOtJournal();

            OtJournalGrid.Items.Refresh();
            OtJournalGrid.SelectedItem = row;
            RefreshSpecialties();
            RefreshProfessions();
            RequestReminderRefresh();
            SaveState();
            MarkTimesheetOtSyncDirty();
            RequestTimesheetRebuild();
        }
        private void MarkRepeatDone(OtJournalEntry row)
        {
            if (row == null)
                return;

            if (!row.IsActionEnabled)
            {
                MessageBox.Show("Для первичного инструктажа действие заблокировано. Отмечайте выполнение только в строке повторного инструктажа.");
                return;
            }

            row.InstructionDate = DateTime.Today;
            row.InstructionType = BuildRepeatInstructionType(GetRepeatIndexForRow(row));
            row.IsPendingRepeat = false;
            row.IsScheduledRepeat = false;
            row.IsRepeatCompleted = true;

            var nextDate = row.NextRepeatDate;
            var nextIndex = GetNextRepeatIndexForPerson(row.FullName);
            var hasScheduled = currentObject?.OtJournal?.Any(x =>
                !ReferenceEquals(x, row)
                && !x.IsDismissed
                && string.Equals(x.FullName?.Trim(), row.FullName?.Trim(), StringComparison.CurrentCultureIgnoreCase)
                && x.IsScheduledRepeat
                && x.InstructionDate.Date == nextDate.Date) == true;

            if (!hasScheduled && currentObject?.OtJournal != null)
            {
                var nextRow = new OtJournalEntry
                {
                    PersonId = row.PersonId,
                    InstructionDate = nextDate,
                    FullName = row.FullName,
                    Specialty = row.Specialty,
                    Rank = row.Rank,
                    Profession = row.Profession,
                    InstructionType = BuildRepeatInstructionType(nextIndex),
                    InstructionNumbers = row.InstructionNumbers,
                    RepeatPeriodMonths = Math.Max(1, row.RepeatPeriodMonths),
                    IsBrigadier = row.IsBrigadier,
                    BrigadierName = row.BrigadierName,
                    IsPendingRepeat = false,
                    IsScheduledRepeat = true,
                    IsRepeatCompleted = false,
                    IsDismissed = false
                };
                nextRow.PropertyChanged += OtJournalEntry_PropertyChanged;
                currentObject.OtJournal.Add(nextRow);
            }

            NormalizeOtRows();
            SortOtJournal();
            RequestReminderRefresh();
            OtJournalGrid.Items.Refresh();
            SaveState();
        }
        private void MarkRepeatDoneRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is OtJournalEntry row)
                MarkRepeatDone(row);
        }

        private void MarkSelectedRepeatDone_Click(object sender, RoutedEventArgs e)
        {
            if (OtJournalGrid.SelectedItem is not OtJournalEntry row)
            {
                MessageBox.Show("Выберите запись в таблице ОТ");
                return;
            }
            MarkRepeatDone(row);
        }

        private void MarkPersonDismissed(OtJournalEntry row)
        {
            if (row == null || currentObject?.OtJournal == null || string.IsNullOrWhiteSpace(row.FullName))
                return;

            var rows = currentObject.OtJournal
                .Where(x => !string.IsNullOrWhiteSpace(x.FullName)
                            && string.Equals(x.FullName.Trim(), row.FullName.Trim(), StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            foreach (var item in rows)
            {
                item.IsDismissed = true;
                item.IsPendingRepeat = false;
                item.IsScheduledRepeat = false;
            }

            RequestReminderRefresh();
            MarkTimesheetOtSyncDirty();
            RequestTimesheetRebuild();
            SaveState();
            MessageBox.Show("Сотрудник отмечен как отсутствующий на объекте.");
        }

        private void MarkPersonDismissedRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is OtJournalEntry row)
                MarkPersonDismissed(row);
        }

        private void MarkSelectedPersonDismissed_Click(object sender, RoutedEventArgs e)
        {
            if (OtJournalGrid.SelectedItem is not OtJournalEntry row)
            {
                MessageBox.Show("Выберите запись в таблице ОТ");
                return;
            }

            MarkPersonDismissed(row);
        }

        private void DeleteSelectedOtRow_Click(object sender, RoutedEventArgs e)
        {
            if (OtJournalGrid.SelectedItem is not OtJournalEntry row || currentObject?.OtJournal == null)
                return;

            currentObject.OtJournal.Remove(row);
            SortOtJournal();
            RefreshBrigadierNames();
            RefreshSpecialties();
            RefreshProfessions();
            RequestReminderRefresh();
            MarkTimesheetOtSyncDirty();
            RequestTimesheetRebuild();
            SaveState();
        }
        private void OtSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            otSearchText = OtSearchTextBox.Text?.Trim() ?? string.Empty;
            CollectionViewSource.GetDefaultView(OtJournalGrid.ItemsSource)?.Refresh();
        }

        private void OtJournalGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit)
                return;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                RefreshBrigadierNames();
                RefreshSpecialties();
                RefreshProfessions();
                MarkTimesheetOtSyncDirty();
                RequestTimesheetRebuild();
                RequestReminderRefresh();
            }));
        }

        private void OtJournalGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                RefreshBrigadierNames();
                RefreshSpecialties();
                RefreshProfessions();
                NormalizeOtRows();
                SortOtJournal();
                RequestReminderRefresh();
                MarkTimesheetOtSyncDirty();
                RequestTimesheetRebuild();
                SaveState();
            }));
        }

        private void OtJournalGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete && OtJournalGrid.SelectedItem is OtJournalEntry row)
            {
                currentObject?.OtJournal?.Remove(row);
                RefreshBrigadierNames();
                RefreshSpecialties();
                RefreshProfessions();
                RequestReminderRefresh();
                MarkTimesheetOtSyncDirty();
                RequestTimesheetRebuild();
                SaveState();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                OtJournalGrid.CommitEdit(DataGridEditingUnit.Cell, true);
                OtJournalGrid.CommitEdit(DataGridEditingUnit.Row, true);

                if (OtJournalGrid.CurrentCell.Column != null)
                {
                    var col = OtJournalGrid.CurrentCell.Column.DisplayIndex;
                    var rowIndex = OtJournalGrid.Items.IndexOf(OtJournalGrid.CurrentItem);

                    if (col < OtJournalGrid.Columns.Count - 1)
                        col++;
                    else
                    {
                        col = 0;
                        rowIndex = Math.Min(rowIndex + 1, OtJournalGrid.Items.Count - 1);
                    }

                    if (rowIndex >= 0 && rowIndex < OtJournalGrid.Items.Count)
                    {
                        OtJournalGrid.SelectedItem = OtJournalGrid.Items[rowIndex];
                        OtJournalGrid.CurrentCell = new DataGridCellInfo(OtJournalGrid.Items[rowIndex], OtJournalGrid.Columns[col]);
                        OtJournalGrid.BeginEdit();
                    }
                }
            }
        }

        private void InitializeTimesheet()
        {
            EnsureTimesheetStorage();
            MarkTimesheetOtSyncDirty();
            RequestTimesheetRebuild();
        }

        private void EnsureTimesheetStorage()
        {
            if (currentObject == null)
                return;

            currentObject.TimesheetPeople ??= new List<TimesheetPersonEntry>();
            currentObject.TimesheetPeople.RemoveAll(x => x == null);

            foreach (var person in currentObject.TimesheetPeople)
            {
                if (person.PersonId == Guid.Empty)
                    person.PersonId = Guid.NewGuid();
            }

            NormalizeTimesheetMonthsWindow();
            RefreshTimesheetPersonSubscriptions();
        }

        private static DateTime GetCurrentMonthDate()
            => new(DateTime.Today.Year, DateTime.Today.Month, 1);

        private static string ToMonthKey(DateTime month)
            => month.ToString("yyyy-MM", CultureInfo.InvariantCulture);

        private static List<DateTime> GetTimesheetWindowMonths()
        {
            var current = GetCurrentMonthDate();
            return new List<DateTime>
            {
                current.AddMonths(-2),
                current.AddMonths(-1),
                current,
                current.AddMonths(1)
            };
        }

        private void NormalizeTimesheetMonthsWindow()
        {
            if (currentObject?.TimesheetPeople == null)
                return;

            var allowedMonths = GetTimesheetWindowMonths();
            var minMonth = allowedMonths.First();
            var maxMonth = allowedMonths.Last();
            var allowedKeys = allowedMonths.Select(ToMonthKey).ToHashSet(StringComparer.Ordinal);
            var forwardKey = ToMonthKey(GetCurrentMonthDate().AddMonths(1));

            if (timesheetMonth < minMonth)
                timesheetMonth = minMonth;
            else if (timesheetMonth > maxMonth)
                timesheetMonth = maxMonth;

            foreach (var person in currentObject.TimesheetPeople.Where(x => x != null))
            {
                person.Months ??= new List<TimesheetMonthEntry>();
                person.Months.RemoveAll(m => m == null || string.IsNullOrWhiteSpace(m.MonthKey) || !allowedKeys.Contains(m.MonthKey));

                var seenMonthKeys = new HashSet<string>(StringComparer.Ordinal);
                for (var i = person.Months.Count - 1; i >= 0; i--)
                {
                    var key = person.Months[i].MonthKey?.Trim();
                    person.Months[i].MonthKey = key;
                    if (string.IsNullOrWhiteSpace(key) || !seenMonthKeys.Add(key))
                        person.Months.RemoveAt(i);
                }

                foreach (var month in person.Months)
                {
                    month.DayEntries ??= new Dictionary<int, TimesheetDayEntry>();
                    month.DayValues ??= new Dictionary<int, string>();

                    foreach (var invalidDay in month.DayEntries.Keys.Where(d => d < 1 || d > 31).ToList())
                        month.DayEntries.Remove(invalidDay);
                    foreach (var invalidDay in month.DayValues.Keys.Where(d => d < 1 || d > 31).ToList())
                        month.DayValues.Remove(invalidDay);

                    foreach (var day in month.DayEntries.Keys.ToList())
                    {
                        month.DayEntries[day] ??= new TimesheetDayEntry();
                        month.DayValues[day] = month.DayEntries[day].Value ?? string.Empty;
                    }

                    foreach (var pair in month.DayValues.ToList())
                    {
                        if (month.DayEntries.ContainsKey(pair.Key))
                            continue;

                        if (string.IsNullOrWhiteSpace(pair.Value))
                            continue;

                        month.DayEntries[pair.Key] = new TimesheetDayEntry
                        {
                            Value = pair.Value.Trim()
                        };
                    }
                }

                var forwardMonth = person.Months.FirstOrDefault(m => string.Equals(m.MonthKey, forwardKey, StringComparison.Ordinal));
                if (forwardMonth == null)
                {
                    person.Months.Add(new TimesheetMonthEntry
                    {
                        MonthKey = forwardKey,
                        DayEntries = new Dictionary<int, TimesheetDayEntry>(),
                        DayValues = new Dictionary<int, string>()
                    });
                }
                else
                {
                    forwardMonth.DayEntries ??= new Dictionary<int, TimesheetDayEntry>();
                    forwardMonth.DayValues ??= new Dictionary<int, string>();
                    forwardMonth.DayEntries.Clear();
                    forwardMonth.DayValues.Clear();
                }
            }
        }

        private void RefreshTimesheetPersonSubscriptions()
        {
            foreach (var person in subscribedTimesheetPeople)
                person.PropertyChanged -= TimesheetPersonEntry_PropertyChanged;

            subscribedTimesheetPeople.Clear();

            if (currentObject?.TimesheetPeople == null)
                return;

            foreach (var person in currentObject.TimesheetPeople.Distinct())
            {
                if (person == null)
                    continue;

                person.PropertyChanged += TimesheetPersonEntry_PropertyChanged;
                subscribedTimesheetPeople.Add(person);
            }
        }

        private void TimesheetPersonEntry_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(TimesheetPersonEntry.Months))
                return;

            Dispatcher.BeginInvoke(new Action(UpdateTimesheetMissingDocsNotification));
        }

        private void SyncTimesheetPeopleWithOtJournal()
        {
            if (currentObject?.OtJournal == null)
                return;

            EnsureTimesheetStorage();
            var timesheetPeople = currentObject.TimesheetPeople.Where(x => x != null).ToList();
            var timesheetById = timesheetPeople.ToDictionary(x => x.PersonId, x => x);
            var timesheetByName = timesheetPeople
                .Where(x => !string.IsNullOrWhiteSpace(x.FullName))
                .GroupBy(x => NormalizePersonNameKey(x.FullName))
                .ToDictionary(x => x.Key, x => x.First(), StringComparer.CurrentCultureIgnoreCase);

            var otGroups = currentObject.OtJournal
                .Where(x => x != null && !x.IsDismissed && !string.IsNullOrWhiteSpace(x.FullName))
                .GroupBy(x => NormalizePersonNameKey(x.FullName))
                .ToList();

            foreach (var group in otGroups)
            {
                var latest = group.OrderByDescending(x => x.InstructionDate).First();

                var resolvedId = group
                    .Select(x => x.PersonId)
                    .FirstOrDefault(x => x != Guid.Empty);

                if (resolvedId == Guid.Empty && timesheetByName.TryGetValue(group.Key, out var byName))
                    resolvedId = byName.PersonId;

                if (resolvedId == Guid.Empty)
                    resolvedId = Guid.NewGuid();

                foreach (var row in group.Where(x => x.PersonId != resolvedId))
                    row.PersonId = resolvedId;

                if (timesheetById.TryGetValue(resolvedId, out var existing))
                {
                    existing.FullName = latest.FullName?.Trim();
                    existing.Specialty = latest.Specialty;
                    existing.Rank = latest.Rank;
                    existing.IsBrigadier = latest.IsBrigadier;
                    existing.BrigadeName = latest.IsBrigadier ? latest.FullName?.Trim() : latest.BrigadierName;
                    continue;
                }

                var created = new TimesheetPersonEntry
                {
                    PersonId = resolvedId,
                    FullName = latest.FullName?.Trim(),
                    Specialty = latest.Specialty,
                    Rank = latest.Rank,
                    IsBrigadier = latest.IsBrigadier,
                    BrigadeName = latest.IsBrigadier ? latest.FullName?.Trim() : latest.BrigadierName
                };
                currentObject.TimesheetPeople.Add(created);
                timesheetById[resolvedId] = created;
                timesheetByName[group.Key] = created;
            }
        }

        private static string NormalizePersonNameKey(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
                return string.Empty;

            return Regex.Replace(fullName.Trim(), @"\s+", " ").ToUpperInvariant();
        }

        private void RebuildTimesheetView(bool force = false)
        {
            if (currentObject == null)
            {
                RefreshTimesheetPersonSubscriptions();
                TimesheetGrid.ItemsSource = null;
                timesheetNeedsRebuild = false;
                timesheetOtSyncDirty = true;
                RequestReminderRefresh();
                return;
            }

            EnsureTimesheetStorage();

            if (!force && !ReferenceEquals(MainTabs?.SelectedItem, TimesheetTab))
            {
                timesheetNeedsRebuild = true;
                return;
            }

            if (timesheetOtSyncDirty)
            {
                SyncTimesheetPeopleWithOtJournal();
                timesheetOtSyncDirty = false;
            }

            RefreshTimesheetPersonSubscriptions();
            RefreshTimesheetBrigades();

            BuildTimesheetColumns();
            RefreshTimesheetRows();
            TimesheetMonthText.Text = timesheetMonth.ToString("MMMM yyyy", CultureInfo.CurrentCulture);
            timesheetNeedsRebuild = false;
            UpdateTimesheetMissingDocsNotification();
        }

        private void RefreshTimesheetBrigades()
        {
            timesheetBrigades.Clear();
            timesheetBrigades.Add("Все бригады");
            timesheetAssignableBrigades.Clear();

            if (currentObject?.TimesheetPeople == null)
                return;

            foreach (var brigade in currentObject.TimesheetPeople
                         .Where(x => x != null)
                         .Select(x => NormalizeBrigadeName(x))
                         .Distinct(StringComparer.CurrentCultureIgnoreCase)
                         .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
            {
                timesheetBrigades.Add(brigade);
                if (!string.Equals(brigade, "Без бригады", StringComparison.CurrentCultureIgnoreCase))
                    timesheetAssignableBrigades.Add(brigade);
            }

            TimesheetBrigadeFilter.ItemsSource = timesheetBrigades;
            if (!timesheetBrigades.Contains(selectedTimesheetBrigade))
                selectedTimesheetBrigade = "Все бригады";
            TimesheetBrigadeFilter.SelectedItem = selectedTimesheetBrigade;
        }

        private string NormalizeBrigadeName(TimesheetPersonEntry row)
        {
            if (row == null)
                return "Без бригады";

            if (row.IsBrigadier)
                return string.IsNullOrWhiteSpace(row.FullName) ? "Без бригады" : row.FullName.Trim();

            return string.IsNullOrWhiteSpace(row.BrigadeName) ? "Без бригады" : row.BrigadeName.Trim();
        }

        private void RefreshTimesheetRows()
        {
            timesheetRows.Clear();
            if (currentObject?.TimesheetPeople == null)
                return;

            var monthKey = timesheetMonth.ToString("yyyy-MM");
            var filtered = currentObject.TimesheetPeople
                .Where(x => x != null)
                .Where(x => selectedTimesheetBrigade == "Все бригады"
                            || string.Equals(NormalizeBrigadeName(x), selectedTimesheetBrigade, StringComparison.CurrentCultureIgnoreCase))
                .OrderBy(x => NormalizeBrigadeName(x), StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.IsBrigadier ? 0 : 1)
                .ThenBy(x => x.FullName ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var grouped = filtered.GroupBy(x => NormalizeBrigadeName(x)).ToList();
            var number = 1;
            foreach (var group in grouped)
            {
                var rows = group.ToList();
                for (var i = 0; i < rows.Count; i++)
                {
                    var vm = new TimesheetRowViewModel(rows[i], monthKey)
                    {
                        Number = number++,
                        IsCrewStart = i == 0,
                        IsCrewEnd = i == rows.Count - 1
                    };
                    vm.RecalculateTotal();
                    timesheetRows.Add(vm);
                }
            }

            TimesheetGrid.ItemsSource = timesheetRows;
        }

        private void BuildTimesheetColumns()
        {
            if (TimesheetGrid == null)
                return;

            TimesheetGrid.Columns.Clear();
            TimesheetGrid.Columns.Add(new DataGridTextColumn { Header = "№", Binding = new Binding(nameof(TimesheetRowViewModel.Number)), IsReadOnly = true, Width = 45 });
            TimesheetGrid.Columns.Add(new DataGridTextColumn { Header = "ФИО", Binding = new Binding(nameof(TimesheetRowViewModel.FullName)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 220 });
            TimesheetGrid.Columns.Add(new DataGridTextColumn { Header = "Часов/день", Binding = new Binding(nameof(TimesheetRowViewModel.DailyWorkHours)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 90 });
            TimesheetGrid.Columns.Add(new DataGridTextColumn { Header = "Специальность", Binding = new Binding(nameof(TimesheetRowViewModel.Specialty)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 170 });
            TimesheetGrid.Columns.Add(new DataGridTextColumn { Header = "Разряд", Binding = new Binding(nameof(TimesheetRowViewModel.Rank)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 70 });
            TimesheetGrid.Columns.Add(new DataGridCheckBoxColumn
            {
                Header = "Бригадир",
                Binding = new Binding(nameof(TimesheetRowViewModel.IsBrigadier)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                Width = 85
            });
            var daysInMonth = DateTime.DaysInMonth(timesheetMonth.Year, timesheetMonth.Month);
            for (var day = 1; day <= daysInMonth; day++)
            {
                if (IsTodayColumnSlot(day))
                    TimesheetGrid.Columns.Add(BuildTodayPresenceColumn(day));

                var date = new DateTime(timesheetMonth.Year, timesheetMonth.Month, day);
                var isWeekend = date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday;
                var column = new DataGridTemplateColumn
                {
                    Header = day.ToString(),
                    Width = 46,
                    CellTemplate = BuildTimesheetDayCellTemplate(day),
                    CellEditingTemplate = BuildTimesheetDayCellTemplate(day)
                };
                var style = new Style(typeof(DataGridCell));
                if (isWeekend)
                    style.Setters.Add(new Setter(DataGridCell.BackgroundProperty, new SolidColorBrush(Color.FromRgb(245, 245, 245))));
                style.Setters.Add(new Setter(DataGridCell.ToolTipProperty, new Binding { Converter = new TimesheetDayCommentToolTipConverter(day) }));
                column.CellStyle = style;
                TimesheetGrid.Columns.Add(column);
            }

            TimesheetGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Итого часов",
                Binding = new Binding(nameof(TimesheetRowViewModel.MonthTotalHours)) { StringFormat = "0.##" },
                IsReadOnly = true,
                Width = 95
            });
        }

        private bool IsTodayColumnSlot(int day)
    => timesheetMonth.Year == DateTime.Today.Year
       && timesheetMonth.Month == DateTime.Today.Month
       && day == DateTime.Today.Day;

        private DataGridTemplateColumn BuildTodayPresenceColumn(int day)
        {
            var checkBoxFactory = new FrameworkElementFactory(typeof(CheckBox));
            checkBoxFactory.SetValue(CheckBox.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            checkBoxFactory.SetValue(CheckBox.VerticalAlignmentProperty, VerticalAlignment.Center);
            checkBoxFactory.SetValue(CheckBox.ForegroundProperty, new SolidColorBrush(Color.FromRgb(146, 64, 14)));
            checkBoxFactory.SetValue(FrameworkElement.TagProperty, day);
            checkBoxFactory.SetBinding(ToggleButton.IsCheckedProperty, new Binding($"PresenceChecked[{day}]")
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });
            checkBoxFactory.AddHandler(ToggleButton.CheckedEvent, new RoutedEventHandler(TodayPresenceChanged));
            checkBoxFactory.AddHandler(ToggleButton.UncheckedEvent, new RoutedEventHandler(TodayPresenceChanged));

            return new DataGridTemplateColumn
            {
                Header = "Сегодня",
                Width = 64,
                CellTemplate = new DataTemplate { VisualTree = checkBoxFactory },
                CellEditingTemplate = new DataTemplate { VisualTree = checkBoxFactory },
                CellStyle = new Style(typeof(DataGridCell))
                {
                    Setters =
                    {
                        new Setter(DataGridCell.BackgroundProperty, new SolidColorBrush(Color.FromRgb(254, 249, 195)))
                    }
                }
            };
        }

        private void TodayPresenceChanged(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element
                && element.DataContext is TimesheetRowViewModel row
                && int.TryParse(element.Tag?.ToString(), out var day)
                && day > 0)
            {
                var shouldMarkPresent = row.GetPresenceChecked(day);
                var autoValue = shouldMarkPresent
                    ? Math.Clamp(row.DailyWorkHours, 1, 24).ToString(CultureInfo.CurrentCulture)
                    : "Н";

                row[day] = autoValue;
            }

            SaveState();
            UpdateTimesheetMissingDocsNotification();
            TimesheetGrid?.Items.Refresh();
        }

        private DataTemplate BuildTimesheetDayCellTemplate(int day)
        {
            var grid = new FrameworkElementFactory(typeof(Grid));

            var editor = new FrameworkElementFactory(typeof(TextBox));
            editor.SetValue(TextBox.BorderThicknessProperty, new Thickness(0));
            editor.SetValue(TextBox.PaddingProperty, new Thickness(2, 0, 10, 0));
            editor.SetBinding(TextBox.TextProperty, new Binding($"[{day}]")
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });
            editor.SetBinding(TextBox.BackgroundProperty, new Binding { Converter = new TimesheetDayBackgroundConverter(day) });
            grid.AppendChild(editor);

            var marker = new FrameworkElementFactory(typeof(TextBlock));
            marker.SetValue(TextBlock.TextProperty, "*");
            marker.SetValue(TextBlock.ForegroundProperty, Brushes.DarkOrange);
            marker.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold);
            marker.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Right);
            marker.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Top);
            marker.SetValue(TextBlock.MarginProperty, new Thickness(0, -1, 1, 0));
            marker.SetBinding(TextBlock.VisibilityProperty, new Binding { Converter = new TimesheetDayCommentVisibilityConverter(day) });
            marker.SetBinding(TextBlock.ToolTipProperty, new Binding { Converter = new TimesheetDayCommentToolTipConverter(day) });
            grid.AppendChild(marker);

            return new DataTemplate { VisualTree = grid };
        }

        private sealed class TimesheetDayBackgroundConverter : IValueConverter
        {
            private readonly int day;
            public TimesheetDayBackgroundConverter(int day) => this.day = day;
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
                               => value is TimesheetRowViewModel row ? row.GetDayBackground(day) : Brushes.Transparent;
            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
                               => Binding.DoNothing;
        }

        private sealed class TimesheetDayCommentVisibilityConverter : IValueConverter
        {
            private readonly int day;
            public TimesheetDayCommentVisibilityConverter(int day) => this.day = day;
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
                => value is TimesheetRowViewModel row && row.HasDayComment(day) ? Visibility.Visible : Visibility.Collapsed;
            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
                => Binding.DoNothing;
        }

        private sealed class TimesheetDayCommentToolTipConverter : IValueConverter
        {
            private readonly int day;
            public TimesheetDayCommentToolTipConverter(int day) => this.day = day;
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (value is not TimesheetRowViewModel row)
                    return string.Empty;

                var comment = row.GetDayComment(day);
                if (string.IsNullOrWhiteSpace(comment))
                    return null;

                return $"Комментарий: {comment}";
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
                => Binding.DoNothing;
        }

        private void TimesheetGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (e.Row.Item is not TimesheetRowViewModel)
                return;
        }

        private void TimesheetGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Row.Item is not TimesheetRowViewModel row)
                return;

            var day = ParseDayFromHeader(e.Column?.Header);
            if (day <= 0)
            {
                SyncTimesheetPersonToOt(row.Source);
                var headerText = e.Column?.Header?.ToString() ?? string.Empty;
                if (string.Equals(headerText, "ФИО", StringComparison.CurrentCultureIgnoreCase)
                    || string.Equals(headerText, "Бригадир", StringComparison.CurrentCultureIgnoreCase))
                {
                    RefreshTimesheetBrigades();
                }
                row.RecalculateTotal();
                TimesheetGrid?.Items.Refresh();
                SaveState();
                UpdateTimesheetMissingDocsNotification();
                return;
            }

            if (e.EditingElement is not TextBox tb)
                return;

            var value = tb.Text?.Trim() ?? string.Empty;
            if (!row.Source.TryApplyDayValue(timesheetMonth.ToString("yyyy-MM"), day, value, out var validationError))
            {
                MessageBox.Show(validationError, "Проверка табеля", MessageBoxButton.OK, MessageBoxImage.Warning);
                e.Cancel = true;
                Dispatcher.BeginInvoke(new Action(RefreshTimesheetRows));
                return;
            }

            row.RecalculateTotal();
            SaveState();
            UpdateTimesheetMissingDocsNotification();

            if (day == DateTime.Today.Day
                && timesheetMonth.Year == DateTime.Today.Year
                && timesheetMonth.Month == DateTime.Today.Month
                && !string.IsNullOrWhiteSpace(value)
                && currentObject?.OtJournal != null)
            {
                var due = currentObject.OtJournal.Any(x =>
                    x.PersonId == row.PersonId &&
                    x.IsPendingRepeat &&
                    !x.IsDismissed);

                if (due)
                    MessageBox.Show($"⚠ {row.FullName}: требуется срочный повторный инструктаж по ОТ.", "Напоминание по ОТ", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private static int ParseDayFromHeader(object header)
            => int.TryParse(header?.ToString(), out var day) ? day : -1;
        private void TimesheetGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit)
                return;

            if (e.Row.Item is not TimesheetRowViewModel row)
                return;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                SyncTimesheetPersonToOt(row.Source);
            }), DispatcherPriority.Background);
        }


        private void TimesheetGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (e.Row.Item is not TimesheetRowViewModel row)
                return;

            e.Row.FontWeight = row.IsBrigadier ? FontWeights.Bold : FontWeights.Normal;
            e.Row.BorderBrush = Brushes.Black;
            e.Row.BorderThickness = new Thickness(0, row.IsCrewStart ? 2 : 0, 0, row.IsCrewEnd ? 2 : 0);
        }

        private void TimesheetGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Cancel = true;
        }
        private void TimesheetGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            var cell = TimesheetGrid.SelectedCells.FirstOrDefault();
            selectedTimesheetRow = cell.Item as TimesheetRowViewModel;
            selectedTimesheetDay = ParseDayFromHeader(cell.Column?.Header);
        }

        private void TimesheetGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete && Keyboard.Modifiers == ModifierKeys.Control)
            {
                DeleteSelectedTimesheetPerson();
                e.Handled = true;
            }
        }

        private void DeleteTimesheetPerson_Click(object sender, RoutedEventArgs e)
            => DeleteSelectedTimesheetPerson();

        private void DeleteSelectedTimesheetPerson()
        {
            if (selectedTimesheetRow == null)
            {
                MessageBox.Show("Выберите строку сотрудника в табеле.");
                return;
            }

            DeleteTimesheetPerson(selectedTimesheetRow);
        }

        private void DeleteTimesheetPerson(TimesheetRowViewModel row)
        {
            if (row == null || currentObject?.TimesheetPeople == null)
                return;

            var name = row.FullName;
            if (MessageBox.Show($"Удалить сотрудника \"{name}\" из табеля?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            if (currentObject?.OtJournal != null)
            {
                foreach (var otRow in currentObject.OtJournal.Where(x =>
                             (row.Source.PersonId != Guid.Empty && x.PersonId == row.Source.PersonId)
                             || (!string.IsNullOrWhiteSpace(x.FullName) && string.Equals(x.FullName.Trim(), row.FullName?.Trim(), StringComparison.CurrentCultureIgnoreCase))))
                {
                    otRow.IsDismissed = true;
                    otRow.IsPendingRepeat = false;
                    otRow.IsScheduledRepeat = false;
                }
            }

            currentObject.TimesheetPeople.Remove(row.Source);
            SaveState();
            RebuildTimesheetView();
        }

        private void EditSelectedDayComment_Click(object sender, RoutedEventArgs e)
        {
            if (!TryGetSelectedDayEntry(out var row, out var day))
                return;

            var currentComment = row.GetDayComment(day);
            var comment = PromptMultiline("Комментарий к дню", $"Введите комментарий ({day} число):", currentComment);
            if (comment == null)
                return;

            row.SetComment(day, comment);
            SaveState();
            UpdateTimesheetMissingDocsNotification();
            TimesheetGrid.Items.Refresh();
        }

        private void ToggleSelectedDayDocument_Click(object sender, RoutedEventArgs e)
        {
            if (!TryGetSelectedDayEntry(out var row, out var day))
                return;

            if (!row.IsNonHourCode(day))
            {
                MessageBox.Show("Документ закрытия нужен только для кодов (не числовых часов).");
                return;
            }

            var current = row.IsDocumentAccepted(day);
            row.SetDocumentAccepted(day, current == true ? false : true);
            SaveState();
            UpdateTimesheetMissingDocsNotification();
            TimesheetGrid.Items.Refresh();
        }

        private bool TryGetSelectedDayEntry(out TimesheetRowViewModel row, out int day)
        {
            row = selectedTimesheetRow;
            day = selectedTimesheetDay;
            if (row == null || day <= 0)
            {
                MessageBox.Show("Сначала выберите ячейку конкретного дня.");
                return false;
            }

            return true;
        }

        private void TimesheetPrevMonth_Click(object sender, RoutedEventArgs e)
        {
            var minMonth = GetCurrentMonthDate().AddMonths(-2);
            if (timesheetMonth <= minMonth)
            {
                timesheetMonth = minMonth;
                RebuildTimesheetView();
                return;
            }

            timesheetMonth = timesheetMonth.AddMonths(-1);
            RebuildTimesheetView();
        }

        private void TimesheetNextMonth_Click(object sender, RoutedEventArgs e)
        {
            var maxMonth = GetCurrentMonthDate().AddMonths(1);
            if (timesheetMonth >= maxMonth)
            {
                timesheetMonth = maxMonth;
                RebuildTimesheetView();
                return;
            }

            timesheetMonth = timesheetMonth.AddMonths(1);
            RebuildTimesheetView();
        }

        private void TimesheetCurrentMonth_Click(object sender, RoutedEventArgs e)
        {
            timesheetMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            RebuildTimesheetView();
        }

        private void TimesheetBrigadeFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedTimesheetBrigade = TimesheetBrigadeFilter.SelectedItem?.ToString() ?? "Все бригады";
            RefreshTimesheetRows();
        }

        private void AddTimesheetPerson_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            var fullName = TimesheetFullNameTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(fullName))
            {
                MessageBox.Show("Введите ФИО.");
                return;
            }

            EnsureTimesheetStorage();
            var workModeText = (TimesheetWorkModeComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty;
            var plannedHours = workModeText.IndexOf("12", StringComparison.CurrentCultureIgnoreCase) >= 0 ? 12 : 8;
            var person = new TimesheetPersonEntry
            {
                PersonId = Guid.NewGuid(),
                FullName = fullName,
                Specialty = TimesheetSpecialtyTextBox.Text?.Trim(),
                Rank = TimesheetRankTextBox.Text?.Trim(),
                IsBrigadier = TimesheetIsBrigadierCheckBox.IsChecked == true,
                BrigadeName = TimesheetBrigadeComboBox.Text?.Trim(),
                DailyWorkHours = plannedHours
            };

            if (person.IsBrigadier && string.IsNullOrWhiteSpace(person.BrigadeName))
                person.BrigadeName = person.FullName;

            currentObject.TimesheetPeople.Add(person);
            EnsurePersonInOtJournal(person);

            TimesheetFullNameTextBox.Text = string.Empty;
            TimesheetSpecialtyTextBox.Text = string.Empty;
            TimesheetRankTextBox.Text = string.Empty;
            TimesheetBrigadeComboBox.Text = string.Empty;
            TimesheetIsBrigadierCheckBox.IsChecked = false;
            TimesheetWorkModeComboBox.SelectedIndex = 0;

            RefreshBrigadierNames();
            RebuildTimesheetView();
            SortOtJournal();
            SaveState();
        }

        private void EnsurePersonInOtJournal(TimesheetPersonEntry person)
        {
            if (currentObject?.OtJournal == null || person == null || string.IsNullOrWhiteSpace(person.FullName))
                return;

            var matchingRows = currentObject.OtJournal
                .Where(x =>
                    (person.PersonId != Guid.Empty && x.PersonId == person.PersonId)
                    || (!string.IsNullOrWhiteSpace(x.FullName) && string.Equals(x.FullName.Trim(), person.FullName?.Trim(), StringComparison.CurrentCultureIgnoreCase)))
                .ToList();

            if (matchingRows.Any(x => !x.IsDismissed))
            {
                UpdateRepeatRequirementByTimesheet(person);
                return;
            }

            var hadInstructionBefore = matchingRows.Any(x => x.IsPrimaryInstruction || IsRepeatInstruction(x) || x.IsRepeatCompleted);
            var requireRepeat = hadInstructionBefore && (matchingRows.Any(x => x.IsDismissed) || HasLongAbsenceInTimesheet(person, 21));

            var ot = new OtJournalEntry
            {
                PersonId = person.PersonId,
                InstructionDate = DateTime.Today,
                FullName = person.FullName,
                Specialty = person.Specialty,
                Rank = person.Rank,
                Profession = person.Specialty,
                InstructionType = requireRepeat
                    ? BuildRepeatInstructionType(GetNextRepeatIndexForPerson(person.FullName))
                    : "Первичный на рабочем месте",
                RepeatPeriodMonths = 3,
                IsBrigadier = person.IsBrigadier,
                BrigadierName = person.IsBrigadier ? null : person.BrigadeName,
                IsPendingRepeat = requireRepeat,
                IsScheduledRepeat = false,
                IsRepeatCompleted = false
            };
            FillInstructionNumbersFromTemplate(ot);
            ot.PropertyChanged += OtJournalEntry_PropertyChanged;
            currentObject.OtJournal.Add(ot);
            NormalizeOtRows();
            SortOtJournal();
            RequestReminderRefresh();
        }

        private void SyncTimesheetPersonToOt(TimesheetPersonEntry person)
        {
            if (person == null || currentObject?.OtJournal == null)
                return;

            var matches = currentObject.OtJournal.Where(x =>
                (person.PersonId != Guid.Empty && x.PersonId == person.PersonId)
                || (!string.IsNullOrWhiteSpace(x.FullName) && string.Equals(x.FullName.Trim(), person.FullName?.Trim(), StringComparison.CurrentCultureIgnoreCase)))
                .ToList();

            isSyncingTimesheetToOt = true;
            try
            {
                foreach (var ot in matches)
                {
                    ot.FullName = person.FullName;
                    ot.Specialty = person.Specialty;
                    ot.Rank = person.Rank;
                    ot.Profession = string.IsNullOrWhiteSpace(ot.Profession) ? person.Specialty : ot.Profession;
                    ot.IsBrigadier = person.IsBrigadier;
                    ot.BrigadierName = person.IsBrigadier ? null : person.BrigadeName;
                }
            }
            finally
            {
                isSyncingTimesheetToOt = false;
            }

            if (matches.Count > 0)
            {
                SortOtJournal();
                RequestReminderRefresh();
            }

            UpdateRepeatRequirementByTimesheet(person);
        }

        private void UpdateRepeatRequirementByTimesheet(TimesheetPersonEntry person)
        {
            if (person == null || currentObject?.OtJournal == null || string.IsNullOrWhiteSpace(person.FullName))
                return;

            if (!HasLongAbsenceInTimesheet(person, 21))
                return;

            var personRows = currentObject.OtJournal
                .Where(x =>
                    !x.IsDismissed
                    && ((person.PersonId != Guid.Empty && x.PersonId == person.PersonId)
                        || string.Equals(x.FullName?.Trim(), person.FullName.Trim(), StringComparison.CurrentCultureIgnoreCase)))
                .OrderByDescending(x => x.InstructionDate)
                .ToList();

            if (personRows.Count == 0 || personRows.Any(x => x.IsPendingRepeat))
                return;

            var lastInstruction = personRows
                .Where(x => x.IsPrimaryInstruction || IsRepeatInstruction(x) || x.IsRepeatCompleted || x.IsScheduledRepeat)
                .FirstOrDefault();
            if (lastInstruction == null)
                return;

            var repeatRow = new OtJournalEntry
            {
                PersonId = person.PersonId,
                InstructionDate = DateTime.Today,
                FullName = lastInstruction.FullName,
                Specialty = lastInstruction.Specialty,
                Rank = lastInstruction.Rank,
                Profession = string.IsNullOrWhiteSpace(lastInstruction.Profession) ? lastInstruction.Specialty : lastInstruction.Profession,
                InstructionType = BuildRepeatInstructionType(GetNextRepeatIndexForPerson(lastInstruction.FullName)),
                InstructionNumbers = lastInstruction.InstructionNumbers,
                RepeatPeriodMonths = Math.Max(1, lastInstruction.RepeatPeriodMonths),
                IsBrigadier = lastInstruction.IsBrigadier,
                BrigadierName = lastInstruction.BrigadierName,
                IsPendingRepeat = true,
                IsScheduledRepeat = false,
                IsRepeatCompleted = false,
                IsDismissed = false
            };
            FillInstructionNumbersFromTemplate(repeatRow);
            repeatRow.PropertyChanged += OtJournalEntry_PropertyChanged;
            currentObject.OtJournal.Add(repeatRow);
            NormalizeOtRows();
            SortOtJournal();
            RequestReminderRefresh();
        }

        private static DateTime? GetLastWorkedDate(TimesheetPersonEntry person)
        {
            if (person?.Months == null || person.Months.Count == 0)
                return null;

            DateTime? result = null;
            foreach (var monthEntry in person.Months)
            {
                if (monthEntry == null || string.IsNullOrWhiteSpace(monthEntry.MonthKey))
                    continue;

                if (!DateTime.TryParseExact(monthEntry.MonthKey + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var monthStart))
                    continue;

                foreach (var dayEntry in monthEntry.DayEntries)
                {
                    var day = dayEntry.Key;
                    if (day < 1 || day > DateTime.DaysInMonth(monthStart.Year, monthStart.Month))
                        continue;

                    var value = dayEntry.Value?.Value?.Trim();
                    var isWorkedDay = false;
                    if (!string.IsNullOrWhiteSpace(value)
                        && (double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out var hours)
                            || double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out hours)))
                    {
                        isWorkedDay = hours > 0.0001;
                    }

                    if (!isWorkedDay && string.Equals(dayEntry.Value?.PresenceMark, "✔", StringComparison.Ordinal))
                        isWorkedDay = true;

                    if (!isWorkedDay)
                        continue;

                    var date = new DateTime(monthStart.Year, monthStart.Month, day);
                    if (!result.HasValue || date > result.Value)
                        result = date;
                }
            }

            return result;
        }

        private static bool HasLongAbsenceInTimesheet(TimesheetPersonEntry person, int daysThreshold)
        {
            if (person == null)
                return false;

            var lastWorked = GetLastWorkedDate(person);
            if (!lastWorked.HasValue)
                return false;

            return (DateTime.Today.Date - lastWorked.Value.Date).TotalDays > Math.Max(0, daysThreshold);
        }

        private void TimesheetPrintFilled_Click(object sender, RoutedEventArgs e) => PrintTimesheet(blank: false);
        private void TimesheetPrintBlank_Click(object sender, RoutedEventArgs e) => PrintTimesheet(blank: true);

        private void PrintTimesheet(bool blank)
        {
            if (timesheetRows.Count == 0)
            {
                MessageBox.Show("Нет строк для печати.");
                return;
            }

            var doc = BuildTimesheetDocument(blank);
            var pd = new PrintDialog();
            if (pd.ShowDialog() != true)
                return;

            doc.PageHeight = pd.PrintableAreaHeight;
            doc.PageWidth = pd.PrintableAreaWidth;
            doc.ColumnWidth = pd.PrintableAreaWidth;
            pd.PrintDocument(((IDocumentPaginatorSource)doc).DocumentPaginator, "Табель");
        }
        private void UpdateTimesheetMissingDocsNotification()
        {
            RequestReminderRefresh();
        }

        private string PromptMultiline(string title, string question, string initialValue)
        {
            var dialog = new Window
            {
                Title = title,
                Owner = this,
                Width = 470,
                Height = 260,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize
            };

            var root = new Grid { Margin = new Thickness(12) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var q = new TextBlock { Text = question, Margin = new Thickness(0, 0, 0, 8) };
            Grid.SetRow(q, 0);
            root.Children.Add(q);

            var tb = new TextBox
            {
                Text = initialValue ?? string.Empty,
                AcceptsReturn = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(tb, 1);
            root.Children.Add(tb);

            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            var ok = new Button { Content = "Сохранить", MinWidth = 90, Margin = new Thickness(0, 0, 6, 0), IsDefault = true };
            var cancel = new Button { Content = "Отмена", MinWidth = 90, IsCancel = true };
            ok.Click += (_, _) => dialog.DialogResult = true;
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            Grid.SetRow(buttons, 2);
            root.Children.Add(buttons);

            dialog.Content = root;
            return dialog.ShowDialog() == true ? tb.Text?.Trim() ?? string.Empty : null;
        }

        private FlowDocument BuildTimesheetDocument(bool blank)
        {
            var doc = new FlowDocument { FontFamily = new FontFamily("Segoe UI"), FontSize = 10 };
            doc.Blocks.Add(new Paragraph(new Run($"Табель за {timesheetMonth:MMMM yyyy}")) { FontSize = 14, FontWeight = FontWeights.Bold });
            doc.Blocks.Add(new Paragraph(new Run($"Фильтр: {selectedTimesheetBrigade}")));

            var table = new Table();
            doc.Blocks.Add(table);
            table.Columns.Add(new TableColumn { Width = new GridLength(25) });
            table.Columns.Add(new TableColumn { Width = new GridLength(130) });
            table.Columns.Add(new TableColumn { Width = new GridLength(90) });
            table.Columns.Add(new TableColumn { Width = new GridLength(40) });
            var daysInMonth = DateTime.DaysInMonth(timesheetMonth.Year, timesheetMonth.Month);
            for (var i = 1; i <= daysInMonth; i++)
                table.Columns.Add(new TableColumn { Width = new GridLength(22) });
            table.Columns.Add(new TableColumn { Width = new GridLength(55) });

            var group = new TableRowGroup();
            table.RowGroups.Add(group);
            var header = new TableRow { FontWeight = FontWeights.Bold, Background = Brushes.LightGray };
            group.Rows.Add(header);
            header.Cells.Add(new TableCell(new Paragraph(new Run("№"))));
            header.Cells.Add(new TableCell(new Paragraph(new Run("ФИО"))));
            header.Cells.Add(new TableCell(new Paragraph(new Run("Спец."))));
            header.Cells.Add(new TableCell(new Paragraph(new Run("Раз."))));
            for (var i = 1; i <= daysInMonth; i++)
                header.Cells.Add(new TableCell(new Paragraph(new Run(i.ToString()))));
            header.Cells.Add(new TableCell(new Paragraph(new Run("Итого"))));

            foreach (var row in timesheetRows)
            {
                var tr = new TableRow();
                if (row.IsBrigadier)
                    tr.FontWeight = FontWeights.Bold;
                group.Rows.Add(tr);

                tr.Cells.Add(new TableCell(new Paragraph(new Run(row.Number.ToString()))));
                tr.Cells.Add(new TableCell(new Paragraph(new Run(row.FullName))));
                tr.Cells.Add(new TableCell(new Paragraph(new Run(row.Specialty))));
                tr.Cells.Add(new TableCell(new Paragraph(new Run(row.Rank))));
                for (var i = 1; i <= daysInMonth; i++)
                    tr.Cells.Add(new TableCell(new Paragraph(new Run(blank ? string.Empty : row.GetDayValue(i)))));
                tr.Cells.Add(new TableCell(new Paragraph(new Run(blank ? string.Empty : row.MonthTotalHours.ToString("0.##")))));
            }

            return doc;
        }

        private void InitializeProductionJournal()
        {
            EnsureProductionJournalStorage();
            RefreshProductionJournalState();
        }

        private void EnsureProductionJournalStorage()
        {
            if (currentObject == null)
                return;

            currentObject.ProductionJournal ??= new List<ProductionJournalEntry>();
            currentObject.ProductionAutoFillSettings ??= new ProductionAutoFillSettings();
            currentObject.SummaryMarksByGroup ??= new Dictionary<string, List<string>>();
        }

        private void RefreshProductionJournalState()
        {
            EnsureProductionJournalStorage();
            NormalizeProductionJournalRows();
            ApplyProductionDisplayMerging();
            productionJournalRows.Clear();

            if (currentObject?.ProductionJournal != null)
            {
                foreach (var row in currentObject.ProductionJournal)
                    productionJournalRows.Add(row);
            }

            if (ProductionJournalGrid != null)
                ProductionJournalGrid.ItemsSource = productionJournalRows;

            RefreshProductionJournalLookups();
            RebuildMountedDemandFromProductionJournal();
            RefreshProductionRemainingInfo();
            if (ProductionJournalGrid != null)
                ProductionJournalGrid.Items.Refresh();
            if (selectedProductionRow != null && !productionJournalRows.Contains(selectedProductionRow))
                selectedProductionRow = null;
            UpdateProductionFormState();
        }

        private void ApplyProductionDisplayMerging()
        {
            if (currentObject?.ProductionJournal == null || currentObject.ProductionJournal.Count == 0)
                return;

            ProductionJournalEntry previous = null;
            foreach (var row in currentObject.ProductionJournal)
            {
                var sameAsPrevious = previous != null
                    && row.Date.Date == previous.Date.Date
                    && string.Equals((row.Weather ?? string.Empty).Trim(), (previous.Weather ?? string.Empty).Trim(), StringComparison.CurrentCultureIgnoreCase);

                row.SuppressDateDisplay = sameAsPrevious;
                row.SuppressWeatherDisplay = sameAsPrevious;
                previous = row;
            }
        }

        private void RefreshProductionJournalLookups()
        {
            productionActions.Clear();
            foreach (var value in new[] { "Монтаж", "Кладка", "Устройство" }
                .Concat(currentObject?.ProductionJournal?.Select(x => x.ActionName) ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase))
            {
                productionActions.Add(value);
            }
            FillLookupCollection(productionTargets,
                (currentObject?.ProductionJournal?.Select(x => x.WorkName) ?? Enumerable.Empty<string>())
                .Concat(currentObject?.MaterialCatalog?
                    .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.TypeName) ?? Enumerable.Empty<string>())
                .Concat(journal
                    .Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.MaterialGroup)));
            FillLookupCollection(productionElements,
                currentObject?.MaterialCatalog?
                    .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.MaterialName)
                    ?? Enumerable.Empty<string>());
            FillLookupCollection(productionWeatherOptions,
                currentObject?.ProductionJournal?.Select(x => x.Weather));
            FillLookupCollection(productionDeviationOptions,
                currentObject?.ProductionJournal?.Select(x => x.Deviations));

            productionBlockOptions.Clear();
            if (currentObject != null)
            {
                for (var i = 1; i <= currentObject.BlocksCount; i++)
                    productionBlockOptions.Add(i.ToString());
            }

            FillLookupCollection(productionMarkOptions,
                (currentObject?.SummaryMarksByGroup?.Values
                    .Where(x => x != null)
                    .SelectMany(x => x ?? Enumerable.Empty<string>())
                    ?? Enumerable.Empty<string>())
                .Concat(LevelMarkHelper.GetDefaultMarks(currentObject)));

            RefreshProductionElementOptions();
        }

        private static void FillLookupCollection(ObservableCollection<string> target, IEnumerable<string> values)
        {
            target.Clear();
            foreach (var value in (values ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x))
            {
                target.Add(value);
            }
        }

        private void RefreshProductionElementOptions()
        {
            var currentText = ProductionElementsBox?.Text?.Trim();
            var selectedWork = ProductionWorkBox?.Text?.Trim();
            var values = new List<string>();

            if (!string.IsNullOrWhiteSpace(selectedWork))
            {
                values.AddRange(currentObject?.MaterialCatalog?
                    .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals(x.TypeName ?? string.Empty, selectedWork, StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.MaterialName)
                    ?? Enumerable.Empty<string>());

                values.AddRange(journal
                    .Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals(x.MaterialGroup ?? string.Empty, selectedWork, StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.MaterialName));
            }
            else
            {
                values.AddRange(currentObject?.MaterialCatalog?
                    .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.MaterialName)
                    ?? Enumerable.Empty<string>());

                values.AddRange(journal
                    .Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.MaterialName));
            }

            values.AddRange(currentObject?.ProductionJournal?
                .SelectMany(x => ParseProductionItems(x.ElementsText))
                .Select(x => x.MaterialName)
                ?? Enumerable.Empty<string>());

            FillLookupCollection(productionElements, values);

            if (ProductionElementsBox != null && !string.IsNullOrWhiteSpace(currentText))
                ProductionElementsBox.Text = currentText;
        }

        private void ProductionWorkBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshProductionElementOptions();
        }

        private void ProductionWorkBox_LostFocus(object sender, RoutedEventArgs e)
        {
            RefreshProductionElementOptions();
        }

        private void InitializeInspectionJournal()
        {
            EnsureInspectionJournalStorage();
            RefreshInspectionJournalState();
        }

        private void EnsureInspectionJournalStorage()
        {
            if (currentObject == null)
                return;

            currentObject.InspectionJournal ??= new List<InspectionJournalEntry>();
        }

        private void RefreshInspectionJournalState()
        {
            EnsureInspectionJournalStorage();
            inspectionJournalRows.Clear();

            if (currentObject?.InspectionJournal != null)
            {
                foreach (var row in currentObject.InspectionJournal
                    .OrderBy(x => x.JournalName)
                    .ThenBy(x => x.InspectionName)
                    .ThenBy(x => x.IsCompletionHistory ? 0 : 1)
                    .ThenByDescending(x => x.LastCompletedDate ?? DateTime.MinValue)
                    .ThenBy(x => x.NextReminderDate))
                {
                    inspectionJournalRows.Add(row);
                }
            }

            if (InspectionJournalGrid != null)
                InspectionJournalGrid.ItemsSource = inspectionJournalRows;

            RefreshInspectionLookups();
            if (InspectionJournalGrid != null)
                InspectionJournalGrid.Items.Refresh();

            if (selectedInspectionRow != null && !inspectionJournalRows.Contains(selectedInspectionRow))
                selectedInspectionRow = null;

            UpdateInspectionFormState();
            RequestReminderRefresh();
        }

        private void RefreshInspectionLookups()
        {
            FillLookupCollection(inspectionJournalNames, currentObject?.InspectionJournal?.Select(x => x.JournalName));
            FillLookupCollection(inspectionNames, currentObject?.InspectionJournal?.Select(x => x.InspectionName));
        }

        private void SaveInspectionForm_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            EnsureInspectionJournalStorage();
            var row = selectedInspectionRow ?? new InspectionJournalEntry();
            inspectionRowSnapshotJson = JsonSerializer.Serialize(row);
            ReadInspectionForm(row);

            if (!ValidateInspectionRow(row, out var message))
            {
                RestoreInspectionSnapshot(row);
                MessageBox.Show(message, "Проверка осмотра", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (selectedInspectionRow == null)
                currentObject.InspectionJournal.Add(row);

            selectedInspectionRow = row;
            RefreshInspectionJournalState();
            InspectionJournalGrid.SelectedItem = row;
            InspectionJournalGrid.ScrollIntoView(row);
            SaveState();
        }

        private void DeleteInspectionRow_Click(object sender, RoutedEventArgs e)
        {
            var row = selectedInspectionRow ?? InspectionJournalGrid?.SelectedItem as InspectionJournalEntry;
            if (row == null || currentObject?.InspectionJournal == null)
            {
                MessageBox.Show("Выберите запись в осмотрах.");
                return;
            }

            currentObject.InspectionJournal.Remove(row);
            selectedInspectionRow = null;
            RefreshInspectionJournalState();
            SaveState();
        }

        private void ClearInspectionForm_Click(object sender, RoutedEventArgs e)
        {
            selectedInspectionRow = null;
            UpdateInspectionFormState(clearFields: true);
            if (InspectionJournalGrid != null)
                InspectionJournalGrid.SelectedItem = null;
        }

        private void InspectionJournalGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (InspectionJournalGrid?.SelectedItem is not InspectionJournalEntry row)
            {
                if (selectedInspectionRow != null)
                    UpdateInspectionFormState(clearFields: false);
                return;
            }

            selectedInspectionRow = row;
            FillInspectionForm(row);
            UpdateInspectionFormState();
        }

        private void UpdateInspectionFormState(bool clearFields = false)
        {
            if (clearFields)
                ResetInspectionFormInputs();
            else if (selectedInspectionRow == null && InspectionReminderStartDatePicker != null && InspectionReminderStartDatePicker.SelectedDate == null)
                ResetInspectionFormInputs();

            if (InspectionSaveButton != null)
                InspectionSaveButton.Content = selectedInspectionRow == null ? "➕ Добавить" : "💾 Обновить";
        }

        private void ResetInspectionFormInputs()
        {
            if (InspectionReminderStartDatePicker == null)
                return;

            InspectionJournalNameBox.Text = string.Empty;
            InspectionNameBox.Text = string.Empty;
            InspectionReminderStartDatePicker.SelectedDate = DateTime.Today;
            InspectionPeriodDaysTextBox.Text = "7";
            InspectionLastCompletedDatePicker.SelectedDate = DateTime.Today;
            InspectionNotesTextBox.Text = string.Empty;
        }

        private void FillInspectionForm(InspectionJournalEntry row)
        {
            if (row == null || InspectionReminderStartDatePicker == null)
                return;

            InspectionJournalNameBox.Text = row.JournalName ?? string.Empty;
            InspectionNameBox.Text = row.InspectionName ?? string.Empty;
            InspectionReminderStartDatePicker.SelectedDate = row.ReminderStartDate;
            InspectionPeriodDaysTextBox.Text = row.ReminderPeriodDays.ToString();
            InspectionLastCompletedDatePicker.SelectedDate = row.LastCompletedDate ?? row.ReminderStartDate;
            InspectionNotesTextBox.Text = row.Notes ?? string.Empty;
        }

        private void ReadInspectionForm(InspectionJournalEntry row)
        {
            row.JournalName = InspectionJournalNameBox.Text?.Trim();
            row.InspectionName = InspectionNameBox.Text?.Trim();
            row.ReminderStartDate = InspectionReminderStartDatePicker.SelectedDate ?? DateTime.Today;
            row.ReminderPeriodDays = int.TryParse(InspectionPeriodDaysTextBox.Text?.Trim(), out var days) && days > 0 ? days : 1;
            row.LastCompletedDate = InspectionLastCompletedDatePicker.SelectedDate;
            row.Notes = InspectionNotesTextBox.Text?.Trim();
        }

        private bool ValidateInspectionRow(InspectionJournalEntry row, out string message)
        {
            message = null;

            if (string.IsNullOrWhiteSpace(row.JournalName))
            {
                message = "Укажите название журнала.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(row.InspectionName))
            {
                message = "Укажите, какой осмотр нужно проводить.";
                return false;
            }

            if (row.ReminderPeriodDays <= 0)
            {
                message = "Периодичность должна быть больше нуля.";
                return false;
            }

            return true;
        }

        private void RestoreInspectionSnapshot(InspectionJournalEntry row)
        {
            if (row == null || string.IsNullOrWhiteSpace(inspectionRowSnapshotJson))
                return;

            var snapshot = JsonSerializer.Deserialize<InspectionJournalEntry>(inspectionRowSnapshotJson);
            if (snapshot == null)
                return;

            row.JournalName = snapshot.JournalName;
            row.InspectionName = snapshot.InspectionName;
            row.ReminderStartDate = snapshot.ReminderStartDate;
            row.ReminderPeriodDays = snapshot.ReminderPeriodDays;
            row.LastCompletedDate = snapshot.LastCompletedDate;
            row.Notes = snapshot.Notes;
        }

        private void MarkInspectionCompleted_Click(object sender, RoutedEventArgs e)
        {
            var row = sender is FrameworkElement fe
                ? fe.DataContext as InspectionJournalEntry
                : selectedInspectionRow ?? InspectionJournalGrid?.SelectedItem as InspectionJournalEntry;

            if (row == null)
            {
                MessageBox.Show("Выберите запись в осмотрах.");
                return;
            }

            if (row.IsCompletionHistory)
            {
                MessageBox.Show("Эта запись уже отмечена как проведенная.");
                return;
            }

            EnsureInspectionJournalStorage();

            var completedDate = DateTime.Today.Date;
            var periodDays = Math.Max(1, row.ReminderPeriodDays);
            var nextDate = completedDate.AddDays(periodDays);

            var completedRow = new InspectionJournalEntry
            {
                JournalName = row.JournalName,
                InspectionName = row.InspectionName,
                ReminderStartDate = row.ReminderStartDate,
                ReminderPeriodDays = row.ReminderPeriodDays,
                LastCompletedDate = completedDate,
                Notes = row.Notes,
                IsCompletionHistory = true
            };

            row.ReminderStartDate = nextDate;
            row.LastCompletedDate = null;
            row.IsCompletionHistory = false;

            currentObject?.InspectionJournal?.Add(completedRow);
            selectedInspectionRow = row;
            RefreshInspectionJournalState();
            if (InspectionJournalGrid != null)
                InspectionJournalGrid.SelectedItem = row;
            SaveState();
        }

        private void AddProductionRow_Click(object sender, RoutedEventArgs e) => SaveProductionForm_Click(sender, e);

        private void SaveProductionForm_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            EnsureProductionJournalStorage();
            var row = selectedProductionRow ?? new ProductionJournalEntry();
            productionRowSnapshotJson = JsonSerializer.Serialize(row);
            ReadProductionForm(row);
            if (!ApplyProductionRowChanges(row))
                return;

            if (selectedProductionRow == null)
                currentObject.ProductionJournal.Add(row);

            selectedProductionRow = row;
            RefreshProductionJournalState();
            ProductionJournalGrid.SelectedItem = row;
            ProductionJournalGrid.ScrollIntoView(row);
            SaveState();
        }

        private void DeleteProductionRow_Click(object sender, RoutedEventArgs e)
        {
            var row = selectedProductionRow ?? ProductionJournalGrid.SelectedItem as ProductionJournalEntry;
            if (row == null || currentObject?.ProductionJournal == null)
            {
                MessageBox.Show("Выберите запись в ПР.");
                return;
            }

            currentObject.ProductionJournal.Remove(row);
            selectedProductionRow = null;
            RefreshProductionJournalState();
            SaveState();
            RefreshSummaryTable();
        }

        private void ClearProductionForm_Click(object sender, RoutedEventArgs e)
        {
            selectedProductionRow = null;
            UpdateProductionFormState(clearFields: true);
            if (ProductionJournalGrid != null)
                ProductionJournalGrid.SelectedItem = null;
        }

        private void ProductionJournalGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProductionJournalGrid.SelectedItem is not ProductionJournalEntry row)
            {
                if (selectedProductionRow != null)
                    UpdateProductionFormState(clearFields: false);
                return;
            }

            selectedProductionRow = row;
            FillProductionForm(row);
            UpdateProductionFormState();
        }

        private void UpdateProductionFormState(bool clearFields = false)
        {
            if (clearFields)
                ResetProductionFormInputs();
            else if (selectedProductionRow == null && ProductionDatePicker != null && ProductionDatePicker.SelectedDate == null)
                ResetProductionFormInputs();

            if (ProductionSaveButton != null)
                ProductionSaveButton.Content = selectedProductionRow == null ? "➕ Добавить" : "💾 Обновить";
        }

        private List<string> GetAvailableProductionElements(string workName)
        {
            var values = new List<string>();
            if (!string.IsNullOrWhiteSpace(workName))
            {
                values.AddRange(currentObject?.MaterialCatalog?
                    .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals(x.TypeName ?? string.Empty, workName.Trim(), StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.MaterialName)
                    ?? Enumerable.Empty<string>());

                values.AddRange(journal
                    .Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals(x.MaterialGroup ?? string.Empty, workName.Trim(), StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.MaterialName));
            }
            else
            {
                values.AddRange(productionElements);
            }

            return values
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private void EditProductionItems_Click(object sender, RoutedEventArgs e)
        {
            var availableNames = GetAvailableProductionElements(ProductionWorkBox?.Text);
            var currentItems = ParseProductionItems(ProductionElementsBox?.Text);
            var result = PromptProductionItems(currentItems, availableNames);
            if (result == null)
                return;

            ProductionElementsBox.Text = FormatProductionItems(result);
        }

        private List<ProductionItemQuantity> PromptProductionItems(IEnumerable<ProductionItemQuantity> currentItems, List<string> availableNames)
        {
            var rows = new ObservableCollection<ProductionItemEditorRow>(
                (currentItems ?? Enumerable.Empty<ProductionItemQuantity>())
                .Select(x => new ProductionItemEditorRow
                {
                    MaterialName = x.MaterialName,
                    Quantity = x.Quantity,
                    AvailableNames = new ObservableCollection<string>(availableNames ?? new List<string>())
                }));

            if (rows.Count == 0)
            {
                rows.Add(new ProductionItemEditorRow
                {
                    Quantity = 1,
                    AvailableNames = new ObservableCollection<string>(availableNames ?? new List<string>())
                });
            }

            var dialog = new Window
            {
                Title = "Элементы и количество",
                Owner = this,
                Width = 760,
                Height = 520,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var hint = new TextBlock
            {
                Text = "Количество вводится отдельно для каждого элемента. \"шт\" писать не нужно.",
                Margin = new Thickness(0, 0, 0, 12),
                Foreground = new SolidColorBrush(Color.FromRgb(71, 85, 105))
            };
            Grid.SetRow(hint, 0);
            root.Children.Add(hint);

            var rowsScroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
            Grid.SetRow(rowsScroll, 1);
            root.Children.Add(rowsScroll);

            var rowsPanel = new StackPanel();
            rowsScroll.Content = rowsPanel;

            void RenderRows()
            {
                rowsPanel.Children.Clear();
                foreach (var row in rows.ToList())
                {
                    var rowGrid = new Grid { Margin = new Thickness(0, 0, 0, 10) };
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                    var elementBox = new ComboBox
                    {
                        IsEditable = true,
                        IsTextSearchEnabled = true,
                        StaysOpenOnEdit = true,
                        ItemsSource = row.AvailableNames,
                        Margin = new Thickness(0, 0, 10, 0)
                    };
                    elementBox.SetBinding(ComboBox.TextProperty, new Binding(nameof(ProductionItemEditorRow.MaterialName))
                    {
                        Source = row,
                        Mode = BindingMode.TwoWay,
                        UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                    });
                    Grid.SetColumn(elementBox, 0);
                    rowGrid.Children.Add(elementBox);

                    var quantityBox = new TextBox
                    {
                        Margin = new Thickness(0, 0, 10, 0),
                        VerticalContentAlignment = VerticalAlignment.Center
                    };
                    quantityBox.SetBinding(TextBox.TextProperty, new Binding(nameof(ProductionItemEditorRow.Quantity))
                    {
                        Source = row,
                        Mode = BindingMode.TwoWay,
                        UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                        StringFormat = "0.##"
                    });
                    Grid.SetColumn(quantityBox, 1);
                    rowGrid.Children.Add(quantityBox);

                    var deleteButton = new Button
                    {
                        Content = "Удалить",
                        Style = FindResource("SecondaryButton") as Style,
                        MinWidth = 90
                    };
                    deleteButton.Click += (_, _) =>
                    {
                        rows.Remove(row);
                        if (rows.Count == 0)
                        {
                            rows.Add(new ProductionItemEditorRow
                            {
                                Quantity = 1,
                                AvailableNames = new ObservableCollection<string>(availableNames ?? new List<string>())
                            });
                        }
                        RenderRows();
                    };
                    Grid.SetColumn(deleteButton, 2);
                    rowGrid.Children.Add(deleteButton);

                    rowsPanel.Children.Add(rowGrid);
                }
            }

            RenderRows();

            var footer = new DockPanel { Margin = new Thickness(0, 12, 0, 0) };
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            var addButton = new Button
            {
                Content = "Добавить строку",
                Style = FindResource("SecondaryButton") as Style,
                MinWidth = 140
            };
            addButton.Click += (_, _) =>
            {
                rows.Add(new ProductionItemEditorRow
                {
                    Quantity = 1,
                    AvailableNames = new ObservableCollection<string>(availableNames ?? new List<string>())
                });
                RenderRows();
            };
            DockPanel.SetDock(addButton, Dock.Left);
            footer.Children.Add(addButton);

            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            DockPanel.SetDock(buttons, Dock.Right);
            footer.Children.Add(buttons);

            var okButton = new Button
            {
                Content = "Сохранить",
                MinWidth = 118,
                IsDefault = true
            };
            var cancelButton = new Button
            {
                Content = "Отмена",
                Style = FindResource("SecondaryButton") as Style,
                MinWidth = 110,
                IsCancel = true,
                Margin = new Thickness(10, 0, 0, 0)
            };
            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);

            List<ProductionItemQuantity> result = null;
            okButton.Click += (_, _) =>
            {
                result = rows
                    .Select(x => new ProductionItemQuantity
                    {
                        MaterialName = x.MaterialName?.Trim(),
                        Quantity = x.Quantity
                    })
                    .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName) && x.Quantity > 0)
                    .ToList();

                dialog.DialogResult = true;
            };

            dialog.Content = root;
            return dialog.ShowDialog() == true ? result : null;
        }

        private void ConfigureProductionAutoFill_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            EnsureProductionJournalStorage();
            var settings = currentObject.ProductionAutoFillSettings ??= new ProductionAutoFillSettings();

            var dialog = new Window
            {
                Title = "Настройки автомастера ПР",
                Owner = this,
                Width = 520,
                Height = 560,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize
            };

            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            TextBox CreateNumericBox(string text, int rowIndex, string caption)
            {
                var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };
                Grid.SetRow(panel, rowIndex);
                panel.Children.Add(new TextBlock { Text = caption, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 6) });
                var box = new TextBox { Text = text };
                panel.Children.Add(box);
                root.Children.Add(panel);
                return box;
            }

            var minBox = CreateNumericBox(settings.MinQuantityPerRow.ToString(), 0, "Минимум в строке");
            var maxBox = CreateNumericBox(settings.MaxQuantityPerRow.ToString(), 1, "Максимум в строке");
            var minRowsBox = CreateNumericBox(settings.MinRowsPerRun.ToString(), 2, "Минимум строк за один запуск");
            var targetRowsBox = CreateNumericBox(settings.TargetRowsPerRun.ToString(), 3, "Целевое число строк");
            var maxRowsBox = CreateNumericBox(settings.MaxRowsPerRun.ToString(), 4, "Максимум строк за один запуск");
            var itemsPerRowBox = CreateNumericBox(settings.MaxItemsPerRow.ToString(), 5, "Максимум элементов в одной строке");

            var preferTypeCheck = new CheckBox
            {
                Content = "Брать материалы только из выбранного типа",
                IsChecked = settings.PreferSelectedTypeOnly,
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(preferTypeCheck, 6);
            root.Children.Add(preferTypeCheck);

            var balanceCheck = new CheckBox
            {
                Content = "Распределять количество более равномерно",
                IsChecked = settings.UseBalancedDistribution,
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(balanceCheck, 7);
            root.Children.Add(balanceCheck);

            var deficitCheck = new CheckBox
            {
                Content = "Сначала закрывать дефицит по сводке",
                IsChecked = settings.PreferDemandDeficit,
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(deficitCheck, 8);
            root.Children.Add(deficitCheck);

            var selectedCellsCheck = new CheckBox
            {
                Content = "Учитывать только выбранные блоки и отметки из формы",
                IsChecked = settings.RespectSelectedBlocksAndMarks,
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(selectedCellsCheck, 9);
            root.Children.Add(selectedCellsCheck);

            var mixedRowsCheck = new CheckBox
            {
                Content = "Разрешать несколько элементов в одной строке",
                IsChecked = settings.AllowMixedMaterialsInRow
            };
            Grid.SetRow(mixedRowsCheck, 10);
            root.Children.Add(mixedRowsCheck);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            Grid.SetRow(footer, 11);
            root.Children.Add(footer);

            var saveButton = new Button { Content = "Сохранить", MinWidth = 120, IsDefault = true };
            var cancelButton = new Button { Content = "Отмена", Style = FindResource("SecondaryButton") as Style, MinWidth = 110, Margin = new Thickness(10, 0, 0, 0), IsCancel = true };
            footer.Children.Add(saveButton);
            footer.Children.Add(cancelButton);

            saveButton.Click += (_, _) =>
            {
                settings.MinQuantityPerRow = Math.Max(1, int.TryParse(minBox.Text, out var min) ? min : 4);
                settings.MaxQuantityPerRow = Math.Max(settings.MinQuantityPerRow, int.TryParse(maxBox.Text, out var max) ? max : 8);
                settings.MinRowsPerRun = Math.Clamp(int.TryParse(minRowsBox.Text, out var minRows) ? minRows : 4, 1, 12);
                settings.TargetRowsPerRun = Math.Clamp(int.TryParse(targetRowsBox.Text, out var targetRows) ? targetRows : 5, settings.MinRowsPerRun, 16);
                settings.MaxRowsPerRun = Math.Clamp(int.TryParse(maxRowsBox.Text, out var maxRows) ? maxRows : 6, settings.TargetRowsPerRun, 20);
                settings.MaxItemsPerRow = Math.Clamp(int.TryParse(itemsPerRowBox.Text, out var maxItems) ? maxItems : 2, 1, 6);
                settings.PreferSelectedTypeOnly = preferTypeCheck.IsChecked == true;
                settings.UseBalancedDistribution = balanceCheck.IsChecked == true;
                settings.PreferDemandDeficit = deficitCheck.IsChecked == true;
                settings.RespectSelectedBlocksAndMarks = selectedCellsCheck.IsChecked == true;
                settings.AllowMixedMaterialsInRow = mixedRowsCheck.IsChecked == true;
                dialog.DialogResult = true;
            };

            dialog.Content = root;
            if (dialog.ShowDialog() == true)
                SaveState();
        }

        private void AutoFillProduction_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            EnsureProductionJournalStorage();
            var settings = currentObject.ProductionAutoFillSettings ??= new ProductionAutoFillSettings();
            var baseDate = ProductionDatePicker?.SelectedDate ?? DateTime.Today;
            var baseAction = string.IsNullOrWhiteSpace(ProductionActionBox?.Text) ? (productionActions.FirstOrDefault() ?? "Монтаж") : ProductionActionBox.Text.Trim();
            var baseWork = ProductionWorkBox?.Text?.Trim();
            var baseBlocks = string.IsNullOrWhiteSpace(ProductionBlocksBox?.Text) ? (productionBlockOptions.FirstOrDefault() ?? "1") : ProductionBlocksBox.Text.Trim();
            var baseMarks = string.IsNullOrWhiteSpace(ProductionMarksBox?.Text) ? (productionMarkOptions.FirstOrDefault() ?? string.Empty) : ProductionMarksBox.Text.Trim();
            var baseBrigade = ProductionBrigadeBox?.Text?.Trim();
            var baseWeather = ProductionWeatherBox?.Text?.Trim();
            var baseDeviation = ProductionDeviationBox?.Text?.Trim();
            var blocks = settings.RespectSelectedBlocksAndMarks
                ? LevelMarkHelper.ParseBlocks(baseBlocks)
                : Enumerable.Range(1, Math.Max(1, currentObject.BlocksCount)).ToList();
            var groupForMarks = ResolveProductionAutoFillGroup(baseWork);
            var marks = settings.RespectSelectedBlocksAndMarks
                ? LevelMarkHelper.ParseMarks(baseMarks)
                : LevelMarkHelper.GetMarksForGroup(currentObject, groupForMarks);

            if (blocks.Count == 0)
                blocks = Enumerable.Range(1, Math.Max(1, currentObject.BlocksCount)).ToList();

            if (marks.Count == 0)
                marks = LevelMarkHelper.GetMarksForGroup(currentObject, groupForMarks);

            if (blocks.Count > 0)
                baseBlocks = string.Join(", ", blocks);
            if (marks.Count > 0)
                baseMarks = string.Join(", ", marks);
            if (string.IsNullOrWhiteSpace(baseWork))
                baseWork = groupForMarks;

            if (ProductionWorkBox != null && !string.IsNullOrWhiteSpace(baseWork))
                ProductionWorkBox.Text = baseWork;
            if (ProductionBlocksBox != null && !string.IsNullOrWhiteSpace(baseBlocks))
                ProductionBlocksBox.Text = baseBlocks;
            if (ProductionMarksBox != null && !string.IsNullOrWhiteSpace(baseMarks))
                ProductionMarksBox.Text = baseMarks;
            RefreshProductionElementOptions();

            var candidates = BuildProductionAutoFillCandidates(settings, groupForMarks, blocks, marks);

            if (candidates.Count == 0)
            {
                MessageBox.Show("Нет доступного прихода или дефицита для автозаполнения ПР.");
                return;
            }

            var targetRows = Math.Clamp(settings.TargetRowsPerRun, settings.MinRowsPerRun, settings.MaxRowsPerRun);
            var plannedRows = BuildProductionAutoFillPlan(candidates, settings, targetRows);
            if (plannedRows.Count == 0)
            {
                MessageBox.Show("Автомастер не смог подобрать строки по текущим ограничениям.");
                return;
            }

            var addedRows = new List<ProductionJournalEntry>();
            foreach (var plannedItems in plannedRows)
            {
                var row = new ProductionJournalEntry
                {
                    Date = baseDate,
                    ActionName = baseAction,
                    WorkName = string.IsNullOrWhiteSpace(baseWork) ? groupForMarks : baseWork,
                    ElementsText = FormatProductionItems(plannedItems),
                    BlocksText = baseBlocks,
                    MarksText = baseMarks,
                    BrigadeName = baseBrigade,
                    Weather = baseWeather,
                    Deviations = baseDeviation,
                    RequiresHiddenWorkAct = ProductionHiddenWorksCheckBox?.IsChecked == true
                };

                if (!ApplyProductionRowChanges(row))
                    continue;

                currentObject.ProductionJournal.Add(row);
                addedRows.Add(row);
            }

            if (addedRows.Count == 0)
            {
                MessageBox.Show("Автозаполнение не добавило строки. Возможно, доступные количества уже исчерпаны.");
                return;
            }

            selectedProductionRow = addedRows.Last();
            RefreshProductionJournalState();
            ProductionJournalGrid.SelectedItem = selectedProductionRow;
            ProductionJournalGrid.ScrollIntoView(selectedProductionRow);
            SaveState();
            MessageBox.Show($"Автозаполнение добавило {addedRows.Count} строк(и) в ПР по выбранным блокам и отметкам.");
        }

        private string ResolveProductionAutoFillGroup(string currentWork)
        {
            if (!string.IsNullOrWhiteSpace(currentWork))
                return currentWork.Trim();

            if (ObjectsTree?.SelectedItem is TreeViewItem node)
            {
                foreach (var currentNode in EnumerateNodeWithParents(node))
                {
                    if (GetNodeKind(currentNode) == "Group")
                        return currentNode.Header?.ToString()?.Trim();

                    if (currentNode.Tag is TreeNodeMeta meta && !string.IsNullOrWhiteSpace(meta.GroupName))
                        return meta.GroupName.Trim();
                }
            }

            return currentObject?.SummaryVisibleGroups?.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))
                ?? productionTargets.FirstOrDefault()
                ?? string.Empty;
        }

        private List<AutoProductionCandidate> BuildProductionAutoFillCandidates(
            ProductionAutoFillSettings settings,
            string groupFilter,
            List<int> blocks,
            List<string> marks)
        {
            var selectedBlocks = (blocks ?? new List<int>()).Where(x => x > 0).Distinct().ToList();
            var selectedMarks = (marks ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var candidates = journal
                .Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                         && !string.IsNullOrWhiteSpace(x.MaterialName)
                         && (!settings.PreferSelectedTypeOnly
                             || string.IsNullOrWhiteSpace(groupFilter)
                             || string.Equals(x.MaterialGroup ?? string.Empty, groupFilter, StringComparison.CurrentCultureIgnoreCase)))
                .GroupBy(x => new
                {
                    Group = (x.MaterialGroup ?? string.Empty).Trim(),
                    Material = (x.MaterialName ?? string.Empty).Trim()
                })
                .Select(x =>
                {
                    var group = x.Key.Group;
                    var material = x.Key.Material;
                    var unit = x.Select(y => y.Unit).FirstOrDefault(y => !string.IsNullOrWhiteSpace(y))
                               ?? GetUnitForMaterial(group, material);
                    var arrived = x.Sum(y => y.Quantity);
                    var mounted = GetMountedQuantityFromProductionJournal(material);
                    var available = Math.Max(0, arrived - mounted);
                    var deficit = CalculateProductionDeficit(group, material, unit, settings, selectedBlocks, selectedMarks);
                    var remaining = settings.PreferDemandDeficit
                        ? Math.Min(available, deficit > 0 ? deficit : available)
                        : available;

                    return new AutoProductionCandidate
                    {
                        Group = group,
                        Material = material,
                        Unit = unit,
                        Available = available,
                        Deficit = deficit,
                        RemainingToPlan = remaining
                    };
                })
                .Where(x => x.Available > 0.0001 && x.RemainingToPlan > 0.0001)
                .OrderByDescending(x => settings.PreferDemandDeficit ? x.Deficit : x.Available)
                .ThenByDescending(x => x.Available)
                .ThenBy(x => x.Material, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (settings.PreferDemandDeficit && candidates.Any(x => x.Deficit > 0.0001))
                candidates = candidates.Where(x => x.Deficit > 0.0001).ToList();

            return candidates;
        }

        private double CalculateProductionDeficit(
            string group,
            string material,
            string unit,
            ProductionAutoFillSettings settings,
            List<int> blocks,
            List<string> marks)
        {
            if (string.IsNullOrWhiteSpace(group) || string.IsNullOrWhiteSpace(material))
                return 0;

            var demand = GetOrCreateDemand(BuildDemandKey(group, material), unit);
            var availableBlocks = settings.RespectSelectedBlocksAndMarks && blocks.Count > 0
                ? blocks
                : Enumerable.Range(1, Math.Max(1, currentObject?.BlocksCount ?? 1)).ToList();
            var availableMarks = settings.RespectSelectedBlocksAndMarks && marks.Count > 0
                ? marks
                : LevelMarkHelper.GetMarksForGroup(currentObject, group);

            double total = 0;
            foreach (var block in availableBlocks)
            {
                foreach (var mark in availableMarks)
                {
                    var planned = GetDemandValue(demand, block, mark);
                    var mounted = GetMountedValue(demand, block, mark);
                    total += Math.Max(0, planned - mounted);
                }
            }

            return total;
        }

        private List<List<ProductionItemQuantity>> BuildProductionAutoFillPlan(
            List<AutoProductionCandidate> candidates,
            ProductionAutoFillSettings settings,
            int targetRows)
        {
            var result = new List<List<ProductionItemQuantity>>();
            if (candidates == null || candidates.Count == 0)
                return result;

            var working = candidates
                .Select(x => new AutoProductionCandidate
                {
                    Group = x.Group,
                    Material = x.Material,
                    Unit = x.Unit,
                    Available = x.Available,
                    Deficit = x.Deficit,
                    RemainingToPlan = x.RemainingToPlan
                })
                .ToList();

            var maxRows = Math.Clamp(settings.MaxRowsPerRun, 1, 20);
            var desiredRows = Math.Clamp(targetRows, 1, maxRows);
            var maxItemsPerRow = settings.AllowMixedMaterialsInRow
                ? Math.Max(1, settings.MaxItemsPerRow)
                : 1;

            while (result.Count < maxRows)
            {
                var remainingCandidates = working
                    .Where(x => x.RemainingToPlan > 0.0001)
                    .OrderByDescending(x => settings.PreferDemandDeficit ? x.Deficit : x.Available)
                    .ThenByDescending(x => x.RemainingToPlan)
                    .ThenBy(x => x.Material, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                if (remainingCandidates.Count == 0)
                    break;

                var rowItems = new List<ProductionItemQuantity>();
                var itemsInRow = Math.Min(maxItemsPerRow, remainingCandidates.Count);

                for (var itemIndex = 0; itemIndex < itemsInRow; itemIndex++)
                {
                    var candidate = remainingCandidates[itemIndex];
                    var rowsRemaining = Math.Max(1, desiredRows - result.Count);
                    var quantity = CalculateAutoFillChunk(candidate.RemainingToPlan, rowsRemaining, settings);
                    if (quantity <= 0.0001)
                        continue;

                    rowItems.Add(new ProductionItemQuantity
                    {
                        MaterialName = candidate.Material,
                        Quantity = quantity
                    });

                    candidate.RemainingToPlan = Math.Max(0, candidate.RemainingToPlan - quantity);
                    candidate.Available = Math.Max(0, candidate.Available - quantity);
                    candidate.Deficit = Math.Max(0, candidate.Deficit - quantity);
                }

                if (rowItems.Count == 0)
                    break;

                result.Add(rowItems);

                if (result.Count >= desiredRows)
                {
                    var anyBigRemainders = working.Any(x => x.RemainingToPlan > settings.MaxQuantityPerRow + 0.0001);
                    if (!anyBigRemainders)
                        break;
                }
            }

            while (result.Count < settings.MinRowsPerRun)
            {
                var extra = working
                    .Where(x => x.RemainingToPlan > 0.0001)
                    .OrderByDescending(x => x.RemainingToPlan)
                    .FirstOrDefault();

                if (extra == null)
                    break;

                var quantity = CalculateAutoFillChunk(extra.RemainingToPlan, 1, settings);
                if (quantity <= 0.0001)
                    break;

                result.Add(new List<ProductionItemQuantity>
                {
                    new ProductionItemQuantity
                    {
                        MaterialName = extra.Material,
                        Quantity = quantity
                    }
                });

                extra.RemainingToPlan = Math.Max(0, extra.RemainingToPlan - quantity);
                if (result.Count >= settings.MaxRowsPerRun)
                    break;
            }

            return result.Take(settings.MaxRowsPerRun).ToList();
        }

        private double CalculateAutoFillChunk(double remaining, int rowsRemaining, ProductionAutoFillSettings settings)
        {
            if (remaining <= 0.0001)
                return 0;

            var min = Math.Max(1, settings.MinQuantityPerRow);
            var max = Math.Max(min, settings.MaxQuantityPerRow);
            var averageTarget = settings.UseBalancedDistribution
                ? Math.Max(min, Math.Min(max, Math.Ceiling(remaining / Math.Max(1, rowsRemaining))))
                : productionAutoRandom.Next(min, max + 1);

            var quantity = Math.Min(remaining, averageTarget);

            if (rowsRemaining <= 1)
                return Math.Round(remaining, 2);

            if (remaining > max * rowsRemaining)
                quantity = Math.Ceiling(remaining / rowsRemaining);
            else if (remaining < min)
                quantity = remaining;

            return Math.Round(Math.Max(1, quantity), 2);
        }

        private void ResetProductionFormInputs()
        {
            if (ProductionDatePicker == null)
                return;

            ProductionDatePicker.SelectedDate = DateTime.Today;
            ProductionActionBox.Text = productionActions.FirstOrDefault() ?? "Монтаж";
            ProductionWorkBox.Text = productionTargets.FirstOrDefault() ?? string.Empty;
            ProductionElementsBox.Text = string.Empty;
            ProductionBlocksBox.Text = productionBlockOptions.FirstOrDefault() ?? "1";
            ProductionMarksBox.Text = productionMarkOptions.FirstOrDefault() ?? string.Empty;
            ProductionBrigadeBox.Text = timesheetAssignableBrigades.FirstOrDefault() ?? string.Empty;
            ProductionWeatherBox.Text = string.Empty;
            ProductionDeviationBox.Text = string.Empty;
            ProductionHiddenWorksCheckBox.IsChecked = false;
            RefreshProductionElementOptions();
        }

        private void FillProductionForm(ProductionJournalEntry row)
        {
            if (row == null || ProductionDatePicker == null)
                return;

            ProductionDatePicker.SelectedDate = row.Date;
            ProductionActionBox.Text = row.ActionName ?? string.Empty;
            ProductionWorkBox.Text = row.WorkName ?? string.Empty;
            RefreshProductionElementOptions();
            ProductionElementsBox.Text = row.ElementsText ?? string.Empty;
            ProductionBlocksBox.Text = row.BlocksText ?? string.Empty;
            ProductionMarksBox.Text = row.MarksText ?? string.Empty;
            ProductionBrigadeBox.Text = row.BrigadeName ?? string.Empty;
            ProductionWeatherBox.Text = row.Weather ?? string.Empty;
            ProductionDeviationBox.Text = row.Deviations ?? string.Empty;
            ProductionHiddenWorksCheckBox.IsChecked = row.RequiresHiddenWorkAct;
        }

        private void ReadProductionForm(ProductionJournalEntry row)
        {
            row.Date = ProductionDatePicker.SelectedDate ?? DateTime.Today;
            row.ActionName = ProductionActionBox.Text?.Trim();
            row.WorkName = ProductionWorkBox.Text?.Trim();
            row.ElementsText = ProductionElementsBox.Text?.Trim();
            row.BlocksText = ProductionBlocksBox.Text?.Trim();
            row.MarksText = ProductionMarksBox.Text?.Trim();
            row.BrigadeName = ProductionBrigadeBox.Text?.Trim();
            row.Weather = ProductionWeatherBox.Text?.Trim();
            row.Deviations = ProductionDeviationBox.Text?.Trim();
            row.RequiresHiddenWorkAct = ProductionHiddenWorksCheckBox.IsChecked == true;
        }

        private bool ApplyProductionRowChanges(ProductionJournalEntry row)
        {
            if (row == null || currentObject?.ProductionJournal == null)
                return false;

            ApplyProductionDefaults(row);

            var originalItems = ParseProductionItems(row.ElementsText);
            var adjustedMessages = new List<string>();
            var adjustedItems = AdjustProductionItems(row, originalItems, adjustedMessages);
            row.ElementsText = FormatProductionItems(adjustedItems);

            if (!ValidateProductionRow(row, out var message))
            {
                RestoreProductionRowSnapshot(row);
                MessageBox.Show(message, "Проверка ПР", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (adjustedItems.Count == 0)
            {
                if (selectedProductionRow != null)
                    currentObject.ProductionJournal.Remove(row);

                selectedProductionRow = null;
                RefreshProductionJournalState();
                RefreshSummaryTable();
                SaveState();
                MessageBox.Show("Для выбранных элементов больше нет доступного количества. Строка не сохранена.", "Проверка ПР", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            EnsureArmoringCompanionRow(row);
            NormalizeProductionJournalRows();
            RefreshProductionJournalLookups();
            RebuildMountedDemandFromProductionJournal();
            RefreshProductionRemainingInfo();
            RefreshSummaryTable();
            SaveState();

            if (adjustedMessages.Count > 0)
            {
                MessageBox.Show(string.Join(Environment.NewLine, adjustedMessages), "Проверка ПР", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            return true;
        }

        private void ApplyProductionDefaults(ProductionJournalEntry row)
        {
            if (row == null)
                return;

            if (string.IsNullOrWhiteSpace(row.ActionName))
                row.ActionName = productionActions.FirstOrDefault() ?? "Монтаж";

            if (string.IsNullOrWhiteSpace(row.WorkName))
            {
                row.WorkName = ParseProductionItems(row.ElementsText)
                    .Select(x => FindMaterialGroupByName(x.MaterialName))
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? row.WorkName;
            }

            if (string.IsNullOrWhiteSpace(row.Deviations))
            {
                var previousDeviation = currentObject?.ProductionJournal?
                    .Where(x => !ReferenceEquals(x, row)
                        && string.Equals(x.ActionName?.Trim(), row.ActionName?.Trim(), StringComparison.CurrentCultureIgnoreCase)
                        && string.Equals(x.WorkName?.Trim(), row.WorkName?.Trim(), StringComparison.CurrentCultureIgnoreCase)
                        && !string.IsNullOrWhiteSpace(x.Deviations))
                    .OrderByDescending(x => x.Date)
                    .Select(x => x.Deviations?.Trim())
                    .FirstOrDefault();

                if (!string.IsNullOrWhiteSpace(previousDeviation))
                    row.Deviations = previousDeviation;
            }

            if (!row.RequiresHiddenWorkAct && !string.IsNullOrWhiteSpace(row.ActionName))
            {
                var previousValue = currentObject?.ProductionJournal?
                    .Where(x => !ReferenceEquals(x, row)
                        && string.Equals(x.ActionName?.Trim(), row.ActionName?.Trim(), StringComparison.CurrentCultureIgnoreCase))
                    .OrderByDescending(x => x.Date)
                    .Select(x => x.RequiresHiddenWorkAct)
                    .FirstOrDefault();

                if (previousValue == true)
                    row.RequiresHiddenWorkAct = true;
            }

            if (string.IsNullOrWhiteSpace(row.Weather))
            {
                var weatherForDate = currentObject?.ProductionJournal?
                    .Where(x => !ReferenceEquals(x, row)
                        && x.Date.Date == row.Date.Date
                        && !string.IsNullOrWhiteSpace(x.Weather))
                    .OrderByDescending(x => x.Date)
                    .Select(x => x.Weather?.Trim())
                    .FirstOrDefault();

                if (!string.IsNullOrWhiteSpace(weatherForDate))
                    row.Weather = weatherForDate;
            }
        }

        private void RestoreProductionRowSnapshot(ProductionJournalEntry row)
        {
            if (row == null || string.IsNullOrWhiteSpace(productionRowSnapshotJson))
                return;

            var snapshot = JsonSerializer.Deserialize<ProductionJournalEntry>(productionRowSnapshotJson);
            if (snapshot == null)
                return;

            row.Date = snapshot.Date;
            row.ActionName = snapshot.ActionName;
            row.WorkName = snapshot.WorkName;
            row.ElementsText = snapshot.ElementsText;
            row.BlocksText = snapshot.BlocksText;
            row.MarksText = snapshot.MarksText;
            row.BrigadeName = snapshot.BrigadeName;
            row.Weather = snapshot.Weather;
            row.Deviations = snapshot.Deviations;
            row.RequiresHiddenWorkAct = snapshot.RequiresHiddenWorkAct;
            row.RemainingInfo = snapshot.RemainingInfo;
        }

        private bool ValidateProductionRow(ProductionJournalEntry row, out string message)
        {
            message = null;

            var items = ParseProductionItems(row.ElementsText);
            if (items.Count == 0)
                return true;

            var blocks = LevelMarkHelper.ParseBlocks(row.BlocksText);
            if (blocks.Count == 0)
            {
                message = "Для записи ПР укажите хотя бы один блок.";
                return false;
            }

            var marks = LevelMarkHelper.ParseMarks(row.MarksText);
            if (marks.Count == 0)
            {
                message = "Для записи ПР укажите хотя бы одну отметку.";
                return false;
            }

            foreach (var item in items)
            {
                var group = FindMaterialGroupByName(item.MaterialName);
                if (string.IsNullOrWhiteSpace(group))
                    continue;

                var alreadyMounted = GetMountedQuantityFromProductionJournal(item.MaterialName, excludeRow: row);
                var arrived = journal
                    .Where(x => x.Category == "Основные"
                        && string.Equals(x.MaterialGroup ?? string.Empty, group, StringComparison.CurrentCultureIgnoreCase)
                        && string.Equals(x.MaterialName ?? string.Empty, item.MaterialName, StringComparison.CurrentCultureIgnoreCase))
                    .Sum(x => x.Quantity);

                if (alreadyMounted + item.Quantity > arrived + 0.0001)
                {
                    message = $"Для \"{item.MaterialName}\" нельзя указать больше, чем пришло. Пришло: {FormatNumber(arrived)}, уже записано в ПР: {FormatNumber(alreadyMounted)}.";
                    return false;
                }
            }

            return true;
        }

        private sealed class ProductionItemQuantity
        {
            public string MaterialName { get; set; }
            public double Quantity { get; set; }
        }

        private List<ProductionItemQuantity> ParseProductionItems(string text)
        {
            var result = new List<ProductionItemQuantity>();
            foreach (var chunk in LevelMarkHelper.SplitText(text))
            {
                var match = Regex.Match(chunk, @"^(?<name>.*?)(?:\s*[-–]\s*|\s+)(?<qty>\d+(?:[.,]\d+)?)\s*(?:шт\.?|шт)?$", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    result.Add(new ProductionItemQuantity
                    {
                        MaterialName = match.Groups["name"].Value.Trim(),
                        Quantity = ParseNumber(match.Groups["qty"].Value)
                    });
                }
                else
                {
                    result.Add(new ProductionItemQuantity
                    {
                        MaterialName = chunk.Trim(),
                        Quantity = 1
                    });
                }
            }

            return result
                .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName) && x.Quantity > 0)
                .ToList();
        }

        private List<ProductionItemQuantity> AdjustProductionItems(ProductionJournalEntry row, List<ProductionItemQuantity> items, List<string> adjustedMessages)
        {
            var result = new List<ProductionItemQuantity>();

            foreach (var item in items)
            {
                var group = FindMaterialGroupByName(item.MaterialName) ?? row.WorkName?.Trim();
                var alreadyMounted = GetMountedQuantityFromProductionJournal(item.MaterialName, excludeRow: row);
                var arrived = journal
                    .Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals(x.MaterialName ?? string.Empty, item.MaterialName, StringComparison.CurrentCultureIgnoreCase)
                             && (string.IsNullOrWhiteSpace(group)
                                 || string.Equals(x.MaterialGroup ?? string.Empty, group, StringComparison.CurrentCultureIgnoreCase)))
                    .Sum(x => x.Quantity);

                var allowed = Math.Max(0, arrived - alreadyMounted);
                var finalQuantity = Math.Min(item.Quantity, allowed);
                if (finalQuantity <= 0)
                {
                    adjustedMessages?.Add($"\"{item.MaterialName}\" не добавлен: доступное количество закончилось.");
                    continue;
                }

                if (finalQuantity + 0.0001 < item.Quantity)
                {
                    adjustedMessages?.Add($"\"{item.MaterialName}\" скорректирован до {FormatNumber(finalQuantity)}. Больше записать нельзя: пришло {FormatNumber(arrived)}, уже смонтировано {FormatNumber(alreadyMounted)}.");
                }

                result.Add(new ProductionItemQuantity
                {
                    MaterialName = item.MaterialName,
                    Quantity = finalQuantity
                });
            }

            return result;
        }

        private static string FormatProductionItems(IEnumerable<ProductionItemQuantity> items)
        {
            return string.Join(", ", (items ?? Enumerable.Empty<ProductionItemQuantity>())
                .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName) && x.Quantity > 0)
                .Select(x => $"{x.MaterialName.Trim()} - {x.Quantity:0.##}"));
        }

        private void NormalizeProductionJournalRows()
        {
            if (currentObject?.ProductionJournal == null || currentObject.ProductionJournal.Count == 0)
                return;

            foreach (var dateGroup in currentObject.ProductionJournal.GroupBy(x => x.Date.Date))
            {
                var weather = dateGroup
                    .Select(x => x.Weather?.Trim())
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));

                if (string.IsNullOrWhiteSpace(weather))
                    continue;

                foreach (var row in dateGroup)
                    row.Weather = weather;
            }

            var merged = new List<ProductionJournalEntry>();
            foreach (var dateGroup in currentObject.ProductionJournal
                .OrderBy(x => x.Date)
                .GroupBy(x => x.Date.Date))
            {
                var map = new Dictionary<string, ProductionJournalEntry>(StringComparer.CurrentCultureIgnoreCase);
                foreach (var row in dateGroup)
                {
                    var key = BuildProductionMergeKey(row);
                    if (!map.TryGetValue(key, out var existing))
                    {
                        map[key] = row;
                        continue;
                    }

                    existing.ElementsText = FormatProductionItems(MergeProductionItems(existing.ElementsText, row.ElementsText));
                }

                merged.AddRange(map.Values.OrderBy(x => x.WorkName).ThenBy(x => x.BlocksText).ThenBy(x => x.MarksText));
            }

            currentObject.ProductionJournal = merged;
        }

        private static string BuildProductionMergeKey(ProductionJournalEntry row)
        {
            return string.Join("|",
                row.Date.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                (row.ActionName ?? string.Empty).Trim().ToUpperInvariant(),
                (row.WorkName ?? string.Empty).Trim().ToUpperInvariant(),
                NormalizeCsvText(row.BlocksText),
                NormalizeCsvText(row.MarksText),
                (row.BrigadeName ?? string.Empty).Trim().ToUpperInvariant(),
                (row.Deviations ?? string.Empty).Trim().ToUpperInvariant(),
                row.RequiresHiddenWorkAct ? "1" : "0");
        }

        private static string NormalizeCsvText(string text)
        {
            return string.Join("|", LevelMarkHelper.SplitText(text ?? string.Empty)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .Select(x => x.ToUpperInvariant()));
        }

        private List<ProductionItemQuantity> MergeProductionItems(string left, string right)
        {
            var merged = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

            void addItems(IEnumerable<ProductionItemQuantity> items)
            {
                foreach (var item in items.Where(x => !string.IsNullOrWhiteSpace(x.MaterialName) && x.Quantity > 0))
                {
                    var key = item.MaterialName.Trim();
                    if (!merged.ContainsKey(key))
                        merged[key] = 0;
                    merged[key] += item.Quantity;
                }
            }

            addItems(ParseProductionItems(left));
            addItems(ParseProductionItems(right));

            return merged
                .OrderBy(x => x.Key)
                .Select(x => new ProductionItemQuantity { MaterialName = x.Key, Quantity = x.Value })
                .ToList();
        }

        private void EnsureArmoringCompanionRow(ProductionJournalEntry row)
        {
            if (row == null
                || currentObject?.ProductionJournal == null
                || !string.Equals(row.ActionName?.Trim(), "Устройство", StringComparison.CurrentCultureIgnoreCase)
                || string.IsNullOrWhiteSpace(row.WorkName)
                || row.Date <= DateTime.MinValue)
                return;

            var workName = row.WorkName.Trim();
            if (workName.IndexOf("бетонир", StringComparison.CurrentCultureIgnoreCase) < 0)
                return;

            var previousDay = row.Date.AddDays(-1).Date;
            var armoringWork = Regex.Replace(workName, "бетонир\\w*", "армирование", RegexOptions.IgnoreCase);
            var previousWeather = currentObject.ProductionJournal
                .Where(x => !ReferenceEquals(x, row)
                         && x.Date.Date == previousDay
                         && !string.IsNullOrWhiteSpace(x.Weather))
                .OrderByDescending(x => x.Date)
                .Select(x => x.Weather)
                .FirstOrDefault() ?? row.Weather;

            var exists = currentObject.ProductionJournal.Any(x =>
                !ReferenceEquals(x, row)
                && x.Date.Date == previousDay
                && string.Equals(x.ActionName?.Trim(), row.ActionName?.Trim(), StringComparison.CurrentCultureIgnoreCase)
                && string.Equals(x.WorkName?.Trim(), armoringWork, StringComparison.CurrentCultureIgnoreCase)
                && string.Equals(x.ElementsText?.Trim(), row.ElementsText?.Trim(), StringComparison.CurrentCultureIgnoreCase)
                && string.Equals(x.BlocksText?.Trim(), row.BlocksText?.Trim(), StringComparison.CurrentCultureIgnoreCase)
                && string.Equals(x.MarksText?.Trim(), row.MarksText?.Trim(), StringComparison.CurrentCultureIgnoreCase));

            if (exists)
                return;

            currentObject.ProductionJournal.Add(new ProductionJournalEntry
            {
                Date = previousDay,
                ActionName = row.ActionName,
                WorkName = armoringWork,
                ElementsText = row.ElementsText,
                BlocksText = row.BlocksText,
                MarksText = row.MarksText,
                BrigadeName = row.BrigadeName,
                Weather = previousWeather,
                Deviations = row.Deviations,
                RequiresHiddenWorkAct = row.RequiresHiddenWorkAct
            });
        }

        private double GetMountedQuantityFromProductionJournal(string materialName, ProductionJournalEntry excludeRow = null)
        {
            if (currentObject?.ProductionJournal == null || string.IsNullOrWhiteSpace(materialName))
                return 0;

            return currentObject.ProductionJournal
                .Where(x => !ReferenceEquals(x, excludeRow))
                .SelectMany(x => ParseProductionItems(x.ElementsText))
                .Where(x => string.Equals(x.MaterialName, materialName, StringComparison.CurrentCultureIgnoreCase))
                .Sum(x => x.Quantity);
        }

        private string FindMaterialGroupByName(string materialName)
        {
            if (string.IsNullOrWhiteSpace(materialName))
                return null;

            var fromCatalog = currentObject?.MaterialCatalog?
                .FirstOrDefault(x => string.Equals(x.MaterialName ?? string.Empty, materialName, StringComparison.CurrentCultureIgnoreCase))
                ?.TypeName;

            if (!string.IsNullOrWhiteSpace(fromCatalog))
                return fromCatalog;

            return journal
                .Where(x => x.Category == "Основные"
                    && string.Equals(x.MaterialName ?? string.Empty, materialName, StringComparison.CurrentCultureIgnoreCase))
                .Select(x => x.MaterialGroup)
                .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
        }

        private void RebuildMountedDemandFromProductionJournal()
        {
            if (currentObject?.Demand == null)
                return;

            foreach (var demand in currentObject.Demand.Values)
            {
                demand.MountedLevels = new Dictionary<int, Dictionary<string, double>>();
                demand.MountedFloors = new Dictionary<int, Dictionary<int, double>>();
            }

            if (currentObject.ProductionJournal == null)
                return;

            foreach (var row in currentObject.ProductionJournal)
            {
                var items = ParseProductionItems(row.ElementsText);
                var blocks = LevelMarkHelper.ParseBlocks(row.BlocksText);
                var marks = LevelMarkHelper.ParseMarks(row.MarksText);
                if (items.Count == 0 || blocks.Count == 0 || marks.Count == 0)
                    continue;

                var divisor = blocks.Count * marks.Count;
                if (divisor <= 0)
                    continue;

                foreach (var item in items)
                {
                    var group = FindMaterialGroupByName(item.MaterialName);
                    if (string.IsNullOrWhiteSpace(group))
                        continue;

                    EnsureSummaryMarksForGroup(group, marks);
                    var demandKey = BuildDemandKey(group, item.MaterialName);
                    var demand = GetOrCreateDemand(demandKey, GetUnitForMaterial(group, item.MaterialName));
                    var unit = demand?.Unit;
                    var isDiscrete = IsDiscreteUnit(unit);

                    if (isDiscrete)
                    {
                        var totalDiscrete = (int)NormalizeQuantityByUnit(item.Quantity, unit);
                        if (totalDiscrete <= 0)
                            continue;

                        var basePerCell = totalDiscrete / divisor;
                        var remainder = totalDiscrete % divisor;
                        var cellIndex = 0;

                        foreach (var block in blocks)
                        {
                            if (!demand.MountedLevels.ContainsKey(block))
                                demand.MountedLevels[block] = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

                            foreach (var mark in marks)
                            {
                                var normalizedMark = mark.Trim();
                                if (!demand.MountedLevels[block].ContainsKey(normalizedMark))
                                    demand.MountedLevels[block][normalizedMark] = 0;

                                var value = basePerCell + (cellIndex < remainder ? 1 : 0);
                                demand.MountedLevels[block][normalizedMark] += value;
                                cellIndex++;
                            }
                        }
                    }
                    else
                    {
                        var perCell = item.Quantity / divisor;
                        foreach (var block in blocks)
                        {
                            if (!demand.MountedLevels.ContainsKey(block))
                                demand.MountedLevels[block] = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

                            foreach (var mark in marks)
                            {
                                var normalizedMark = mark.Trim();
                                if (!demand.MountedLevels[block].ContainsKey(normalizedMark))
                                    demand.MountedLevels[block][normalizedMark] = 0;

                                demand.MountedLevels[block][normalizedMark] += perCell;
                            }
                        }
                    }
                }
            }
        }

        private string GetUnitForMaterial(string group, string materialName)
        {
            return journal
                .Where(x => x.Category == "Основные"
                    && string.Equals(x.MaterialGroup ?? string.Empty, group ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals(x.MaterialName ?? string.Empty, materialName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase))
                .Select(x => x.Unit)
                .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? "шт";
        }

        private void EnsureSummaryMarksForGroup(string group, IEnumerable<string> marks)
        {
            if (currentObject == null || string.IsNullOrWhiteSpace(group))
                return;

            currentObject.SummaryMarksByGroup ??= new Dictionary<string, List<string>>();
            if (!currentObject.SummaryMarksByGroup.TryGetValue(group, out var existing) || existing == null)
            {
                existing = LevelMarkHelper.GetMarksForGroup(currentObject, group);
                currentObject.SummaryMarksByGroup[group] = existing;
            }

            foreach (var mark in marks.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()))
            {
                if (!existing.Contains(mark, StringComparer.CurrentCultureIgnoreCase))
                    existing.Add(mark);
            }
        }

        private void RefreshProductionRemainingInfo()
        {
            if (currentObject?.ProductionJournal == null)
                return;

            foreach (var row in currentObject.ProductionJournal)
                row.RemainingInfo = BuildProductionRemainingInfo(row);
        }

        private string BuildProductionRemainingInfo(ProductionJournalEntry row)
        {
            var items = ParseProductionItems(row.ElementsText);
            var blocks = LevelMarkHelper.ParseBlocks(row.BlocksText);
            var marks = LevelMarkHelper.ParseMarks(row.MarksText);
            if (items.Count == 0 || blocks.Count == 0 || marks.Count == 0)
                return string.Empty;

            var parts = new List<string>();
            foreach (var item in items)
            {
                var group = FindMaterialGroupByName(item.MaterialName);
                if (string.IsNullOrWhiteSpace(group))
                    continue;

                var demandKey = BuildDemandKey(group, item.MaterialName);
                var demand = GetOrCreateDemand(demandKey, GetUnitForMaterial(group, item.MaterialName));
                var unit = demand?.Unit ?? GetUnitForMaterial(group, item.MaterialName);
                var plannedTotal = GetTotalDemandOnBuilding(demand);
                if (plannedTotal <= 0)
                    continue;

                foreach (var block in blocks)
                {
                    foreach (var mark in marks)
                    {
                        var planned = GetDemandValue(demand, block, mark);
                        var mounted = GetMountedValue(demand, block, mark);
                        var remaining = NormalizeQuantityByUnit(Math.Max(0, planned - mounted), unit);
                        if (remaining <= 0)
                            continue;

                        parts.Add($"{item.MaterialName}: Б{block} {mark} — остаток {FormatNumberByUnit(remaining, unit)}");
                    }
                }
            }

            return string.Join(Environment.NewLine, parts);
        }

        private static double GetTotalDemandOnBuilding(MaterialDemand demand)
        {
            if (demand?.Levels == null)
                return 0;

            return demand.Levels.Values
                .Where(x => x != null)
                .SelectMany(x => x.Values)
                .Sum();
        }


        // ================= МЕНЮ =================

        private void CreateObject_Click(object sender, RoutedEventArgs e)
        {
            var w = new CreateObjectWindow { Owner = this };
            if (w.ShowDialog() == true)
            {
                currentObject = new ProjectObject
                {
                    Name = w.ObjectName,
                    BlocksCount = 1   // ← КРИТИЧНО
                };

                journal.Clear();
                EnsureOtJournalStorage();
                BindOtJournal();
                ArrivalPanel.SetObject(currentObject, journal);



                EnsureInspectionJournalStorage();
                RefreshInspectionJournalState();
                RefreshDocumentLibraries();
                SaveState();
                RefreshTreePreserveState();
                ApplyProjectUiSettings();
              
            }
        }

        private void ObjectSettings_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            var w = new ObjectSettingsWindow(currentObject)
            {
                Owner = this
            };

            if (w.ShowDialog() == true)
            {
                SaveState();
                RefreshTreePreserveState();
            }
        }

        private void ApplyMaterialBindingChanges(IEnumerable<TreeSettingsWindow.MaterialBindingChange> changes)
        {
            if (changes == null || currentObject == null)
                return;

            foreach (var change in changes)
            {
                foreach (var rec in journal.Where(j =>
                             string.Equals(j.MaterialName ?? string.Empty, change.OldMaterialName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)
                      && string.Equals(j.Category ?? string.Empty, change.OldCategoryName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)
                      && string.Equals(j.MaterialGroup ?? string.Empty, change.OldTypeName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)
                      && string.Equals(j.SubCategory ?? string.Empty, change.OldSubTypeName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)))
                {
                    rec.Category = change.NewCategoryName;
                    rec.MaterialGroup = change.NewTypeName;
                    rec.SubCategory = change.NewSubTypeName;
                    rec.MaterialName = change.NewMaterialName;
                }

                MoveDemandBinding(change);
            }

            SyncLegacyMaterialsFromCatalog();
            RebuildArchiveFromCurrentData();
        }

        private void MoveDemandBinding(TreeSettingsWindow.MaterialBindingChange change)
        {
            if (currentObject?.Demand == null)
                return;

            var oldKey = $"{change.OldTypeName}::{change.OldMaterialName}";
            var newKey = $"{change.NewTypeName}::{change.NewMaterialName}";

            if (string.Equals(oldKey, newKey, StringComparison.CurrentCultureIgnoreCase))
                return;

            if (!currentObject.Demand.TryGetValue(oldKey, out var sourceDemand) || sourceDemand == null)
                return;

            currentObject.Demand.Remove(oldKey);
            if (!currentObject.Demand.TryGetValue(newKey, out var targetDemand) || targetDemand == null)
            {
                currentObject.Demand[newKey] = sourceDemand;
                return;
            }

            MergeDemand(targetDemand, sourceDemand);
        }

        private static void MergeDemand(MaterialDemand target, MaterialDemand source)
        {
            if (target == null || source == null)
                return;

            if (string.IsNullOrWhiteSpace(target.Unit) && !string.IsNullOrWhiteSpace(source.Unit))
                target.Unit = source.Unit;

            MergeLevelMap(target.Levels, source.Levels);
            MergeLevelMap(target.MountedLevels, source.MountedLevels);
            MergeFloorMap(target.Floors, source.Floors);
            MergeFloorMap(target.MountedFloors, source.MountedFloors);
        }

        private static void MergeLevelMap(Dictionary<int, Dictionary<string, double>> target, Dictionary<int, Dictionary<string, double>> source)
        {
            if (target == null || source == null)
                return;

            foreach (var blockPair in source)
            {
                if (!target.TryGetValue(blockPair.Key, out var targetMarks) || targetMarks == null)
                {
                    target[blockPair.Key] = new Dictionary<string, double>(blockPair.Value, StringComparer.CurrentCultureIgnoreCase);
                    continue;
                }

                foreach (var markPair in blockPair.Value)
                {
                    if (!targetMarks.ContainsKey(markPair.Key))
                        targetMarks[markPair.Key] = 0;
                    targetMarks[markPair.Key] += markPair.Value;
                }
            }
        }

        private static void MergeFloorMap(Dictionary<int, Dictionary<int, double>> target, Dictionary<int, Dictionary<int, double>> source)
        {
            if (target == null || source == null)
                return;

            foreach (var blockPair in source)
            {
                if (!target.TryGetValue(blockPair.Key, out var targetFloors) || targetFloors == null)
                {
                    target[blockPair.Key] = new Dictionary<int, double>(blockPair.Value);
                    continue;
                }

                foreach (var floorPair in blockPair.Value)
                {
                    if (!targetFloors.ContainsKey(floorPair.Key))
                        targetFloors[floorPair.Key] = 0;
                    targetFloors[floorPair.Key] += floorPair.Value;
                }
            }
        }

        private void TreeSettings_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }
            currentObject.MaterialCatalog ??= new List<MaterialCatalogItem>();

            var materialNames = journal
                    .Where(j => !string.IsNullOrWhiteSpace(j.MaterialName)
                    && string.Equals(j.Category, "Основные", StringComparison.CurrentCultureIgnoreCase))
                .Select(j => new TreeSettingsWindow.MaterialSplitRuleSource
                {
                    MaterialName = j.MaterialName,

                    CategoryName = j.Category ?? string.Empty,
                    TypeName = j.MaterialGroup ?? string.Empty,
                    SubTypeName = j.SubCategory ?? string.Empty
                })
                                .Concat(currentObject.MaterialCatalog
                    .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName)
                    && string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => new TreeSettingsWindow.MaterialSplitRuleSource
                    {
                        MaterialName = x.MaterialName,
                        CategoryName = x.CategoryName ?? string.Empty,
                        TypeName = x.TypeName ?? string.Empty,
                        SubTypeName = x.SubTypeName ?? string.Empty,
                        Level4Name = x.ExtraLevels != null && x.ExtraLevels.Count > 0 ? x.ExtraLevels[0] : string.Empty,
                        Level5Name = x.ExtraLevels != null && x.ExtraLevels.Count > 1 ? x.ExtraLevels[1] : string.Empty,
                        Level6Name = x.ExtraLevels != null && x.ExtraLevels.Count > 2 ? x.ExtraLevels[2] : string.Empty
                    }))
                .ToList();

            var w = new TreeSettingsWindow(
                materialNames,
                currentObject.MaterialTreeSplitRules ?? new(),
                currentObject.AutoSplitMaterialNames ?? new List<string>())
            {
                Owner = this
            };

            if (w.ShowDialog() != true)
                return;

            currentObject.MaterialTreeSplitRules = w.ResultRules;
            currentObject.AutoSplitMaterialNames = w.ResultAutoSplitMaterials ?? new List<string>();
            currentObject.MaterialCatalog = w.ResultCatalog ?? new List<MaterialCatalogItem>();
            var validMainMaterials = currentObject.MaterialCatalog
    .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase))
    .Select(x => x.MaterialName)
    .Where(x => !string.IsNullOrWhiteSpace(x))
    .ToHashSet(StringComparer.CurrentCultureIgnoreCase);

            journal.RemoveAll(x => x.Category == "Основные" && !validMainMaterials.Contains(x.MaterialName ?? string.Empty));
            currentObject.Demand = currentObject.Demand
                .Where(kv =>
                {
                    var keyParts = kv.Key.Split("::", StringSplitOptions.None);
                    return keyParts.Length == 2 && validMainMaterials.Contains(keyParts[1]);
                })
                .ToDictionary(kv => kv.Key, kv => kv.Value);

            if (w.ResultBindingChanges?.Count > 0)
            {
                PushUndo();
                ApplyMaterialBindingChanges(w.ResultBindingChanges);
                CleanupMaterialsAfterDelete();
            }
            SyncLegacyMaterialsFromCatalog();
            RebuildArchiveFromCurrentData();
            SaveState();
            RefreshTreePreserveState();
            ArrivalPanel.SetObject(currentObject, journal);
            RefreshArrivalTypes();
            RefreshArrivalNames();
            ApplyAllFilters();
        }

        // ================= КНОПКИ =================



        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (!filteredJournal.Any())
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            var win = new ExportModeWindow() { Owner = this };
            if (win.ShowDialog() != true)
                return;

            ExportMode mode = win.Mode;

            var dlg = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = "ЖВК.xlsx"
            };

            if (dlg.ShowDialog() != true)
                return;

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("ЖВК");

                if (mode == ExportMode.Merged)
                    ExportMerged(ws);
                else
                    ExportDetailed(ws);


                wb.SaveAs(dlg.FileName);
            }

            MessageBox.Show("Экспорт завершён");
        }

        string Normalize(string v)
        {
            if (string.IsNullOrWhiteSpace(v))
                return null;

            v = v.Trim();

            // любые пустые формы
            if (v == "—" || v == "-" || v == "--" || v == "_" || v == "null" || v == "None")
                return null;

            return v;
        }



        void ExportMerged(IXLWorksheet ws)
        {
            int row = 1;

            // ===== ЗАГОЛОВОК =====
            ws.Cell(row, 1).Value = "Дата";
            ws.Cell(row, 2).Value = "ТТН";
            ws.Cell(row, 3).Value = "Наименование";
            ws.Cell(row, 4).Value = "СТБ";
            ws.Cell(row, 5).Value = "Ед.";
            ws.Cell(row, 6).Value = "Кол-во";
            ws.Cell(row, 7).Value = "Поставщик";
            ws.Cell(row, 8).Value = "Паспорт";

            ws.Range(row, 1, row, 8).Style.Font.Bold = true;
            ws.Range(row, 1, row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(row, 1, row, 8).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = XLColor.FromHtml("#E9EEF6");
            row++;

            var structured = filteredJournal
                .Where(j => j.Category == "Основные")
                .GroupBy(j => j.Date.Date)
                .OrderByDescending(g => g.Key)
                .ToList();

            foreach (var day in structured)
            {
                int dayStart = row;

                var ttnGroups = day
                    .GroupBy(x => x.Ttn)
                    .ToList();

                foreach (var grp in ttnGroups)
                {
                    var items = grp.ToList();
                    int grpStart = row;
                    int rows = items.Count;

                    // STB
                    string firstStb = Normalize(items[0].Stb);
                    bool stbSame = items.All(x => Normalize(x.Stb) == firstStb);
                    string mergedStb = stbSame ? (firstStb ?? "—") : null;

                    // UNIT
                    string firstUnit = Normalize(items[0].Unit);
                    bool unitSame = items.All(x => Normalize(x.Unit) == firstUnit);
                    string mergedUnit = unitSame ? (firstUnit ?? "—") : null;

                    // SUPPLIER
                    string firstSupplier = Normalize(items[0].Supplier);
                    bool supplierSame = items.All(x => Normalize(x.Supplier) == firstSupplier);
                    string mergedSupplier = supplierSame ? (firstSupplier ?? "—") : null;

                    // выводим строки
                    foreach (var x in items)
                    {
                        ws.Cell(row, 3).Value = Normalize(x.MaterialName) ?? "—";
                        ws.Cell(row, 6).Value = x.Quantity;
                        ws.Cell(row, 8).Value = Normalize(x.Passport) ?? "—";

                        row++;
                    }

                    // merge TTN
                    ws.Range(grpStart, 2, row - 1, 2).Merge();
                    ws.Cell(grpStart, 2).Value = grp.Key;
                    ws.Cell(grpStart, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(grpStart, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    // merge STB
                    if (mergedStb != null)
                    {
                        ws.Range(grpStart, 4, row - 1, 4).Merge();
                        ws.Cell(grpStart, 4).Value = mergedStb;
                    }
                    else
                    {
                        for (int i = 0; i < rows; i++)
                            ws.Cell(grpStart + i, 4).Value = Normalize(items[i].Stb) ?? "—";
                    }

                    // merge UNIT
                    if (mergedUnit != null)
                    {
                        ws.Range(grpStart, 5, row - 1, 5).Merge();
                        ws.Cell(grpStart, 5).Value = mergedUnit;
                    }
                    else
                    {
                        for (int i = 0; i < rows; i++)
                            ws.Cell(grpStart + i, 5).Value = Normalize(items[i].Unit) ?? "—";

                    }

                    // merge SUPPLIER
                    if (mergedSupplier != null)
                    {
                        ws.Range(grpStart, 7, row - 1, 7).Merge();
                        ws.Cell(grpStart, 7).Value = mergedSupplier;
                        ws.Cell(grpStart, 7).Style.Alignment.WrapText = true;
                    }
                    else
                    {
                        for (int i = 0; i < rows; i++)
                            ws.Cell(grpStart + i, 7).Value = Normalize(items[i].Supplier) ?? "—";

                    }

                    // заливка всего блока
                    var c = GetSoftColor(grp.Key);
                    var fill = XLColor.FromColor(System.Drawing.Color.FromArgb(55, c.R, c.G, c.B));
                    ws.Range(grpStart, 2, row - 1, 8).Style.Fill.BackgroundColor = fill;

                    // рамка блока
                    ws.Range(grpStart, 2, row - 1, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // merge DATE
                ws.Range(dayStart, 1, row - 1, 1).Merge();
                ws.Cell(dayStart, 1).Value = day.Key.ToString("dd.MM.yyyy");
                ws.Cell(dayStart, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(dayStart, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Range(dayStart, 1, row - 1, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            }

            ws.Columns().AdjustToContents();
            ws.Range(1, 1, row - 1, 8).SetAutoFilter();
        }







        void ExportDetailed(IXLWorksheet ws)
        {
            int row = 1;

            // Заголовок
            ws.Cell(row, 1).Value = "Дата";
            ws.Cell(row, 2).Value = "ТТН";
            ws.Cell(row, 3).Value = "Наименование";
            ws.Cell(row, 4).Value = "СТБ";
            ws.Cell(row, 5).Value = "Ед.";
            ws.Cell(row, 6).Value = "Кол-во";
            ws.Cell(row, 7).Value = "Поставщик";
            ws.Cell(row, 8).Value = "Паспорт";

            ws.Range(row, 1, row, 8).Style.Font.Bold = true;
            ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = XLColor.FromHtml("#E9EEF6");
            row++;

            var days = filteredJournal
                .Where(j => j.Category == "Основные")
                .GroupBy(j => j.Date.Date)
                .OrderByDescending(g => g.Key);

            foreach (var day in days)
            {
                int dayStart = row;

                var dayGroups = day.GroupBy(x => x.Ttn);

                foreach (var grp in dayGroups)
                {
                    var items = grp.ToList();
                    int grpStart = row;
                    int rows = items.Count;

                    // === анализ одинаковости ===
                    string firstStb = Normalize(items[0].Stb);
                    bool stbSame = items.All(x => Normalize(x.Stb) == firstStb);
                    string mergedStb = stbSame ? (firstStb ?? "—") : null;

                    string firstUnit = Normalize(items[0].Unit);
                    bool unitSame = items.All(x => Normalize(x.Unit) == firstUnit);
                    string mergedUnit = unitSame ? (firstUnit ?? "—") : null;

                    string firstSupplier = Normalize(items[0].Supplier);
                    bool supplierSame = items.All(x => Normalize(x.Supplier) == firstSupplier);
                    string mergedSupplier = supplierSame ? (firstSupplier ?? "—") : null;


                    // === строки ===
                    foreach (var x in items)
                    {
                        ws.Cell(row, 2).Value = x.Ttn;
                        ws.Cell(row, 3).Value = x.MaterialName;
                        ws.Cell(row, 6).Value = x.Quantity;
                        ws.Cell(row, 8).Value = x.Passport ?? "—";
                        ws.Cell(row, 4).Value = Normalize(x.Stb) ?? "—";
                        ws.Cell(row, 5).Value = Normalize(x.Unit) ?? "—";
                        ws.Cell(row, 7).Value = Normalize(x.Supplier) ?? "—";


                        var c = GetSoftColor(x.Ttn);
                        var draw = System.Drawing.Color.FromArgb(35, c.R, c.G, c.B);
                        ws.Range(row, 2, row, 8).Style.Fill.BackgroundColor = XLColor.FromColor(draw);

                        row++;
                    }

                    // === STB ===
                    if (mergedStb != null)
                    {
                        ws.Range(grpStart, 4, row - 1, 4).Merge();
                        ws.Cell(grpStart, 4).Value = mergedStb;
                    }
                    else
                    {
                        for (int i = 0; i < rows; i++)
                            ws.Cell(grpStart + i, 4).Value = Normalize(items[i].Stb) ?? "—";
                    }

                    // === UNIT ===
                    if (mergedUnit != null)
                    {
                        ws.Range(grpStart, 5, row - 1, 5).Merge();
                        ws.Cell(grpStart, 5).Value = mergedUnit;
                    }
                    else
                    {
                        for (int i = 0; i < rows; i++)
                            ws.Cell(grpStart + i, 5).Value = Normalize(items[i].Unit) ?? "—";
                    }

                    // === SUPPLIER ===
                    if (mergedSupplier != null)
                    {
                        ws.Range(grpStart, 7, row - 1, 7).Merge();
                        ws.Cell(grpStart, 7).Value = mergedSupplier;
                        ws.Cell(grpStart, 7).Style.Alignment.WrapText = true;
                    }
                    else
                    {
                        for (int i = 0; i < rows; i++)
                            ws.Cell(grpStart + i, 7).Value = Normalize(items[i].Supplier) ?? "—";
                    }


                }

                ws.Range(dayStart, 1, row - 1, 1).Merge();
                ws.Cell(dayStart, 1).Value = day.Key.ToString("dd.MM.yyyy");
                ws.Range(dayStart, 1, row - 1, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            }


            // автоподгон
            ws.Columns().AdjustToContents();
            ws.Rows().AdjustToContents();
            ws.Range(1, 1, row - 1, 8).SetAutoFilter();
        }




        private void LockButton_Checked(object sender, RoutedEventArgs e)
        {
            isLocked = true;

            RefreshSummaryTable();
        }


        private void LockButton_Unchecked(object sender, RoutedEventArgs e)
        {
            isLocked = false;

            RefreshSummaryTable();
        }


        private void ArrivalFilterButton_Click(object sender, RoutedEventArgs e)
        {
            ArrivalFiltersOverlay.Visibility = Visibility.Visible;
        }

        private void CloseArrivalFilters_Click(object sender, RoutedEventArgs e)
        {
            ArrivalFiltersOverlay.Visibility = Visibility.Collapsed;
        }

        // ================= ПРИХОД =================

        private void OnArrivalAdded(Arrival arrival)
        {
            PushUndo(); // ⬅️ ВОТ ЭТОГО НЕ ХВАТАЛО


            currentObject.MaterialCatalog ??= new List<MaterialCatalogItem>();

            foreach (var i in arrival.Items)
            {
                var rowGroup = i.MaterialGroup?.Trim() ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(i.MaterialName))
                {
                    var categoryName = arrival.Category ?? string.Empty;
                    var typeName = arrival.Category == "Основные" ? rowGroup : (arrival.SubCategory ?? string.Empty);
                    var subTypeName = arrival.Category == "Основные" ? string.Empty : rowGroup;

                      if (!currentObject.MaterialCatalog.Any(x =>
                          string.Equals(x.CategoryName ?? string.Empty, categoryName, StringComparison.CurrentCultureIgnoreCase)
                        && string.Equals(x.TypeName ?? string.Empty, typeName, StringComparison.CurrentCultureIgnoreCase)
                        && string.Equals(x.SubTypeName ?? string.Empty, subTypeName, StringComparison.CurrentCultureIgnoreCase)
                        && string.Equals(x.MaterialName ?? string.Empty, i.MaterialName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        currentObject.MaterialCatalog.Add(new MaterialCatalogItem
                        {
                            CategoryName = categoryName,
                            TypeName = typeName,
                            SubTypeName = subTypeName,
                            MaterialName = i.MaterialName
                        });
                    }
                }

                  if (arrival.Category == "Основные")
                  {
                      if (!currentObject.MaterialGroups.Any(g => g.Name == rowGroup))
                          currentObject.MaterialGroups.Add(new MaterialGroup { Name = rowGroup });

                      if (!currentObject.MaterialNamesByGroup.ContainsKey(rowGroup))
                          currentObject.MaterialNamesByGroup[rowGroup] = new List<string>();

                      // === список на дереве ===
                      if (!currentObject.MaterialNamesByGroup[rowGroup]
                              .Contains(i.MaterialName))
                      {
                          currentObject.MaterialNamesByGroup[rowGroup]
                              .Add(i.MaterialName);
                      }

                      // === список для ComboBox ===
                      var archive = currentObject.Archive;

                      if (!archive.Groups.Contains(rowGroup))
                          archive.Groups.Add(rowGroup);

                      if (!archive.Materials.ContainsKey(rowGroup))
                          archive.Materials[rowGroup] = new();

                      if (!archive.Materials[rowGroup]
                              .Contains(i.MaterialName))
                      {
                          archive.Materials[rowGroup].Add(i.MaterialName);
                      }
                  }

                // === запись журнала ===
                journal.Add(new JournalRecord
                {
                      Date = i.Date,
                      ObjectName = currentObject.Name,
                      Category = arrival.Category,
                      SubCategory = arrival.SubCategory,
                      MaterialGroup = rowGroup,
                      MaterialName = i.MaterialName,
                    Unit = i.Unit,
                    Quantity = i.Quantity,
                    Passport = i.Passport,
                    Ttn = arrival.TtnNumber,
                    Stb = i.Stb,
                    Supplier = i.Supplier
                });
            }



            SyncLegacyMaterialsFromCatalog();
            SaveState();
            RefreshTreePreserveState();
   

            // важно: обновляем панель прихода
            ArrivalPanel.SetObject(currentObject, journal);
            // === обновляем чипы типов и материалов ===
            RefreshArrivalTypes();
            RefreshArrivalNames();


        }

        private void SyncLegacyMaterialsFromCatalog()
        {
            if (currentObject == null)
                return;

            currentObject.MaterialCatalog ??= new List<MaterialCatalogItem>();

            foreach (var item in currentObject.MaterialCatalog.Where(x => !string.IsNullOrWhiteSpace(x.MaterialName)
               && string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)))
            {
                if (!string.Equals(item.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase))
                    continue;

                var groupName = item.TypeName ?? string.Empty;
                if (string.IsNullOrWhiteSpace(groupName))
                    continue;

                if (!currentObject.MaterialGroups.Any(g => string.Equals(g.Name, groupName, StringComparison.CurrentCultureIgnoreCase)))
                    currentObject.MaterialGroups.Add(new MaterialGroup { Name = groupName });

                if (!currentObject.MaterialNamesByGroup.ContainsKey(groupName))
                    currentObject.MaterialNamesByGroup[groupName] = new List<string>();

                if (!currentObject.MaterialNamesByGroup[groupName].Contains(item.MaterialName))
                    currentObject.MaterialNamesByGroup[groupName].Add(item.MaterialName);

                if (!currentObject.Archive.Groups.Contains(groupName))
                    currentObject.Archive.Groups.Add(groupName);

                if (!currentObject.Archive.Materials.ContainsKey(groupName))
                    currentObject.Archive.Materials[groupName] = new List<string>();

                if (!currentObject.Archive.Materials[groupName].Contains(item.MaterialName))
                    currentObject.Archive.Materials[groupName].Add(item.MaterialName);
            }
        }
        private void RebuildArchiveFromCurrentData()
        {
            if (currentObject == null)
                return;

            var archive = new ObjectArchive();

            foreach (var rec in journal)
            {
                if (!string.IsNullOrWhiteSpace(rec.MaterialGroup))
                {
                    if (!archive.Groups.Contains(rec.MaterialGroup))
                        archive.Groups.Add(rec.MaterialGroup);

                    if (!archive.Materials.ContainsKey(rec.MaterialGroup))
                        archive.Materials[rec.MaterialGroup] = new List<string>();

                    if (!string.IsNullOrWhiteSpace(rec.MaterialName) && !archive.Materials[rec.MaterialGroup].Contains(rec.MaterialName))
                        archive.Materials[rec.MaterialGroup].Add(rec.MaterialName);
                }

                if (!string.IsNullOrWhiteSpace(rec.Unit) && !archive.Units.Contains(rec.Unit))
                    archive.Units.Add(rec.Unit);

                if (!string.IsNullOrWhiteSpace(rec.Supplier) && !archive.Suppliers.Contains(rec.Supplier))
                    archive.Suppliers.Add(rec.Supplier);

                if (!string.IsNullOrWhiteSpace(rec.Passport) && !archive.Passports.Contains(rec.Passport))
                    archive.Passports.Add(rec.Passport);

                if (!string.IsNullOrWhiteSpace(rec.Stb) && !archive.Stb.Contains(rec.Stb))
                    archive.Stb.Add(rec.Stb);
            }

            currentObject.Archive = archive;
        }
        private void CleanupMaterialsAfterDelete()
        {
            // Какие группы реально используются
            var usedGroups = journal
                .Select(j => j.MaterialGroup)
                .Distinct()
                .ToHashSet();

            // 1. Удаляем пустые группы
            currentObject.MaterialGroups
                .RemoveAll(g => !usedGroups.Contains(g.Name));

            // 2. Удаляем пустые материалы
            foreach (var g in currentObject.MaterialNamesByGroup.Keys.ToList())
            {
                var usedMaterials = journal
                    .Where(j => j.MaterialGroup == g)
                    .Select(j => j.MaterialName)
                    .Distinct()
                    .ToHashSet();

                currentObject.MaterialNamesByGroup[g]
                    .RemoveAll(m => !usedMaterials.Contains(m));

                // если в группе вообще ничего не осталось
                if (currentObject.MaterialNamesByGroup[g].Count == 0)
                    currentObject.MaterialNamesByGroup.Remove(g);
            }
        }

        // ================= ДЕРЕВО =================

        private void RefreshTreePreserveState()
        {
            ObjectsTree.Items.Clear();
            if (currentObject == null)
                return;
            currentObject.MaterialTreeSplitRules ??= new Dictionary<string, string>();

            var newRoot = new TreeViewItem
            {
                Header = currentObject.Name,
                Tag = "Object",
                IsExpanded = true
            };

            var mainNode = new TreeViewItem
            {
                Header = "Основные",
                Tag = "Category",
                IsExpanded = true
            };

            var extraNode = new TreeViewItem
            {
                Header = "Допы",
                Tag = "Category",
                IsExpanded = true
            };

            // ===== ОСНОВНЫЕ =====
            var mainGroups = journal
                .Where(j => j.Category == "Основные")
                .GroupBy(j => j.MaterialGroup)
                .OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase);

            foreach (var g in mainGroups)
            {
                var groupNode = new TreeViewItem
                {
                    Header = g.Key,
                    Tag = "Group",
                    IsExpanded = false
                };

                var materialsBySubType = g
                      .Select(x => x.MaterialName)
                      .Where(x => !string.IsNullOrWhiteSpace(x))
                      .Distinct(StringComparer.CurrentCultureIgnoreCase)
                      .GroupBy(material => GetMainMaterialSubType(g.Key, material) ?? string.Empty)
                      .OrderBy(x => string.IsNullOrWhiteSpace(x.Key) ? "~~~~" : x.Key, StringComparer.CurrentCultureIgnoreCase)
                      .ToList();

                foreach (var subTypeGroup in materialsBySubType)
                {
                    var targetParent = groupNode;

                    if (!string.IsNullOrWhiteSpace(subTypeGroup.Key))
                    {
                        targetParent = new TreeViewItem
                        {
                            Header = subTypeGroup.Key,
                            Tag = "SubType",
                            IsExpanded = false
                        };
                        groupNode.Items.Add(targetParent);
                    }

                    foreach (var m in subTypeGroup.OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
                        AddMaterialTreeNodes(targetParent, m, g.Key, "Основные", subTypeGroup.Key);
                }

                mainNode.Items.Add(groupNode);
            }

            // ===== ДОПЫ =====
            var extraGroups = journal
                .Where(j => j.Category == "Допы")
                 .GroupBy(j => j.SubCategory)
                .OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase);

            foreach (var g in extraGroups)
            {
                var subNode = new TreeViewItem
                {
                    Header = g.Key,
                    Tag = "SubCategory",
                    IsExpanded = false
                };

                foreach (var m in g.Select(x => x.MaterialName)
                              .Distinct()
                              .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
                {
                    AddMaterialTreeNodes(subNode, m, null, "Допы", g.Key);
                }

                extraNode.Items.Add(subNode);
            }

            newRoot.Items.Add(mainNode);
            newRoot.Items.Add(extraNode);

            ObjectsTree.Items.Add(newRoot);
        }
        private string GetMainMaterialSubType(string groupName, string materialName)
        {
            if (currentObject?.MaterialCatalog == null)
                return string.Empty;

            return currentObject.MaterialCatalog
                .Where(x => string.Equals(x.CategoryName ?? string.Empty, "Основные", StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals(x.TypeName ?? string.Empty, groupName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals(x.MaterialName ?? string.Empty, materialName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase))
                .Select(x => x.SubTypeName ?? string.Empty)
                .FirstOrDefault() ?? string.Empty;
        }

        private void AddMaterialTreeNodes(TreeViewItem parent, string materialName, string groupName, string category, string subCategory)
        {
            var segments = GetSegmentsForMaterial(materialName);
            if (segments.Count == 0)
                segments.Add(materialName);

            ItemsControl current = parent;
            var prefix = new List<string>();

            for (int i = 0; i < segments.Count; i++)
            {
                var isFinal = i == segments.Count - 1;
                prefix.Add(segments[i]);

                if (!isFinal)
                {
                    var existingNode = FindChildNode(current, segments[i]);
                    if (existingNode != null)
                    {
                        current = existingNode;
                        continue;
                    }
                }

                var node = new TreeViewItem
                {
                    Header = segments[i],
                    Tag = new TreeNodeMeta
                    {
                        Kind = isFinal ? "Material" : "MaterialPart",
                        MaterialName = isFinal ? materialName : null,
                        GroupName = groupName,
                        Category = category,
                        SubCategory = subCategory,
                        PrefixSegments = prefix.ToList()
                    },
                    IsExpanded = false
                };

                current.Items.Add(node);
                current = node;
            }
        }

        private TreeViewItem FindChildNode(ItemsControl parent, string header)
        {
            foreach (var child in parent.Items)
            {
                if (child is TreeViewItem node
                    && string.Equals(node.Header?.ToString(), header, StringComparison.CurrentCultureIgnoreCase))
                    return node;
            }

            return null;
        }

        public static List<string> GetSegmentsFromText(string materialName)
        {
            if (string.IsNullOrWhiteSpace(materialName))
                return new List<string>();

            return Regex.Matches(materialName, @"[A-Za-zА-Яа-яЁё]+|\d+(?:[\.,]\d+)?")
                 .Select(m => m.Value)
                .ToList();
        }

        private List<string> GetSegmentsForMaterial(string materialName)
        {
            if (string.IsNullOrWhiteSpace(materialName))
                return new List<string>();

            if (currentObject?.MaterialTreeSplitRules != null
                && currentObject.MaterialTreeSplitRules.TryGetValue(materialName, out var rule)
                && !string.IsNullOrWhiteSpace(rule))
            {
                return rule
                    .Split('|', StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToList();
            }

            if (currentObject?.AutoSplitMaterialNames != null
                && currentObject.AutoSplitMaterialNames.Contains(materialName, StringComparer.CurrentCultureIgnoreCase))
            {
                var autoSegments = Regex.Matches(materialName, @"[A-Za-zА-Яа-яЁё]+|\d+")
                    .Select(x => x.Value.Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToList();

                if (autoSegments.Count > 0)
                    return autoSegments;
            }

            return new List<string> { materialName };
        }
        private void ObjectsTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            SyncArrivalMatrixSelectionWithTree();
            ApplyAllFilters();
        }

        private void SyncArrivalMatrixSelectionWithTree()
        {
            if (!arrivalMatrixMode || ObjectsTree?.SelectedItem is not TreeViewItem node)
                return;

            string group = null;
            foreach (var currentNode in EnumerateNodeWithParents(node))
            {
                if (GetNodeKind(currentNode) == "Group")
                {
                    group = currentNode.Header?.ToString();
                    break;
                }

                if (currentNode.Tag is TreeNodeMeta meta && !string.IsNullOrWhiteSpace(meta.GroupName))
                {
                    group = meta.GroupName;
                    break;
                }
            }

            if (string.IsNullOrWhiteSpace(group))
                return;

            if (selectedArrivalTypes.Count == 1 && selectedArrivalTypes.Contains(group))
                return;

            selectedArrivalTypes.Clear();
            selectedArrivalTypes.Add(group);
            selectedArrivalNames.Clear();
            RefreshArrivalTypes();
            RefreshArrivalNames();
        }

        private string GetNodeKind(TreeViewItem node)
        {
            if (node.Tag is TreeNodeMeta meta)
                return meta.Kind;

            return node.Tag as string;
        }

        private TreeViewItem FindParentNode(DependencyObject child)
        {
            var parent = VisualTreeHelper.GetParent(child);

            while (parent != null)
            {
                if (parent is TreeViewItem tvi)
                    return tvi;

                parent = VisualTreeHelper.GetParent(parent);
            }

            return null;
        }
        private IEnumerable<TreeViewItem> EnumerateNodeWithParents(TreeViewItem node)
        {
            var current = node;
            while (current != null)
            {
                yield return current;
                current = FindParentNode(current);
            }
        }

        private void ObjectsTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            treeDragStart = e.GetPosition(null);
        }

        private void ObjectsTree_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || isLocked)
                return;

            var position = e.GetPosition(null);
            if (Math.Abs(position.X - treeDragStart.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(position.Y - treeDragStart.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            if (ObjectsTree.SelectedItem is not TreeViewItem selected)
                return;

            var kind = GetNodeKind(selected);
            if (kind != "Group" && kind != "Material")
                return;

            DragDrop.DoDragDrop(selected, selected, DragDropEffects.Move);
        }

        private void ObjectsTree_Drop(object sender, DragEventArgs e)
        {
            if (isLocked)
                return;

            if (!e.Data.GetDataPresent(typeof(TreeViewItem)))
                return;

            if (e.Data.GetData(typeof(TreeViewItem)) is not TreeViewItem sourceNode)
                return;

            if (e.OriginalSource is not DependencyObject dep)
                return;

            var targetNode = FindParentNode(dep);
            if (targetNode == null || ReferenceEquals(sourceNode, targetNode))
                return;

            var sourceKind = GetNodeKind(sourceNode);
            var targetKind = GetNodeKind(targetNode);

           

            if (sourceKind == "Material")
            {
                if (sourceNode.Tag is not TreeNodeMeta sourceMeta)
                    return;

               
                if (targetKind == "Group")
                {
                    var targetGroup = targetNode.Header?.ToString();
                    if (string.IsNullOrWhiteSpace(targetGroup) || targetGroup == sourceMeta.GroupName)
                        return;

                    PushUndo();

                    foreach (var rec in journal.Where(j => j.MaterialName == sourceMeta.MaterialName && j.MaterialGroup == sourceMeta.GroupName))
                        rec.MaterialGroup = targetGroup;

                    foreach (var item in currentObject.MaterialCatalog.Where(x =>
                                 string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                              && string.Equals(x.TypeName ?? string.Empty, sourceMeta.GroupName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)
                              && string.Equals(x.MaterialName ?? string.Empty, sourceMeta.MaterialName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        item.TypeName = targetGroup;
                    }

                    MoveDemandBinding(new TreeSettingsWindow.MaterialBindingChange
                    {
                        OldCategoryName = "Основные",
                        NewCategoryName = "Основные",
                        OldTypeName = sourceMeta.GroupName,
                        NewTypeName = targetGroup,
                        OldSubTypeName = sourceMeta.SubCategory,
                        NewSubTypeName = sourceMeta.SubCategory,
                        OldMaterialName = sourceMeta.MaterialName,
                        NewMaterialName = sourceMeta.MaterialName
                    });

                    
                        CleanupMaterialsAfterDelete();
                    SyncLegacyMaterialsFromCatalog();
                    RebuildArchiveFromCurrentData();
                    SaveState();
                    RefreshTreePreserveState();
                    ApplyAllFilters();
                    return;
                }

                
                if (targetKind == "Material" || targetKind == "MaterialPart")
                {
                    if (targetNode.Tag is not TreeNodeMeta targetMeta || targetMeta.PrefixSegments == null || targetMeta.PrefixSegments.Count == 0)
                        return;

                    
                    var sourceSegments = GetSegmentsForMaterial(sourceMeta.MaterialName);
                    var sourceLeaf = sourceSegments.LastOrDefault() ?? sourceMeta.MaterialName;

                    var targetPrefix = targetKind == "Material"
                        ? targetMeta.PrefixSegments.Take(targetMeta.PrefixSegments.Count - 1).ToList()
                        : targetMeta.PrefixSegments.ToList();

                    if (targetPrefix.Count == 0)
                        return;

                    var newSegments = targetPrefix.Concat(new[] { sourceLeaf }).ToList();
                    var newRule = string.Join("|", newSegments);
                    var oldRule = string.Join("|", sourceSegments);

                    if (string.Equals(newRule, oldRule, StringComparison.CurrentCultureIgnoreCase))
                        return;

                    PushUndo();
                    currentObject.MaterialTreeSplitRules ??= new Dictionary<string, string>();
                    currentObject.MaterialTreeSplitRules[sourceMeta.MaterialName] = newRule;

                    SyncLegacyMaterialsFromCatalog();
                    SaveState();
                    RefreshTreePreserveState();
                    ApplyAllFilters();
                    return;
                }

                return;
            }
            

                if (sourceKind == "Group" && targetKind == "Group")
                {
                    var sourceName = sourceNode.Header?.ToString();
                    var targetName = targetNode.Header?.ToString();

                    if (string.IsNullOrWhiteSpace(sourceName) || string.IsNullOrWhiteSpace(targetName) || sourceName == targetName)
                        return;

                    PushUndo();

                    foreach (var rec in journal.Where(j => j.MaterialGroup == sourceName))
                        rec.MaterialGroup = targetName;

                    foreach (var item in currentObject.MaterialCatalog.Where(x =>
                                 string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                              && string.Equals(x.TypeName ?? string.Empty, sourceName ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        item.TypeName = targetName;
                    }

                    foreach (var key in currentObject.Demand.Keys
                                 .Where(x => x.StartsWith(sourceName + "::", StringComparison.CurrentCultureIgnoreCase))
                                 .ToList())
                    {
                        var materialName = key.Split(new[] { "::" }, StringSplitOptions.None).Skip(1).FirstOrDefault() ?? string.Empty;
                        MoveDemandBinding(new TreeSettingsWindow.MaterialBindingChange
                        {
                            OldCategoryName = "Основные",
                            NewCategoryName = "Основные",
                            OldTypeName = sourceName,
                            NewTypeName = targetName,
                            OldSubTypeName = string.Empty,
                            NewSubTypeName = string.Empty,
                            OldMaterialName = materialName,
                            NewMaterialName = materialName
                        });
                    }

                    CleanupMaterialsAfterDelete();
                    SyncLegacyMaterialsFromCatalog();
                    RebuildArchiveFromCurrentData();
                    SaveState();
                    RefreshTreePreserveState();
                    ApplyAllFilters();
                }

        }


        // ================= ПКМ =================

        private void RenameTreeItem_Click(object sender, RoutedEventArgs e)
        {
            if (isLocked)
            {
                MessageBox.Show("Редактирование заблокировано");
                return;
            }

            if (ObjectsTree.SelectedItem is not TreeViewItem node)
                return;

            if (GetNodeKind(node) == "Object")
                return;

            var oldName = node.Header.ToString();

            var input = Microsoft.VisualBasic.Interaction.InputBox(
                "Новое название:",
                "Переименование",
                oldName);

            if (string.IsNullOrWhiteSpace(input) || input == oldName)
                return;
            PushUndo(); // ⬅️ ВАЖНО: сохраняем состояние ДО переименования

            if (GetNodeKind(node) == "Group")
            {
                var g = currentObject.MaterialGroups.First(x => x.Name == oldName);
                g.Name = input;

                if (currentObject.MaterialNamesByGroup.ContainsKey(oldName))
                {
                    currentObject.MaterialNamesByGroup[input] =
                        currentObject.MaterialNamesByGroup[oldName];
                    currentObject.MaterialNamesByGroup.Remove(oldName);
                }

                foreach (var j in journal.Where(x => x.MaterialGroup == oldName))
                    j.MaterialGroup = input;

                foreach (var item in currentObject.MaterialCatalog.Where(x =>
                             string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                          && string.Equals(x.TypeName ?? string.Empty, oldName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    item.TypeName = input;
                }

                foreach (var key in currentObject.Demand.Keys
                             .Where(x => x.StartsWith(oldName + "::", StringComparison.CurrentCultureIgnoreCase))
                             .ToList())
                {
                    var materialName = key.Split(new[] { "::" }, StringSplitOptions.None).Skip(1).FirstOrDefault() ?? string.Empty;
                    MoveDemandBinding(new TreeSettingsWindow.MaterialBindingChange
                    {
                        OldCategoryName = "Основные",
                        NewCategoryName = "Основные",
                        OldTypeName = oldName,
                        NewTypeName = input,
                        OldSubTypeName = string.Empty,
                        NewSubTypeName = string.Empty,
                        OldMaterialName = materialName,
                        NewMaterialName = materialName
                    });
                }
            }

            if (GetNodeKind(node) == "Material")
            {
                var oldMaterialName = node.Tag is TreeNodeMeta meta ? meta.MaterialName : oldName;
                var wasAutoSplit = currentObject.AutoSplitMaterialNames.Any(x => string.Equals(x, oldMaterialName, StringComparison.CurrentCultureIgnoreCase));
                foreach (var kv in currentObject.MaterialNamesByGroup)
                {
                    var idx = kv.Value.IndexOf(oldMaterialName);
                    if (idx >= 0)
                        kv.Value[idx] = input;
                }

                foreach (var j in journal.Where(x => x.MaterialName == oldMaterialName))
                    j.MaterialName = input;

                foreach (var item in currentObject.MaterialCatalog.Where(x =>
                             string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                          && string.Equals(x.MaterialName ?? string.Empty, oldMaterialName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    item.MaterialName = input;
                }

                if (currentObject.MaterialTreeSplitRules.TryGetValue(oldMaterialName, out var rule))
                {
                    if (!wasAutoSplit)
                        currentObject.MaterialTreeSplitRules[input] = rule;
                    currentObject.MaterialTreeSplitRules.Remove(oldMaterialName);
                }

                if (currentObject.AutoSplitMaterialNames.RemoveAll(x => string.Equals(x, oldMaterialName, StringComparison.CurrentCultureIgnoreCase)) > 0
                    && !currentObject.AutoSplitMaterialNames.Contains(input, StringComparer.CurrentCultureIgnoreCase))
                {
                    currentObject.AutoSplitMaterialNames.Add(input);
                }

                foreach (var key in currentObject.Demand.Keys
                             .Where(x => x.EndsWith("::" + oldMaterialName, StringComparison.CurrentCultureIgnoreCase))
                             .ToList())
                {
                    var groupName = key.Split(new[] { "::" }, StringSplitOptions.None).FirstOrDefault() ?? string.Empty;
                    MoveDemandBinding(new TreeSettingsWindow.MaterialBindingChange
                    {
                        OldCategoryName = "Основные",
                        NewCategoryName = "Основные",
                        OldTypeName = groupName,
                        NewTypeName = groupName,
                        OldSubTypeName = string.Empty,
                        NewSubTypeName = string.Empty,
                        OldMaterialName = oldMaterialName,
                        NewMaterialName = input
                    });
                }
            }

            CleanupMaterialsAfterDelete();
            SyncLegacyMaterialsFromCatalog();
            RebuildArchiveFromCurrentData();
            SaveState();
            RefreshTreePreserveState();
           
        }

        private void DeleteTreeItem_Click(object sender, RoutedEventArgs e)
        {
            if (isLocked)
            {
                MessageBox.Show("Редактирование заблокировано");
                return;
            }

            if (ObjectsTree.SelectedItem is not TreeViewItem node)
                return;

            if (GetNodeKind(node) == "Object")
                return;

            var name = node.Header.ToString();

            if (MessageBox.Show($"Удалить \"{name}\"?",
                "Подтверждение",
                MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;

            if (GetNodeKind(node) == "Group")
            {
                currentObject.MaterialGroups.RemoveAll(g => g.Name == name);
                currentObject.MaterialNamesByGroup.Remove(name);
                journal.RemoveAll(j => j.MaterialGroup == name);
                currentObject.MaterialCatalog.RemoveAll(x =>
                    string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals(x.TypeName ?? string.Empty, name, StringComparison.CurrentCultureIgnoreCase));

                foreach (var key in currentObject.Demand.Keys
                             .Where(x => x.StartsWith(name + "::", StringComparison.CurrentCultureIgnoreCase))
                             .ToList())
                {
                    currentObject.Demand.Remove(key);
                }
            }

            if (GetNodeKind(node) == "Material")
            {
                var materialName = node.Tag is TreeNodeMeta meta ? meta.MaterialName : name;
                foreach (var kv in currentObject.MaterialNamesByGroup)
                kv.Value.Remove(materialName);

                journal.RemoveAll(j => j.MaterialName == materialName);
                currentObject.MaterialTreeSplitRules.Remove(materialName);
                currentObject.AutoSplitMaterialNames.RemoveAll(x => string.Equals(x, materialName, StringComparison.CurrentCultureIgnoreCase));
                currentObject.MaterialCatalog.RemoveAll(x =>
                    string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals(x.MaterialName ?? string.Empty, materialName, StringComparison.CurrentCultureIgnoreCase));

                foreach (var key in currentObject.Demand.Keys
                             .Where(x => x.EndsWith("::" + materialName, StringComparison.CurrentCultureIgnoreCase))
                             .ToList())
                {
                    currentObject.Demand.Remove(key);
                }
            }

            CleanupMaterialsAfterDelete();
            RebuildArchiveFromCurrentData();
            SaveState();
            RefreshTreePreserveState();
          
        }

        private void ArrivalFilters_Changed(object sender, RoutedEventArgs e)
        {
            ApplyAllFilters();
        }



        private void ApplyAllFilters()
        {
            IEnumerable<JournalRecord> data = journal;
            // === ПРИХОД: КАТЕГОРИЯ ОСНОВНЫЕ/ДОПЫ ===
            bool showMain = ArrivalMainCheck?.IsChecked == true;
            bool showExtra = ArrivalExtraCheck?.IsChecked == true;

            data = data.Where(j =>
                (showMain && j.Category == "Основные")
                || (showExtra && j.Category == "Допы")
            );



            // ===== ДОПОЛНИТЕЛЬНЫЕ ФИЛЬТРЫ =====
            // ===== ДОПОЛНИТЕЛЬНЫЕ ФИЛЬТРЫ (ДОПЫ ПО УМОЛЧАНИЮ СКРЫТЫ) =====
            // === ПРИХОД: ДОПОЛНИТЕЛЬНЫЕ ПОДТИПЫ ===
            bool showLowCost = ArrivalLowCostCheck?.IsChecked == true;
            bool showInternal = ArrivalInternalCheck?.IsChecked == true;

            data = data.Where(j =>
                j.Category != "Допы"
                || (
                    (showLowCost && j.SubCategory == "Малоценка")
                    || (showInternal && j.SubCategory == "Внутренние")
                )
            );
            // === ПРИХОД: ФИЛЬТР ПО ТИПАМ ===
            if (selectedArrivalTypes.Count > 0)
            {
                data = data.Where(j => selectedArrivalTypes.Contains(j.MaterialGroup));
            }
            // === ПРИХОД: ФИЛЬТР ПО НАИМЕНОВАНИЯМ ===
            if (selectedArrivalNames.Count > 0)
            {
                data = data.Where(j => selectedArrivalNames.Contains(j.MaterialName));
            }




            if (ObjectsTree.SelectedItem is TreeViewItem node)
            {
                foreach (var currentNode in EnumerateNodeWithParents(node))
                {

                    var kind = GetNodeKind(currentNode);
                    var value = currentNode.Header?.ToString();

                    if (kind == "Group")
                        data = data.Where(j => j.MaterialGroup == value);
                    else if (kind == "SubCategory")
                        data = data.Where(j => j.SubCategory == value);
                    else if (kind == "Category")
                        data = data.Where(j => j.Category == value);
                    else if (currentNode.Tag is TreeNodeMeta nodeMeta && nodeMeta.PrefixSegments?.Count > 0)
                    {
                        if (kind == "Material")
                        {
                            var materialName = nodeMeta.MaterialName ?? value;
                            data = data.Where(j => j.MaterialName == materialName);
                        }
                        else if (kind == "MaterialPart")
                        {
                            var prefixSegments = nodeMeta.PrefixSegments;
                            data = data.Where(j =>
                            {
                                var segments = GetSegmentsForMaterial(j.MaterialName);
                                if (segments.Count < prefixSegments.Count)
                                    return false;

                                for (int i = 0; i < prefixSegments.Count; i++)
                                {
                                    if (!string.Equals(segments[i], prefixSegments[i], StringComparison.CurrentCultureIgnoreCase))
                                        return false;
                                }

                                return true;
                            });
                        }
                    }
                }
            }


            // === ПРИХОД: ДАТЫ ===
            if (ArrivalDateFrom?.SelectedDate != null)
                data = data.Where(j => j.Date >= ArrivalDateFrom.SelectedDate);

            if (ArrivalDateTo?.SelectedDate != null)
                data = data.Where(j => j.Date <= ArrivalDateTo.SelectedDate);


            var arrivalSearch = ArrivalSearchBox?.Text?.Trim();
            if (!string.IsNullOrWhiteSpace(arrivalSearch))
            {
                data = data.Where(j =>
                    (j.MaterialName ?? string.Empty).Contains(arrivalSearch, StringComparison.CurrentCultureIgnoreCase)
                    || (j.Ttn ?? string.Empty).Contains(arrivalSearch, StringComparison.CurrentCultureIgnoreCase)
                    || (j.Supplier ?? string.Empty).Contains(arrivalSearch, StringComparison.CurrentCultureIgnoreCase));
            }

            if (arrivalMatrixMode)
            {
                if (selectedArrivalTypes.Count == 0)
                {
                    filteredJournal = new List<JournalRecord>();
                    RenderJvk();
                    RenderArrivalMatrixPlaceholder("Откройте фильтры и выберите один тип материала.");
                    if (ArrivalLegacyGrid != null)
                        ArrivalLegacyGrid.ItemsSource = filteredJournal;
                    RefreshSummaryTable();
                    return;
                }

                var selectedArrivalType = selectedArrivalTypes
                    .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                    .First();

                if (selectedArrivalTypes.Count > 1)
                {
                    selectedArrivalTypes.Clear();
                    selectedArrivalTypes.Add(selectedArrivalType);
                    RefreshArrivalTypes();
                    RefreshArrivalNames();
                }

                data = data.Where(j => string.Equals(j.MaterialGroup, selectedArrivalType, StringComparison.CurrentCultureIgnoreCase));
            }

            if (selectedArrivalTypes.Count > 0)
                data = data.Where(j => selectedArrivalTypes.Contains(j.MaterialGroup));

            if (selectedArrivalNames.Count > 0)
                data = data.Where(j => selectedArrivalNames.Contains(j.MaterialName));

            // === ПРИХОД: СОРТ ПО УМОЛЧАНИЮ ===
            data = data.OrderByDescending(j => j.Date);


            filteredJournal = data.ToList();


            RenderJvk();
            if (ArrivalLegacyGrid != null)
                ArrivalLegacyGrid.ItemsSource = filteredJournal;

            if (initialUiPrepared && IsArrivalTabActive() && arrivalMatrixMode)
                RenderArrivalMatrix();
            else
                RenderArrivalMatrixPlaceholder();
            RefreshSummaryTable();

        }
        private void ArrivalClearFilters_Click(object sender, RoutedEventArgs e)
        {
            selectedArrivalTypes.Clear();
            selectedArrivalNames.Clear();

            ArrivalMainCheck.IsChecked = true;
            ArrivalExtraCheck.IsChecked = true;
            ArrivalLowCostCheck.IsChecked = true;
            ArrivalInternalCheck.IsChecked = true;

            ArrivalDateFrom.SelectedDate = null;
            ArrivalDateTo.SelectedDate = null;

            ArrivalSearchBox.Text = "";

            RefreshArrivalTypes();
            RefreshArrivalNames();
            ApplyAllFilters();

            ArrivalFiltersOverlay.Visibility = Visibility.Collapsed;
        }

        private void ArrivalLegacyGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            ScheduleArrivalLegacyRefresh();
        }

        private void ArrivalLegacyGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            ScheduleArrivalLegacyRefresh();
        }

        private void ScheduleArrivalLegacyRefresh()
        {
            if (arrivalLegacyRefreshPending)
                return;

            arrivalLegacyRefreshPending = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                arrivalLegacyRefreshPending = false;
                RebuildArchiveFromCurrentData();
                SaveState();
                ArrivalPanel.SetObject(currentObject, journal);
                RefreshArrivalTypes();
                RefreshArrivalNames();
                ApplyAllFilters();
            }), DispatcherPriority.Background);
        }

        private sealed class ArrivalMatrixColumn
        {
            public DateTime Date { get; set; }
            public string Ttn { get; set; }
            public string Supplier { get; set; }
            public string Passport { get; set; }
        }

        private bool IsArrivalTabActive()
            => MainTabs?.SelectedItem is TabItem item
               && string.Equals(item.Header?.ToString(), "Приход", StringComparison.CurrentCulture);

        private void RenderArrivalMatrixPlaceholder(string text = "Откройте фильтры и выберите один тип материала.")
        {
            if (ArrivalMatrixHost == null)
                return;

            ArrivalMatrixHost.Children.Clear();
            ArrivalMatrixHost.RowDefinitions.Clear();
            ArrivalMatrixHost.ColumnDefinitions.Clear();
            ArrivalMatrixHost.Children.Add(new TextBlock
            {
                Text = text,
                Margin = new Thickness(12),
                Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                TextWrapping = TextWrapping.Wrap
            });
        }

        private void RenderArrivalMatrix()
        {
            if (ArrivalMatrixHost == null)
                return;

            ArrivalMatrixHost.SnapsToDevicePixels = true;
            ArrivalMatrixHost.UseLayoutRounding = true;
            ArrivalMatrixHost.Children.Clear();
            ArrivalMatrixHost.RowDefinitions.Clear();
            ArrivalMatrixHost.ColumnDefinitions.Clear();

            var data = filteredJournal
                .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName))
                .ToList();

            if (data.Count == 0)
            {
                ArrivalMatrixHost.Children.Add(new TextBlock
                {
                    Text = "Нет данных по выбранным фильтрам.",
                    Margin = new Thickness(12),
                    Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139))
                });
                return;
            }

            var columns = data
                .GroupBy(x => new
                {
                    Date = x.Date.Date,
                    Ttn = (x.Ttn ?? string.Empty).Trim(),
                    Supplier = (x.Supplier ?? string.Empty).Trim(),
                    Passport = (x.Passport ?? string.Empty).Trim()
                })
                .Select(x => new ArrivalMatrixColumn
                {
                    Date = x.Key.Date,
                    Ttn = x.Key.Ttn,
                    Supplier = x.Key.Supplier,
                    Passport = x.Key.Passport
                })
                .OrderBy(x => x.Date)
                .ThenBy(x => x.Ttn, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var materials = data
                .Select(x => x.MaterialName?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var materialColumnWidth = Math.Max(280, Math.Min(420, materials.Max(x => (x?.Length ?? 0) * 9 + 50)));
            ArrivalMatrixHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(materialColumnWidth) });
            ArrivalMatrixHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(64) });
            foreach (var _ in columns)
                ArrivalMatrixHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(56) });

            ArrivalMatrixHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(72) });
            ArrivalMatrixHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(72) });
            ArrivalMatrixHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(108) });
            ArrivalMatrixHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(78) });
            foreach (var _ in materials)
                ArrivalMatrixHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(34) });

            AddArrivalMatrixCell(0, 0, string.Empty, "#FFFFFF", rowSpan: 2);
            AddArrivalMatrixCell(2, 0, "Наименование", "#F9FAFB", rowSpan: 2, fontWeight: FontWeights.SemiBold, fontSize: 18, textAlignment: TextAlignment.Center);

            AddArrivalMatrixSideLabel(0, "Дата");
            AddArrivalMatrixSideLabel(1, "ТТН");
            AddArrivalMatrixSideLabel(2, "Поставщик");
            AddArrivalMatrixSideLabel(3, "Паспорта");

            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var column = columns[columnIndex];
                var gridColumn = columnIndex + 2;

                AddArrivalMatrixHeaderCell(0, gridColumn, column.Date.ToString("dd.MM.yyyy"));
                AddArrivalMatrixHeaderCell(1, gridColumn, string.IsNullOrWhiteSpace(column.Ttn) ? "—" : column.Ttn);
                AddArrivalMatrixHeaderCell(2, gridColumn, string.IsNullOrWhiteSpace(column.Supplier) ? "—" : column.Supplier);
                AddArrivalMatrixPassportCell(3, gridColumn, string.IsNullOrWhiteSpace(column.Passport) ? "—" : column.Passport);
            }

            for (var materialIndex = 0; materialIndex < materials.Count; materialIndex++)
            {
                var material = materials[materialIndex];
                var rowIndex = materialIndex + 4;
                var rowBackground = materialIndex % 2 == 0 ? "#FFFFFF" : "#F9FAFB";

                AddArrivalMatrixCell(rowIndex, 0, material, rowBackground, textAlignment: TextAlignment.Left, padding: new Thickness(10, 6, 10, 6));
                AddArrivalMatrixCell(rowIndex, 1, string.Empty, "#EEF4FF");

                for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
                {
                    var column = columns[columnIndex];
                    var quantity = data
                        .Where(x => string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase)
                                 && x.Date.Date == column.Date
                                 && string.Equals((x.Ttn ?? string.Empty).Trim(), column.Ttn, StringComparison.CurrentCultureIgnoreCase)
                                 && string.Equals((x.Supplier ?? string.Empty).Trim(), column.Supplier, StringComparison.CurrentCultureIgnoreCase)
                                 && string.Equals((x.Passport ?? string.Empty).Trim(), column.Passport, StringComparison.CurrentCultureIgnoreCase))
                        .Sum(x => x.Quantity);

                    AddArrivalMatrixCell(rowIndex, columnIndex + 2, quantity > 0 ? FormatNumber(quantity) : string.Empty, rowBackground);
                }
            }
        }

        private void AddArrivalMatrixSideLabel(int row, string text)
        {
            var border = CreateArrivalMatrixBorder("#EEF4FF");
            Grid.SetRow(border, row);
            Grid.SetColumn(border, 1);

            border.Child = new TextBlock
            {
                Text = text,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                LayoutTransform = new RotateTransform(90),
                FontSize = 16
            };

            ArrivalMatrixHost.Children.Add(border);
        }

        private void AddArrivalMatrixHeaderCell(int row, int column, string text)
        {
            var border = CreateArrivalMatrixBorder("#FFFFFF");
            Grid.SetRow(border, row);
            Grid.SetColumn(border, column);

            border.Child = new TextBlock
            {
                Text = text,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                LayoutTransform = new RotateTransform(90),
                Margin = new Thickness(3)
            };

            ArrivalMatrixHost.Children.Add(border);
        }

        private void AddArrivalMatrixPassportCell(int row, int column, string text)
        {
            var border = CreateArrivalMatrixBorder("#FFFFFF");
            Grid.SetRow(border, row);
            Grid.SetColumn(border, column);

            border.Child = new TextBlock
            {
                Text = text,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(4)
            };

            ArrivalMatrixHost.Children.Add(border);
        }

        private void AddArrivalMatrixCell(int row, int column, string text, string background, int rowSpan = 1, FontWeight? fontWeight = null, double fontSize = 14, TextAlignment textAlignment = TextAlignment.Center, Thickness? padding = null)
        {
            var border = CreateArrivalMatrixBorder(background);
            Grid.SetRow(border, row);
            Grid.SetColumn(border, column);
            if (rowSpan > 1)
                Grid.SetRowSpan(border, rowSpan);

            var textBlock = new TextBlock
            {
                Text = text,
                FontSize = fontSize,
                TextAlignment = textAlignment,
                TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = textAlignment == TextAlignment.Left ? HorizontalAlignment.Left : HorizontalAlignment.Center,
                Margin = padding ?? new Thickness(6, 4, 6, 4)
            };

            if (fontWeight.HasValue)
                textBlock.FontWeight = fontWeight.Value;

            border.Child = textBlock;
            ArrivalMatrixHost.Children.Add(border);
        }

        private static Border CreateArrivalMatrixBorder(string background)
        {
            return new Border
            {
                Background = (Brush)new BrushConverter().ConvertFromString(background),
                BorderBrush = new SolidColorBrush(Color.FromRgb(229, 231, 235)),
                BorderThickness = new Thickness(0.6)
            };
        }


        // ================= СОХРАНЕНИЕ =================

        private void SaveState()
        {
            var json = BuildCurrentStateJson();

            var tempFileName = $"{currentSaveFileName}.tmp";
            File.WriteAllText(tempFileName, json);
            File.Copy(tempFileName, currentSaveFileName, overwrite: true);
            File.Delete(tempFileName);
            lastSavedStateSnapshot = json;
        }

        private string BuildCurrentStateJson()
        {
            return JsonSerializer.Serialize(new AppState
            {
                CurrentObject = currentObject,
                Journal = journal
            });
        }
        private AppState CloneState()
        {
            var currentObjectClone = currentObject == null
                ? null
                : JsonSerializer.Deserialize<ProjectObject>(JsonSerializer.Serialize(currentObject));

            var journalClone = JsonSerializer.Deserialize<List<JournalRecord>>(
                JsonSerializer.Serialize(journal)) ?? new List<JournalRecord>();

            return new AppState
            {
                CurrentObject = currentObjectClone,
                Journal = journalClone
            };
        }

        private const int MaxUndoSteps = 10;

        private void PushUndo()
        {
            // если превышаем лимит — удаляем самый старый шаг
            if (undoStack.Count >= MaxUndoSteps)
            {
                var temp = undoStack.Reverse().Take(MaxUndoSteps - 1).Reverse().ToList();
                undoStack.Clear();
                foreach (var s in temp)
                    undoStack.Push(s);
            }

            undoStack.Push(CloneState());
            redoStack.Clear();
            UpdateUndoRedoButtons();
        }



        private void RestoreState(AppState state)
        {
            currentObject = state.CurrentObject;
            journal = state.Journal ?? new();
            EnsureProjectUiSettings();
            EnsureDocumentLibraries();

            ArrivalPanel.SetObject(currentObject, journal);

            RefreshTreePreserveState();

            RefreshSummaryTable();
            EnsureOtJournalStorage();
            BindOtJournal();
            RefreshBrigadierNames();
            MarkTimesheetOtSyncDirty();
            RequestTimesheetRebuild();
            EnsureProductionJournalStorage();
            RefreshProductionJournalState();
            EnsureInspectionJournalStorage();
            RefreshInspectionJournalState();
            RefreshDocumentLibraries();
            ApplyProjectUiSettings();
            SaveState();
        }

        private void UpdateUndoRedoButtons()
        {
            UndoButton.IsEnabled = undoStack.Count > 0;
            RedoButton.IsEnabled = redoStack.Count > 0;
        }

        private void MigrateDemandLevelsAndProductionJournal()
        {
            if (currentObject == null)
                return;

            currentObject.SummaryMarksByGroup ??= new Dictionary<string, List<string>>();
            currentObject.ProductionJournal ??= new List<ProductionJournalEntry>();
            currentObject.InspectionJournal ??= new List<InspectionJournalEntry>();
            currentObject.PdfDocuments ??= new List<DocumentTreeNode>();
            currentObject.EstimateDocuments ??= new List<DocumentTreeNode>();
            currentObject.AutoSplitMaterialNames ??= new List<string>();
            currentObject.UiSettings ??= new ProjectUiSettings();

            if (currentObject.Demand == null)
                return;

            foreach (var pair in currentObject.Demand)
            {
                var group = pair.Key.Split(new[] { "::" }, StringSplitOptions.None).FirstOrDefault() ?? string.Empty;
                var demand = pair.Value;
                demand.Levels ??= new Dictionary<int, Dictionary<string, double>>();
                demand.MountedLevels ??= new Dictionary<int, Dictionary<string, double>>();

                if (demand.Levels.Count == 0 && demand.Floors != null)
                {
                    foreach (var block in demand.Floors)
                    {
                        demand.Levels[block.Key] = block.Value.ToDictionary(
                            x => LevelMarkHelper.GetLegacyMarkLabel(x.Key),
                            x => x.Value,
                            StringComparer.CurrentCultureIgnoreCase);
                    }
                }

                if (demand.MountedLevels.Count == 0 && demand.MountedFloors != null)
                {
                    foreach (var block in demand.MountedFloors)
                    {
                        demand.MountedLevels[block.Key] = block.Value.ToDictionary(
                            x => LevelMarkHelper.GetLegacyMarkLabel(x.Key),
                            x => x.Value,
                            StringComparer.CurrentCultureIgnoreCase);
                    }
                }

                if (!currentObject.SummaryMarksByGroup.TryGetValue(group, out var marks) || marks == null || marks.Count == 0)
                {
                    marks = demand.Levels.Values
                        .SelectMany(x => x.Keys)
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct(StringComparer.CurrentCultureIgnoreCase)
                        .ToList();

                    if (marks.Count == 0)
                        marks = LevelMarkHelper.GetDefaultMarks(currentObject);

                    currentObject.SummaryMarksByGroup[group] = marks;
                }
            }
        }

        private void LoadState()
        {
            if (!File.Exists(currentSaveFileName))
                return;

            AppState? state;
            try
            {
                state = JsonSerializer.Deserialize<AppState>(
                    File.ReadAllText(currentSaveFileName));
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось загрузить сохранённые данные. Файл состояния повреждён или имеет неверный формат.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                    "Ошибка загрузки",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            currentObject = state?.CurrentObject;
            journal = state?.Journal ?? new();
            lastSavedStateSnapshot = File.ReadAllText(currentSaveFileName);
            MigrateDemandLevelsAndProductionJournal();
            // === ВОССТАНОВЛЕНИЕ АРХИВА ИЗ СТАРЫХ ДАННЫХ ===
            if (currentObject != null)
            {
                currentObject.MaterialTreeSplitRules ??= new Dictionary<string, string>();
                if (currentObject.Archive == null)
                    currentObject.Archive = new ObjectArchive();

                var archive = currentObject.Archive;

                // группы
                foreach (var g in currentObject.MaterialGroups)
                {
                    if (!archive.Groups.Contains(g.Name))
                        archive.Groups.Add(g.Name);

                    if (!archive.Materials.ContainsKey(g.Name))
                        archive.Materials[g.Name] = new();

                    if (currentObject.MaterialNamesByGroup.TryGetValue(g.Name, out var list))
                    {
                        foreach (var m in list)
                            if (!archive.Materials[g.Name].Contains(m))
                                archive.Materials[g.Name].Add(m);
                    }
                }

                // из журнала добираем остальное
                foreach (var rec in journal)
                {
                    if (!string.IsNullOrWhiteSpace(rec.Unit) && !archive.Units.Contains(rec.Unit))
                        archive.Units.Add(rec.Unit);

                    if (!string.IsNullOrWhiteSpace(rec.Supplier) && !archive.Suppliers.Contains(rec.Supplier))
                        archive.Suppliers.Add(rec.Supplier);

                    if (!string.IsNullOrWhiteSpace(rec.Passport) && !archive.Passports.Contains(rec.Passport))
                        archive.Passports.Add(rec.Passport);

                    if (!string.IsNullOrWhiteSpace(rec.Stb) && !archive.Stb.Contains(rec.Stb))
                        archive.Stb.Add(rec.Stb);
                }
            }

            // === АВТОФОРМИРОВАНИЕ АРХИВА ИЗ СТАРЫХ ДАННЫХ ===
            if (currentObject != null && currentObject.Archive == null)
            {
                currentObject.Archive = new ObjectArchive();

                // группы
                foreach (var g in currentObject.MaterialGroups)
                {
                    if (!currentObject.Archive.Groups.Contains(g.Name))
                        currentObject.Archive.Groups.Add(g.Name);

                    if (!currentObject.Archive.Materials.ContainsKey(g.Name))
                        currentObject.Archive.Materials[g.Name] = new();
                }

                // материалы
                foreach (var kv in currentObject.MaterialNamesByGroup)
                {
                    foreach (var m in kv.Value)
                    {
                        if (!currentObject.Archive.Materials[kv.Key].Contains(m))
                            currentObject.Archive.Materials[kv.Key].Add(m);
                    }
                }

                // дополняем из журнала всё остальное
                foreach (var rec in journal)
                {
                    if (!string.IsNullOrWhiteSpace(rec.Unit) && !currentObject.Archive.Units.Contains(rec.Unit))
                        currentObject.Archive.Units.Add(rec.Unit);

                    if (!string.IsNullOrWhiteSpace(rec.Supplier) && !currentObject.Archive.Suppliers.Contains(rec.Supplier))
                        currentObject.Archive.Suppliers.Add(rec.Supplier);

                    if (!string.IsNullOrWhiteSpace(rec.Passport) && !currentObject.Archive.Passports.Contains(rec.Passport))
                        currentObject.Archive.Passports.Add(rec.Passport);

                    if (!string.IsNullOrWhiteSpace(rec.Stb) && !currentObject.Archive.Stb.Contains(rec.Stb))
                        currentObject.Archive.Stb.Add(rec.Stb);
                }
            }
            SeedDemandFromArrivalsIfMissing();
            EnsureOtJournalStorage();
            EnsureProjectUiSettings();
            EnsureDocumentLibraries();
            RefreshDocumentLibraries();

        }

        private bool HasUnsavedChanges()
        {
            var currentSnapshot = BuildCurrentStateJson();
            return !string.Equals(currentSnapshot, lastSavedStateSnapshot, StringComparison.Ordinal);
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (closeConfirmed)
                return;

            CommitOpenEdits();
            if (!HasUnsavedChanges())
                return;

            var result = MessageBox.Show(
                "Сохранить текущие изменения?",
                "Подтверждение закрытия",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
            {
                e.Cancel = true;
                return;
            }

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    SaveState();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Не удалось сохранить изменения.\n{ex.Message}",
                        "Ошибка сохранения",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    e.Cancel = true;
                    return;
                }
            }

            closeConfirmed = true;
        }

        private void SeedDemandFromArrivalsIfMissing()
        {
            if (currentObject == null || journal == null || journal.Count == 0)
                return;

            currentObject.Demand ??= new Dictionary<string, MaterialDemand>(StringComparer.CurrentCultureIgnoreCase);

            var grouped = journal
                .Where(x => !string.IsNullOrWhiteSpace(x.MaterialGroup) && !string.IsNullOrWhiteSpace(x.MaterialName))
                .GroupBy(x => new
                {
                    Group = x.MaterialGroup.Trim(),
                    Material = x.MaterialName.Trim()
                })
                .ToList();

            foreach (var group in grouped)
            {
                var demandKey = BuildDemandKey(group.Key.Group, group.Key.Material);
                if (currentObject.Demand.ContainsKey(demandKey))
                    continue;

                var arrivedTotal = group.Sum(x => Math.Max(0, x.Quantity));
                if (arrivedTotal <= 0.0001)
                    continue;

                var marks = LevelMarkHelper.GetMarksForGroup(currentObject, group.Key.Group);
                if (marks.Count == 0)
                    marks = new List<string> { "0.000" };

                var blocks = Math.Max(1, currentObject.BlocksCount);
                var cells = Math.Max(1, blocks * marks.Count);
                var unit = group.Select(x => x.Unit).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))
                           ?? GetUnitForMaterial(group.Key.Group, group.Key.Material);
                var demand = GetOrCreateDemand(demandKey, unit);
                var isDiscrete = IsDiscreteUnit(unit);
                var totalPlan = isDiscrete
                    ? Math.Ceiling(Math.Max(1, arrivedTotal * 1.25))
                    : Math.Round(Math.Max(1, arrivedTotal * 1.25), 2, MidpointRounding.AwayFromZero);
                var remaining = totalPlan;
                var leftCells = cells;
                var discreteBase = isDiscrete ? (int)totalPlan / cells : 0;
                var discreteRemainder = isDiscrete ? (int)totalPlan % cells : 0;

                for (var block = 1; block <= blocks; block++)
                {
                    if (!demand.Levels.ContainsKey(block))
                        demand.Levels[block] = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

                    foreach (var mark in marks)
                    {
                        leftCells = Math.Max(1, leftCells);
                        double value;
                        if (isDiscrete)
                        {
                            value = discreteBase + (discreteRemainder > 0 ? 1 : 0);
                            if (discreteRemainder > 0)
                                discreteRemainder--;
                        }
                        else
                        {
                            value = Math.Round(remaining / leftCells, 2, MidpointRounding.AwayFromZero);
                        }

                        if (value < 0)
                            value = 0;

                        demand.Levels[block][mark] = value;
                        remaining = Math.Max(0, remaining - value);
                        leftCells--;
                    }
                }
            }
        }

        private void CommitOpenEdits()
        {
            OtJournalGrid?.CommitEdit(DataGridEditingUnit.Cell, true);
            OtJournalGrid?.CommitEdit(DataGridEditingUnit.Row, true);
            TimesheetGrid?.CommitEdit(DataGridEditingUnit.Cell, true);
            TimesheetGrid?.CommitEdit(DataGridEditingUnit.Row, true);
            ProductionJournalGrid?.CommitEdit(DataGridEditingUnit.Cell, true);
            ProductionJournalGrid?.CommitEdit(DataGridEditingUnit.Row, true);
            Keyboard.ClearFocus();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            SaveState();
            MessageBox.Show("Данные сохранены");
        }

        private void SaveAs_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();

            var dlg = new SaveFileDialog
            {
                Filter = "Файл проекта ConstructionControl (*.json)|*.json",
                FileName = System.IO.Path.GetFileName(currentSaveFileName)
            };

            if (dlg.ShowDialog() != true)
                return;

            var previousSaveFileName = currentSaveFileName;
            currentSaveFileName = dlg.FileName;
            CopyProjectStorage(previousSaveFileName, currentSaveFileName);
            EnsureDocumentLibraries();
            SaveState();
            MessageBox.Show("Проект сохранён в новый файл.");
        }

        private static string BuildStorageRootPath(string saveFileName)
        {
            if (string.IsNullOrWhiteSpace(saveFileName))
                return string.Empty;

            var fullSavePath = System.IO.Path.GetFullPath(saveFileName);
            var folder = System.IO.Path.GetDirectoryName(fullSavePath);
            var baseName = System.IO.Path.GetFileNameWithoutExtension(fullSavePath);
            if (string.IsNullOrWhiteSpace(folder) || string.IsNullOrWhiteSpace(baseName))
                return string.Empty;

            return System.IO.Path.Combine(folder, $"{baseName}_files");
        }

        private void CopyProjectStorage(string sourceSaveFileName, string targetSaveFileName)
        {
            var sourceRoot = BuildStorageRootPath(sourceSaveFileName);
            var targetRoot = BuildStorageRootPath(targetSaveFileName);
            if (string.IsNullOrWhiteSpace(sourceRoot) || string.IsNullOrWhiteSpace(targetRoot))
                return;

            if (!Directory.Exists(sourceRoot))
                return;

            if (string.Equals(sourceRoot, targetRoot, StringComparison.OrdinalIgnoreCase))
                return;

            Directory.CreateDirectory(targetRoot);
            foreach (var sourceFile in Directory.EnumerateFiles(sourceRoot, "*", SearchOption.AllDirectories))
            {
                var relativePath = System.IO.Path.GetRelativePath(sourceRoot, sourceFile);
                var targetFile = System.IO.Path.Combine(targetRoot, relativePath);
                var targetFolder = System.IO.Path.GetDirectoryName(targetFile);
                if (!string.IsNullOrWhiteSpace(targetFolder))
                    Directory.CreateDirectory(targetFolder);
                File.Copy(sourceFile, targetFile, overwrite: true);
            }
        }

        private void AppSettings_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            CommitOpenEdits();
            EnsureProjectUiSettings();

            var window = new SettingsWindow(currentObject.UiSettings)
            {
                Owner = this
            };

            if (window.ShowDialog() != true)
                return;

            currentObject.UiSettings = window.ResultSettings ?? new ProjectUiSettings();
            ApplyProjectUiSettings();
            SaveState();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            SaveState();
            LoadState();
            if (currentObject != null)
                ArrivalPanel.SetObject(currentObject, journal);
            RefreshTreePreserveState();
            RefreshSummaryTable();
            RefreshArrivalTypes();
            RefreshArrivalNames();
            RefreshProductionJournalState();
            RefreshInspectionJournalState();
            ApplyAllFilters();
            RequestReminderRefresh(immediate: true);
        }

        private void ExportAllData_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();

            var dlg = new SaveFileDialog
            {
                Filter = "ConstructionControl backup (*.ccbak)|*.ccbak",
                FileName = $"backup_{DateTime.Now:yyyyMMdd_HHmm}.ccbak"
            };
            if (dlg.ShowDialog() != true)
                return;

            var state = new AppState { CurrentObject = currentObject, Journal = journal };
            File.WriteAllText(dlg.FileName, JsonSerializer.Serialize(state));
            MessageBox.Show("Резервная копия сохранена.");
        }

        private void ImportAllData_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();

            var dlg = new OpenFileDialog
            {
                Filter = "ConstructionControl backup (*.ccbak)|*.ccbak|JSON (*.json)|*.json"
            };
            if (dlg.ShowDialog() != true)
                return;

            var state = JsonSerializer.Deserialize<AppState>(File.ReadAllText(dlg.FileName));
            if (state == null)
                return;

            PushUndo();
            RestoreState(state);
            RebuildArchiveFromCurrentData();
            SaveState();
              RefreshTreePreserveState();
              RefreshArrivalTypes();
              RefreshArrivalNames();
              RefreshDocumentLibraries();
              ApplyAllFilters();
          }
        private void LockToggle_Checked(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            LockButton_Checked(sender, e);
        }

        private void LockToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            LockButton_Unchecked(sender, e);
        }



        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            SaveState();
            Close();
        }

        private void ClearObject_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            var firstConfirm = MessageBox.Show(
                "Будут удалены все данные текущего объекта: приход, сводка, ОТ, табель, ПР, осмотры, дерево материалов, ПДФ и сметы.\n\nПродолжить?",
                "Очистка объекта",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (firstConfirm != MessageBoxResult.Yes)
                return;

            var secondConfirm = MessageBox.Show(
                "Подтвердите полную очистку объекта.",
                "Очистка объекта",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (secondConfirm != MessageBoxResult.Yes)
                return;

            var code = Random.Shared.Next(100000, 1000000).ToString(CultureInfo.InvariantCulture);
            var entered = Microsoft.VisualBasic.Interaction.InputBox(
                $"Для окончательного подтверждения введите код: {code}",
                "Код подтверждения",
                string.Empty)?.Trim();

            if (!string.Equals(entered, code, StringComparison.Ordinal))
            {
                MessageBox.Show("Код введен неверно. Очистка отменена.", "Очистка объекта", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            PushUndo();
            ClearCurrentObjectData();
            SaveState();
            MessageBox.Show("Объект полностью очищен.", "Очистка объекта", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ClearCurrentObjectData()
        {
            if (currentObject == null)
                return;

            journal.Clear();

            currentObject.Demand = new Dictionary<string, MaterialDemand>();
            currentObject.Archive = new ObjectArchive();
            currentObject.MaterialGroups = new List<MaterialGroup>();
            currentObject.MaterialCatalog = new List<MaterialCatalogItem>();
            currentObject.MaterialTreeSplitRules = new Dictionary<string, string>();
            currentObject.AutoSplitMaterialNames = new List<string>();
            currentObject.MaterialNamesByGroup = new Dictionary<string, List<string>>();
            currentObject.StbByGroup = new Dictionary<string, string>();
            currentObject.SupplierByGroup = new Dictionary<string, string>();
            currentObject.ArrivalHistory = new List<ArrivalItem>();
            currentObject.SummaryVisibleGroups = new List<string>();
            currentObject.SummaryMarksByGroup = new Dictionary<string, List<string>>();
            currentObject.OtJournal = new List<OtJournalEntry>();
            currentObject.TimesheetPeople = new List<TimesheetPersonEntry>();
            currentObject.ProductionJournal = new List<ProductionJournalEntry>();
            currentObject.ProductionAutoFillSettings = new ProductionAutoFillSettings();
            currentObject.InspectionJournal = new List<InspectionJournalEntry>();
            currentObject.PdfDocuments = new List<DocumentTreeNode>();
            currentObject.EstimateDocuments = new List<DocumentTreeNode>();

            EnsureProjectUiSettings();
            EnsureDocumentLibraries();

            ArrivalPanel.SetObject(currentObject, journal);
            EnsureOtJournalStorage();
            BindOtJournal();
            MarkTimesheetOtSyncDirty();
            RequestTimesheetRebuild(force: true);
            EnsureProductionJournalStorage();
            RefreshProductionJournalState();
            EnsureInspectionJournalStorage();
            RefreshInspectionJournalState();
            RefreshDocumentLibraries();
            RefreshTreePreserveState();
            RefreshSummaryTable();
            RefreshArrivalTypes();
            RefreshArrivalNames();
            ApplyAllFilters();
            RequestReminderRefresh(immediate: true);
        }





        private void ImportExcel_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = "Выберите файл Excel с приходами"
            };

            if (dlg.ShowDialog() != true)
                return;

            using var wb = new XLWorkbook(dlg.FileName);

            var sheetNames = wb.Worksheets
                .Select(s => s.Name)
                .ToList();
            var importWindow = new ExcelImportWindow(dlg.FileName, sheetNames, currentObject)
            {
                Owner = this
            };

            if (importWindow.ShowDialog() != true)
                return;
            foreach (var rec in importWindow.ImportedRecords)
            {
                PushUndo();
                rec.ObjectName = currentObject.Name;
                journal.Add(rec);
                // === ПОПОЛНЕНИЕ АРХИВА ===
                var archive = currentObject.Archive;

                if (!string.IsNullOrWhiteSpace(rec.MaterialGroup))
                {
                    if (!archive.Groups.Contains(rec.MaterialGroup))
                        archive.Groups.Add(rec.MaterialGroup);

                    if (!archive.Materials.ContainsKey(rec.MaterialGroup))
                        archive.Materials[rec.MaterialGroup] = new();

                    if (!archive.Materials[rec.MaterialGroup].Contains(rec.MaterialName))
                        archive.Materials[rec.MaterialGroup].Add(rec.MaterialName);
                }

                if (!string.IsNullOrWhiteSpace(rec.Unit) && !archive.Units.Contains(rec.Unit))
                    archive.Units.Add(rec.Unit);

                if (!string.IsNullOrWhiteSpace(rec.Supplier) && !archive.Suppliers.Contains(rec.Supplier))
                    archive.Suppliers.Add(rec.Supplier);

                if (!string.IsNullOrWhiteSpace(rec.Passport) && !archive.Passports.Contains(rec.Passport))
                    archive.Passports.Add(rec.Passport);

                if (!string.IsNullOrWhiteSpace(rec.Stb) && !archive.Stb.Contains(rec.Stb))
                    archive.Stb.Add(rec.Stb);


                // ====== ОБРАБОТКА ТОЛЬКО ОСНОВНЫХ ======
                if (rec.Category == "Основные")
                {
                    if (!currentObject.MaterialGroups.Any(g => g.Name == rec.MaterialGroup))
                    {
                        currentObject.MaterialGroups.Add(new MaterialGroup
                        {
                            Name = rec.MaterialGroup
                        });

                        currentObject.MaterialNamesByGroup[rec.MaterialGroup] = new List<string>();
                    }

                    if (!currentObject.MaterialNamesByGroup[rec.MaterialGroup]
                            .Contains(rec.MaterialName))
                    {
                        currentObject.MaterialNamesByGroup[rec.MaterialGroup]
                            .Add(rec.MaterialName);
                    }
                }
            }



            // ====== обновляем UI ======
            SaveState();
            RefreshTreePreserveState();

            RefreshSummaryTable();
            ArrivalPanel.SetObject(currentObject, journal);

            if (importWindow.DemandUpdated)
                RefreshSummaryTable();


        }

        public void RefreshTree()
        {
            RefreshTreePreserveState();
        }

        public void RefreshJournal()
        {
            ApplyAllFilters();
        }
        public void RefreshAfterArchiveChange()
        {
            RefreshTreePreserveState();
            ApplyAllFilters();
            RefreshSummaryTable();
            ArrivalPanel.SetObject(currentObject, journal);
            RefreshArrivalTypes();
            RefreshArrivalNames();
        }

        private void OpenArchive_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            var w = new ArchiveWindow(currentObject, journal)
            {
                Owner = this
            };


            if (w.ShowDialog() == true)
            {
                // после изменений — обновляем всё1
                SaveState();
                RefreshTreePreserveState();
                ApplyAllFilters();
                RefreshSummaryTable();
                ArrivalPanel.SetObject(currentObject, journal);
            }
        }





        public void RefreshSummaryTable()
        {
            if (SummaryPanel == null)
                return;

            SummaryPanel.Items.Clear();

            if (currentObject == null)
                return;

            var mainRecords = journal
                .Where(j => j.Category == "Основные");


            var journalGroups = mainRecords
                .Select(j => j.MaterialGroup)
                .Distinct()
                .ToHashSet();
            var recordsByGroupAndMaterial = mainRecords
                 .GroupBy(j => (j.MaterialGroup, j.MaterialName))
                     .ToDictionary(g => g.Key, g => g.ToList());


            var groupOrder = currentObject.MaterialGroups
                .Select(g => g.Name)
                .Where(name => journalGroups.Contains(name))
                .ToList();

            if (groupOrder.Count == 0)
                groupOrder = journalGroups.OrderBy(g => g).ToList();

            RenderSummaryFilters(groupOrder);

            var visibleGroups = currentObject.SummaryVisibleGroups.Count == 0
                ? new List<string>()
                : groupOrder.Where(g => currentObject.SummaryVisibleGroups.Contains(g)).ToList();

            RenderSummaryHeader();

            foreach (var g in visibleGroups)
            {
                RenderMaterialGroup(g);

                var materialNames = GetMaterialsForGroup(g);

                foreach (var mat in materialNames)
                {
                    if (!recordsByGroupAndMaterial.TryGetValue((g, mat), out var records))
                        records = new List<JournalRecord>();

                    string unit = records.FirstOrDefault()?.Unit ?? string.Empty;
                    string position = records
                        .Select(r => r.Position)
                        .FirstOrDefault(p => !string.IsNullOrWhiteSpace(p)) ?? string.Empty;

                    double totalArrival = records.Sum(x => x.Quantity);

                    RenderMaterialRow(g, mat, unit, totalArrival, position);
                }
            }

            RenderSummaryFooter();
        }
        void RenderSummaryHeader()
        {
            var note = new TextBlock
            {
                Text = summaryMountedMode ? "Формат ячейки: смонтировано / пришло" : "Формат ячейки: план / пришло",
                Foreground = new SolidColorBrush(Color.FromRgb(107, 114, 128)),
                Margin = new Thickness(0, 0, 0, 8)
            };

            SummaryPanel.Items.Add(note);
        }

        void RenderMaterialGroup(string group)
        {
            summaryBlocks = BuildSummaryBlocks(group);
            summaryColumns = new List<SummaryColumnInfo>();

            var headerBorder = new Border
            {
                Background = GetColor(group),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10, 6, 10, 6),
                Margin = new Thickness(0, 10, 0, 6)
            };

            headerBorder.Child = new TextBlock
            {
                Text = group,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(31, 41, 55))
            };

            SummaryPanel.Items.Add(headerBorder);

            summaryGrid = new Grid
            {
                Margin = new Thickness(0, 0, 0, 14)
            };

            SummaryPanel.Items.Add(summaryGrid);

            summaryGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            summaryGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 70 });
            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 260 });

            int colIndex = 2;

            foreach (var block in summaryBlocks)
            {
                foreach (var mark in block.Levels)
                {
                    summaryColumns.Add(new SummaryColumnInfo
                    {
                        ColumnIndex = colIndex,
                        Block = block.Block,
                        Level = mark
                    });
                    summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 42 });
                    colIndex++;
                }

                summaryColumns.Add(new SummaryColumnInfo
                {
                    ColumnIndex = colIndex,
                    Block = block.Block,
                    IsBlockTotal = true
                });
                summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 54 });
                colIndex++;
            }

            summaryTotalColumn = colIndex++;
            summaryNotArrivedColumn = colIndex++;
            summaryArrivedColumn = colIndex++;

            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 90 });
            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 90 });
            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 70 });

            var headerBg = new SolidColorBrush(Color.FromRgb(243, 244, 246));

            AddCell(summaryGrid, 0, 0, "Позиция", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);
            AddCell(summaryGrid, 0, 1, "Наименование", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);

            int blockStart = 2;
            foreach (var block in summaryBlocks)
            {
                int blockColumns = block.Levels.Count + 1;

                AddCell(summaryGrid, 0, blockStart, $"Блок {block.Block}", bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, colspan: blockColumns, noWrap: true);

                int floorCol = blockStart;
                for (var levelIndex = 0; levelIndex < block.Levels.Count; levelIndex++)
                {
                    AddSummaryLevelHeaderCell(summaryGrid, 1, floorCol, group, levelIndex, block.Levels[levelIndex], headerBg);
                    floorCol++;
                }

                AddCell(summaryGrid, 1, floorCol, "Итого", bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);
                blockStart += blockColumns;
            }

            AddCell(summaryGrid, 0, summaryTotalColumn, summaryMountedMode ? "Смонтировано" : "Всего на здание", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);
            AddCell(summaryGrid, 0, summaryNotArrivedColumn, summaryMountedMode ? "В остатке" : "Не доехало", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);
            AddCell(summaryGrid, 0, summaryArrivedColumn, "Пришло", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);

            summaryRowIndex = 2;
        }

        void RenderMaterialRow(string group, string mat, string unit, double totalArrival, string position)
        {
            if (summaryGrid == null)
                return;

            totalArrival = NormalizeQuantityByUnit(totalArrival, unit);
            var blockTotals = new Dictionary<int, double>();

            summaryGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            string demandKey = BuildDemandKey(group, mat);
            var demand = GetOrCreateDemand(demandKey, unit);
            var allocations = AllocateArrival(demand, totalArrival);

            double totalPlanned = 0;
            var blockArrivedTotals = new Dictionary<int, double>();
            var blockFilled = new Dictionary<int, bool>();
            var blockOverage = new Dictionary<int, bool>();

            foreach (var block in summaryBlocks)
            {
                double blockTotal = 0;
                foreach (var mark in block.Levels)
                {
                    blockTotal += NormalizeQuantityByUnit(summaryMountedMode
                        ? GetMountedValue(demand, block.Block, mark)
                        : GetDemandValue(demand, block.Block, mark), unit);
                }

                blockTotals[block.Block] = blockTotal;  // теперь ок
                totalPlanned += blockTotal;

                double arrivedTotal = allocations.TryGetValue(block.Block, out var arrivedFloors)
                    ? arrivedFloors.Values.Sum()
                    : 0;

                blockArrivedTotals[block.Block] = arrivedTotal;
                blockFilled[block.Block] = blockTotal > 0 && Math.Abs(arrivedTotal - blockTotal) < 0.0001;
                blockOverage[block.Block] = blockTotal > 0 && arrivedTotal > blockTotal;
            }

            bool rowComplete = summaryBlocks.Count > 0
                && summaryBlocks.All(block => blockTotals.TryGetValue(block.Block, out var blockTotal)
                    && blockTotal > 0
                    && blockFilled.TryGetValue(block.Block, out var filled) && filled);
            bool rowOverage = totalArrival > totalPlanned;
            var filledHighlight = new SolidColorBrush(Color.FromRgb(220, 252, 231));
            var blockHighlight = new SolidColorBrush(Color.FromRgb(219, 234, 254));
            var warningHighlight = new SolidColorBrush(Color.FromRgb(254, 243, 199));

            AddCell(summaryGrid, summaryRowIndex, 0, position, align: TextAlignment.Center, noWrap: true, minWidth: 60);
            AddCell(summaryGrid, summaryRowIndex, 1, mat, noWrap: true, minWidth: 220);

            foreach (var col in summaryColumns)
            {
                if (col.IsBlockTotal)
                {
                    double blockTotal = blockTotals.TryGetValue(col.Block, out var val) ? val : 0;
                   
                    bool blockComplete = blockFilled.TryGetValue(col.Block, out var complete) && complete;
                    bool blockIsOverage = blockOverage.TryGetValue(col.Block, out var over) && over;
                    Brush cellBg = blockIsOverage ? warningHighlight : (blockComplete ? blockHighlight : null);
                    AddCell(summaryGrid, summaryRowIndex, col.ColumnIndex, FormatNumberByUnit(blockTotal, unit), align: TextAlignment.Right, bg: cellBg, noWrap: true, minWidth: 44);
                }
                else if (!string.IsNullOrWhiteSpace(col.Level))
                {
                    double plan = NormalizeQuantityByUnit(GetDemandValue(demand, col.Block, col.Level), unit);
                    double mounted = NormalizeQuantityByUnit(GetMountedValue(demand, col.Block, col.Level), unit);
                    double arrived = allocations.TryGetValue(col.Block, out var blockDict)
                        && blockDict.TryGetValue(col.Level, out var arr)
                        ? NormalizeQuantityByUnit(arr, unit)
                        : 0;

                    double compareBase = summaryMountedMode ? mounted : plan;
                    bool floorOverage = compareBase > 0 ? arrived > compareBase : arrived > 0;
                    bool floorFilled = compareBase > 0 && Math.Abs(arrived - compareBase) < 0.0001;
                    Brush cellBg = floorOverage ? warningHighlight : (floorFilled ? filledHighlight : null);
                    AddDiagonalSummaryCell(summaryGrid, summaryRowIndex, col.ColumnIndex, summaryMountedMode ? mounted : plan, arrived, demandKey, col.Block, col.Level, unit, cellBg, 44, true, summaryMountedMode);
                }
            }

            double notArrived = summaryMountedMode
               ? Math.Max(0, totalArrival - totalPlanned)
               : Math.Max(0, totalPlanned - totalArrival);
            Brush arrivedBg = rowOverage ? warningHighlight : null;
            AddCell(summaryGrid, summaryRowIndex, summaryTotalColumn, FormatNumberByUnit(totalPlanned, unit), align: TextAlignment.Right, bg: rowComplete ? blockHighlight : null, noWrap: true, minWidth: 70);
            AddCell(summaryGrid, summaryRowIndex, summaryNotArrivedColumn, FormatNumberByUnit(notArrived, unit), align: TextAlignment.Right, noWrap: true, minWidth: 70);
            AddCell(summaryGrid, summaryRowIndex, summaryArrivedColumn, FormatNumberByUnit(totalArrival, unit), align: TextAlignment.Right, bg: arrivedBg, noWrap: true, minWidth: 70);

            summaryRowIndex++;
        }

        void RenderSummaryFooter()
        {
            summaryGrid = null;
            summaryColumns = null;
            summaryBlocks = null;
        }

        private void RenderSummaryFilters(List<string> groups)
        {
            
            if (currentObject == null)
                return;

            summaryFilterUpdating = true;
            summaryFilterGroups = groups.ToList();

            SummaryTypeFilterPanel?.Children.Clear();
            SummarySubTypeFilterPanel?.Children.Clear();

            if (groups.Count == 0)
            {
                UpdateSummaryFilterSubtitle(groups, new List<string>());
                summaryFilterUpdating = false;
                return;
            }

            var selectedGroups = currentObject.SummaryVisibleGroups.Where(groups.Contains).ToList();

            if (selectedGroups.Count == 0)
                selectedGroups = new List<string> { groups[0] };
            else if (selectedGroups.Count > 1)
                selectedGroups = selectedGroups.Take(1).ToList();

            currentObject.SummaryVisibleGroups = selectedGroups;

            if (SummaryTypeFilterPanel != null)
            {
                var radioStyle = FindResource("SummaryFilterRadio") as Style;
               

                foreach (var group in groups)
                {
                    var radio = new RadioButton
                    {
                        Content = group,
                        Margin = new Thickness(0, 2, 0, 2),
                        GroupName = "SummaryTypeFilter",
                        IsChecked = selectedGroups.Count == 1 && selectedGroups[0] == group,
                        Tag = group,
                        Style = radioStyle
                    };
                    radio.Checked += SummaryFilterOptionChanged;
                    SummaryTypeFilterPanel.Children.Add(radio);
                }
            }

            RenderSummarySubTypeFilter(selectedGroups.FirstOrDefault());
            UpdateSummaryFilterSubtitle(groups, selectedGroups);
            summaryFilterUpdating = false;
        }

        private void RenderSummarySubTypeFilter(string selectedGroup)
        {
            if (SummarySubTypeFilterPanel == null)
                return;

            SummarySubTypeFilterPanel.Children.Clear();
            if (string.IsNullOrWhiteSpace(selectedGroup))
                return;

            var subTypes = currentObject.MaterialCatalog
                .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                         && string.Equals(x.TypeName ?? string.Empty, selectedGroup, StringComparison.CurrentCultureIgnoreCase))
                .Select(x => x.SubTypeName ?? string.Empty)
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .OrderBy(x => x)
                .ToList();

            if (subTypes.Count == 0)
            {
                summarySelectedSubType = string.Empty;
                return;
            }

            var radioStyle = FindResource("SummaryFilterRadio") as Style;
            var values = new List<string> { "Все" };
            values.AddRange(subTypes);
            if (!values.Contains(summarySelectedSubType))
                summarySelectedSubType = "Все";

            foreach (var subType in values)
            {
                var radio = new RadioButton
                {
                    Content = subType,
                    GroupName = "SummarySubTypeFilter",
                    Margin = new Thickness(0, 2, 0, 2),
                    IsChecked = subType == summarySelectedSubType,
                    Tag = subType,
                    Style = radioStyle
                };
                radio.Checked += SummarySubTypeFilter_Checked;
                SummarySubTypeFilterPanel.Children.Add(radio);
            }
        }
        private void SummaryFilterOptionChanged(object sender, RoutedEventArgs e)
        {
            if (summaryFilterUpdating || currentObject == null)
                return;

            if (sender is not RadioButton radio)
                return;

            var selectedGroup = radio.Tag?.ToString();
            var selected = string.IsNullOrWhiteSpace(selectedGroup) ? new List<string>() : new List<string> { selectedGroup };

            currentObject.SummaryVisibleGroups = selected;
            RenderSummarySubTypeFilter(selectedGroup);
            UpdateSummaryFilterSubtitle(summaryFilterGroups, selected);

            
            RefreshSummaryTable();
        }
        private void SummarySubTypeFilter_Checked(object sender, RoutedEventArgs e)
        {
            if (summaryFilterUpdating)
                return;

            if (sender is RadioButton radio)
            {
                summarySelectedSubType = radio.Tag?.ToString() ?? "Все";
                RefreshSummaryTable();
            }
        }

        private void UpdateSummaryFilterSubtitle(List<string> groups, List<string> selectedGroups)
        {
            if (SummaryFilterSubtitle == null)
                return;

            if (groups == null || groups.Count == 0)
            {
                SummaryFilterSubtitle.Text = "Нет доступных типов";
                return;
            }

            if (selectedGroups == null || selectedGroups.Count == 0)
            {
                SummaryFilterSubtitle.Text = groups[0];
                return;
            }

            SummaryFilterSubtitle.Text = selectedGroups[0];
        }

        private void OpenDemandEditor_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
                return;

            var rows = currentObject.MaterialCatalog
                .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                         && !string.IsNullOrWhiteSpace(x.TypeName)
                         && !string.IsNullOrWhiteSpace(x.MaterialName))
                .Select(x => new DemandEditorWindow.DemandMaterialRow
                {
                    Group = x.TypeName,
                    Material = x.MaterialName,
                    Unit = journal.Where(j => j.MaterialName == x.MaterialName)
                                  .Select(j => j.Unit)
                                  .FirstOrDefault(u => !string.IsNullOrWhiteSpace(u)) ?? string.Empty
                })
                               .DistinctBy(x => $"{x.Group}::{x.Material}")
                .OrderBy(x => x.Group)
                .ThenBy(x => x.Material)
                .ToList();

            var window = new DemandEditorWindow(currentObject, rows)
            {
                Owner = this
            };

            if (window.ShowDialog() == true)
                RefreshSummaryTable();
        }

        private void OpenReorderWindow_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            var availableGroups = summaryFilterGroups
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            if (availableGroups.Count == 0)
            {
                MessageBox.Show("Сначала добавьте типы в сводке.");
                return;
            }

            var defaultSelectedGroups = (currentObject.SummaryVisibleGroups ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Where(x => availableGroups.Contains(x, StringComparer.CurrentCultureIgnoreCase))
                .ToList();
            if (defaultSelectedGroups.Count == 0)
                defaultSelectedGroups.Add(availableGroups[0]);

            var blocks = BuildSummaryBlocks(defaultSelectedGroups[0]);
            var marks = blocks.SelectMany(x => x.Levels).Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();

            var groupOptions = new ObservableCollection<SelectableOption>(availableGroups.Select(x => new SelectableOption
            {
                Value = x,
                IsSelected = defaultSelectedGroups.Contains(x, StringComparer.CurrentCultureIgnoreCase)
            }));

            var blockOptions = new ObservableCollection<SelectableOption>(blocks.Select(x => new SelectableOption
            {
                Value = $"Блок {x.Block}",
                IsSelected = true
            }));
            var markOptions = new ObservableCollection<SelectableOption>(marks.Select(x => new SelectableOption
            {
                Value = x,
                IsSelected = true
            }));
            var materialOptions = new ObservableCollection<SelectableOption>();
            var previewRows = new ObservableCollection<SummaryReorderPreviewRow>();

            var dialog = new Window
            {
                Title = "Дозаказать",
                Owner = this,
                Width = 1080,
                Height = 680,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(280) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var title = new TextBlock
            {
                Text = "Выберите типы, блоки, отметки и материалы для дозаказа.",
                FontWeight = FontWeights.SemiBold,
                FontSize = 15,
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(title, 0);
            Grid.SetColumnSpan(title, 5);
            root.Children.Add(title);

            FrameworkElement BuildSelectionPanel(string caption, ObservableCollection<SelectableOption> options, int column)
            {
                var border = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(12),
                    Padding = new Thickness(10),
                    Margin = new Thickness(0, 0, 10, 0),
                    Background = Brushes.White
                };
                Grid.SetRow(border, 1);
                Grid.SetColumn(border, column);

                var panel = new DockPanel();
                border.Child = panel;
                var captionText = new TextBlock
                {
                    Text = caption,
                    FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(0, 0, 0, 8)
                };
                panel.Children.Add(captionText);
                DockPanel.SetDock(captionText, Dock.Top);

                var list = new ListBox
                {
                    BorderThickness = new Thickness(0),
                    Background = Brushes.Transparent,
                    ItemsSource = options
                };
                var template = new DataTemplate(typeof(SelectableOption));
                var factory = new FrameworkElementFactory(typeof(CheckBox));
                factory.SetBinding(ToggleButton.IsCheckedProperty, new Binding(nameof(SelectableOption.IsSelected)) { Mode = BindingMode.TwoWay });
                factory.SetBinding(ContentControl.ContentProperty, new Binding(nameof(SelectableOption.Value)));
                template.VisualTree = factory;
                list.ItemTemplate = template;
                panel.Children.Add(list);
                return border;
            }

            root.Children.Add(BuildSelectionPanel("Типы", groupOptions, 0));
            root.Children.Add(BuildSelectionPanel("Блоки", blockOptions, 1));
            root.Children.Add(BuildSelectionPanel("Отметки", markOptions, 2));
            root.Children.Add(BuildSelectionPanel("Материалы", materialOptions, 3));

            var previewGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                ItemsSource = previewRows
            };
            previewGrid.Columns.Add(new DataGridTextColumn { Header = "Тип", Binding = new Binding(nameof(SummaryReorderPreviewRow.Group)) });
            previewGrid.Columns.Add(new DataGridTextColumn { Header = "Наименование", Binding = new Binding(nameof(SummaryReorderPreviewRow.Material)) });
            previewGrid.Columns.Add(new DataGridTextColumn { Header = "Блок", Binding = new Binding(nameof(SummaryReorderPreviewRow.Block)) });
            previewGrid.Columns.Add(new DataGridTextColumn { Header = "Отметка", Binding = new Binding(nameof(SummaryReorderPreviewRow.Mark)) });
            previewGrid.Columns.Add(new DataGridTextColumn { Header = "Количество", Binding = new Binding(nameof(SummaryReorderPreviewRow.Quantity)) { StringFormat = "0.##" } });
            previewGrid.Columns.Add(new DataGridTextColumn { Header = "Ед.", Binding = new Binding(nameof(SummaryReorderPreviewRow.Unit)) });
            Grid.SetRow(previewGrid, 1);
            Grid.SetColumn(previewGrid, 4);
            DataGridSizingHelper.SetEnableSmartSizing(previewGrid, true);
            root.Children.Add(previewGrid);

            void SyncMaterialOptions(bool selectAllByDefault)
            {
                var selectedGroups = groupOptions
                    .Where(x => x.IsSelected && !string.IsNullOrWhiteSpace(x.Value))
                    .Select(x => x.Value.Trim())
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                var selectedBefore = materialOptions
                    .Where(x => x.IsSelected)
                    .Select(x => x.Value)
                    .ToHashSet(StringComparer.CurrentCultureIgnoreCase);

                materialOptions.Clear();

                foreach (var selectedGroup in selectedGroups)
                {
                    foreach (var material in GetMaterialsForGroup(selectedGroup))
                    {
                        var key = $"{selectedGroup}::{material}";
                        materialOptions.Add(new SelectableOption
                        {
                            Value = key,
                            IsSelected = selectAllByDefault || selectedBefore.Contains(key)
                        });
                    }
                }
            }

            void RefreshPreview()
            {
                previewRows.Clear();
                var selectedBlocks = blockOptions
                    .Where(x => x.IsSelected)
                    .Select(x => int.TryParse(x.Value.Replace("Блок", string.Empty).Trim(), out var block) ? block : 0)
                    .Where(x => x > 0)
                    .ToHashSet();
                var selectedMarks = markOptions.Where(x => x.IsSelected).Select(x => x.Value).ToHashSet(StringComparer.CurrentCultureIgnoreCase);
                var selectedMaterials = materialOptions.Where(x => x.IsSelected).Select(x => x.Value).ToList();

                foreach (var row in BuildSummaryReorderRows(selectedMaterials, selectedBlocks, selectedMarks))
                    previewRows.Add(row);
            }

            foreach (var option in groupOptions)
            {
                option.PropertyChanged += (_, _) =>
                {
                    SyncMaterialOptions(selectAllByDefault: false);
                    foreach (var materialOption in materialOptions)
                        materialOption.PropertyChanged += (_, _) => RefreshPreview();
                    RefreshPreview();
                };
            }

            foreach (var option in blockOptions.Cast<SelectableOption>().Concat(markOptions))
            {
                option.PropertyChanged += (_, _) => RefreshPreview();
            }

            SyncMaterialOptions(selectAllByDefault: true);
            foreach (var materialOption in materialOptions)
                materialOption.PropertyChanged += (_, _) => RefreshPreview();
            RefreshPreview();

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            Grid.SetRow(footer, 2);
            Grid.SetColumnSpan(footer, 5);
            root.Children.Add(footer);

            var exportButton = new Button { Content = "Экспорт в Word", MinWidth = 140 };
            var closeButton = new Button { Content = "Закрыть", Style = FindResource("SecondaryButton") as Style, MinWidth = 110, Margin = new Thickness(10, 0, 0, 0), IsCancel = true };
            footer.Children.Add(exportButton);
            footer.Children.Add(closeButton);

            exportButton.Click += (_, _) =>
            {
                RefreshPreview();
                if (previewRows.Count == 0)
                {
                    MessageBox.Show("По выбранным блокам и отметкам нет позиций для дозаказа.");
                    return;
                }

                var dlg = new SaveFileDialog
                {
                    Filter = "Word (*.docx)|*.docx",
                    FileName = $"Дозаказать_{DateTime.Today:yyyyMMdd}.docx"
                };

                if (dlg.ShowDialog() != true)
                    return;

                var selectedGroups = groupOptions.Where(x => x.IsSelected).Select(x => x.Value).ToList();
                ExportSummaryReorderToWord(dlg.FileName, selectedGroups, previewRows.ToList());
                MessageBox.Show("Файл Word сформирован.");
            };

            dialog.Content = root;
            dialog.ShowDialog();
        }

        private List<SummaryReorderPreviewRow> BuildSummaryReorderRows(IEnumerable<string> selectedMaterialKeys, IEnumerable<int> selectedBlocks, IEnumerable<string> selectedMarks)
        {
            var rows = new List<SummaryReorderPreviewRow>();
            var materials = (selectedMaterialKeys ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x) && x.Contains("::"))
                .Select(x =>
                {
                    var parts = x.Split(new[] { "::" }, 2, StringSplitOptions.None);
                    return new
                    {
                        Group = parts[0].Trim(),
                        Material = parts.Length > 1 ? parts[1].Trim() : string.Empty
                    };
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.Group) && !string.IsNullOrWhiteSpace(x.Material))
                .DistinctBy(x => $"{x.Group}::{x.Material}")
                .ToList();
            var blocks = (selectedBlocks ?? Enumerable.Empty<int>()).Where(x => x > 0).ToHashSet();
            var marks = (selectedMarks ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase);

            if (materials.Count == 0 || blocks.Count == 0 || marks.Count == 0)
                return rows;

            foreach (var materialSelection in materials
                .OrderBy(x => x.Group, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.Material, StringComparer.CurrentCultureIgnoreCase))
            {
                var group = materialSelection.Group;
                var material = materialSelection.Material;
                var summaryBlockSet = BuildSummaryBlocks(group);
                var records = journal
                    .Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals(x.MaterialGroup ?? string.Empty, group, StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals(x.MaterialName ?? string.Empty, material, StringComparison.CurrentCultureIgnoreCase))
                    .ToList();

                var unit = records.Select(x => x.Unit).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? GetUnitForMaterial(group, material);
                var demand = GetOrCreateDemand(BuildDemandKey(group, material), unit);
                var totalArrivedOnBuilding = records.Sum(x => x.Quantity);
                var totalNeedOnBuilding = summaryBlockSet
                    .SelectMany(x => x.Levels.Select(level => GetDemandValue(demand, x.Block, level)))
                    .Sum();
                var remainingBuildingDeficit = Math.Max(0, totalNeedOnBuilding - totalArrivedOnBuilding);
                if (remainingBuildingDeficit <= 0.0001)
                    continue;

                var allocations = AllocateArrivalForBlocks(summaryBlockSet, demand, totalArrivedOnBuilding);
                var isDeficitExhausted = false;

                foreach (var block in summaryBlockSet.Where(x => blocks.Contains(x.Block)))
                {
                    foreach (var mark in block.Levels.Where(marks.Contains))
                    {
                        var planned = GetDemandValue(demand, block.Block, mark);
                        var arrived = allocations.TryGetValue(block.Block, out var blockMap)
                            && blockMap.TryGetValue(mark, out var arrivedValue)
                            ? arrivedValue
                            : 0;
                        var reorder = Math.Max(0, planned - arrived);
                        reorder = Math.Min(reorder, remainingBuildingDeficit);
                        if (reorder <= 0.0001)
                            continue;

                        rows.Add(new SummaryReorderPreviewRow
                        {
                            Group = group,
                            Material = material,
                            Block = block.Block,
                            Mark = mark,
                            Quantity = reorder,
                            Unit = unit
                        });

                        remainingBuildingDeficit -= reorder;
                        if (remainingBuildingDeficit <= 0.0001)
                        {
                            isDeficitExhausted = true;
                            break;
                        }
                    }

                    if (isDeficitExhausted)
                        break;
                }
            }

            return rows;
        }

        private void ExportSummaryReorderToWord(string filePath, IEnumerable<string> selectedGroups, List<SummaryReorderPreviewRow> rows)
        {
            using var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(
                filePath,
                DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            var body = new DocumentFormat.OpenXml.Wordprocessing.Body();
            mainPart.Document.Append(body);

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("Дозаказ материалов"))));

            if (!string.IsNullOrWhiteSpace(currentObject?.Name))
            {
                body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text($"Объект: {currentObject.Name}"))));
            }

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text($"Дата формирования: {DateTime.Today:dd.MM.yyyy}"))));

            var selectedTypesText = string.Join(", ", (selectedGroups ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase));
            if (!string.IsNullOrWhiteSpace(selectedTypesText))
            {
                body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text($"Типы: {selectedTypesText}"))));
            }

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text(string.Empty))));

            foreach (var groupRows in rows
                         .OrderBy(x => x.Group, StringComparer.CurrentCultureIgnoreCase)
                         .ThenBy(x => x.Material, StringComparer.CurrentCultureIgnoreCase)
                         .ThenBy(x => x.Block)
                         .ThenBy(x => x.Mark)
                         .GroupBy(x => x.Group))
            {
                body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text(groupRows.Key))));

                foreach (var row in groupRows)
                {
                    var line = $"• {row.Material}: Блок {row.Block}, отметка {row.Mark}, количество {FormatNumberByUnit(row.Quantity, row.Unit)} {row.Unit}";
                    body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text(line))));
                }

                body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text(string.Empty))));
            }

            mainPart.Document.Save();
        }

        private void SummaryModeSwitch_Changed(object sender, RoutedEventArgs e)
        {
            summaryMountedMode = SummaryMountedModeSwitch?.IsChecked == true;
            RefreshSummaryTable();
        }


        private List<string> GetMaterialsForGroup(string group)
        {
            if (currentObject.MaterialCatalog?.Count > 0)
            {
                var catalogMaterials = currentObject.MaterialCatalog
                    .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals(x.TypeName ?? string.Empty, group, StringComparison.CurrentCultureIgnoreCase)
                             && (string.IsNullOrWhiteSpace(summarySelectedSubType)
                                 || summarySelectedSubType == "Все"
                                 || string.Equals(x.SubTypeName ?? string.Empty, summarySelectedSubType, StringComparison.CurrentCultureIgnoreCase)))
                    .Select(x => x.MaterialName)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .OrderBy(n => n)
                    .ToList();

                if (catalogMaterials.Count > 0)
                    return catalogMaterials;
            }
            if (currentObject.MaterialNamesByGroup.TryGetValue(group, out var list) && list.Count > 0)
                return list;

            return journal
                .Where(j => j.Category == "Основные" && j.MaterialGroup == group)
                .Select(j => j.MaterialName)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
        }
        private void TreePinToggle_Checked(object sender, RoutedEventArgs e)
        {
            isTreePinned = true;
            UpdateTreePanelState(forceVisible: true);
        }

        private void TreePinToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            isTreePinned = false;
            UpdateTreePanelState(forceVisible: false);
        }

        private void TreeHoverZone_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isTreePinned)
                UpdateTreePanelState(forceVisible: true);
        }

        private void TreePanel_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isTreePinned)
                UpdateTreePanelState(forceVisible: true);
        }

        private void TreePanel_MouseLeave(object sender, MouseEventArgs e)
        {
            if (isTreePinned)
                return;

            if (!TreePanelBorder.IsMouseOver)
                UpdateTreePanelState(forceVisible: false);
        }

        private void ContentPanel_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isTreePinned)
                UpdateTreePanelState(forceVisible: false);
        }

        private bool IsDocumentLibraryTabSelected()
            => ReferenceEquals(MainTabs?.SelectedItem, PdfTab) || ReferenceEquals(MainTabs?.SelectedItem, EstimateTab);

        private void UpdateTreePanelState(bool forceVisible)
        {
            if (TreeColumn == null || TreePanelBorder == null || TreeHoverColumn == null || TreeHoverStrip == null)
                return;

            EnsureProjectUiSettings();

            if (currentObject?.UiSettings?.DisableTree == true)
            {
                TreePanelBorder.Visibility = Visibility.Collapsed;
                TreeHoverStrip.Visibility = Visibility.Collapsed;
                TreeColumn.Width = new GridLength(0);
                TreeHoverColumn.Width = new GridLength(0);
                if (TreePinToggle != null)
                    TreePinToggle.IsChecked = false;
                return;
            }

            if (IsDocumentLibraryTabSelected())
            {
                TreePanelBorder.Visibility = Visibility.Collapsed;
                TreeHoverStrip.Visibility = Visibility.Collapsed;
                TreeColumn.Width = new GridLength(0);
                TreeHoverColumn.Width = new GridLength(0);
                return;
            }

            bool show = isTreePinned || forceVisible;
            TreeHoverStrip.Visibility = Visibility.Visible;
            TreePanelBorder.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            TreeHoverColumn.Width = new GridLength(28);
            TreeColumn.Width = show ? new GridLength(260) : new GridLength(0);
        }

        private void PdfTreePinToggle_Checked(object sender, RoutedEventArgs e)
        {
            isPdfTreePinned = true;
            UpdatePdfTreePanelState(forceVisible: true);
        }

        private void PdfTreePinToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            isPdfTreePinned = false;
            UpdatePdfTreePanelState(forceVisible: false);
        }

        private void PdfTreeHoverZone_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isPdfTreePinned)
                UpdatePdfTreePanelState(forceVisible: true);
        }

        private void PdfTreePanel_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isPdfTreePinned)
                UpdatePdfTreePanelState(forceVisible: true);
        }

        private void PdfTreePanel_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!isPdfTreePinned && !PdfTreePanelBorder.IsMouseOver)
                UpdatePdfTreePanelState(forceVisible: false);
        }

        private void PdfPreviewArea_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isPdfTreePinned)
                UpdatePdfTreePanelState(forceVisible: false);
        }

        private void UpdatePdfTreePanelState(bool forceVisible)
        {
            if (PdfTreeColumn == null || PdfTreePanelBorder == null || PdfTreeHoverColumn == null || PdfTreeHoverStrip == null)
                return;

            bool showStrip = ReferenceEquals(MainTabs?.SelectedItem, PdfTab);
            if (!showStrip)
            {
                PdfTreePanelBorder.Visibility = Visibility.Collapsed;
                PdfTreeHoverStrip.Visibility = Visibility.Collapsed;
                PdfTreeColumn.Width = new GridLength(0);
                PdfTreeHoverColumn.Width = new GridLength(0);
                return;
            }

            bool show = isPdfTreePinned || forceVisible;
            PdfTreeHoverStrip.Visibility = Visibility.Visible;
            PdfTreeHoverColumn.Width = new GridLength(28);
            PdfTreePanelBorder.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            PdfTreeColumn.Width = show ? new GridLength(320) : new GridLength(0);
        }

        private void EstimateTreePinToggle_Checked(object sender, RoutedEventArgs e)
        {
            isEstimateTreePinned = true;
            UpdateEstimateTreePanelState(forceVisible: true);
        }

        private void EstimateTreePinToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            isEstimateTreePinned = false;
            UpdateEstimateTreePanelState(forceVisible: false);
        }

        private void EstimateTreeHoverZone_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isEstimateTreePinned)
                UpdateEstimateTreePanelState(forceVisible: true);
        }

        private void EstimateTreePanel_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isEstimateTreePinned)
                UpdateEstimateTreePanelState(forceVisible: true);
        }

        private void EstimateTreePanel_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!isEstimateTreePinned && !EstimateTreePanelBorder.IsMouseOver)
                UpdateEstimateTreePanelState(forceVisible: false);
        }

        private void EstimatePreviewArea_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!isEstimateTreePinned)
                UpdateEstimateTreePanelState(forceVisible: false);
        }

        private void UpdateEstimateTreePanelState(bool forceVisible)
        {
            if (EstimateTreeColumn == null || EstimateTreePanelBorder == null || EstimateTreeHoverColumn == null || EstimateTreeHoverStrip == null)
                return;

            bool showStrip = ReferenceEquals(MainTabs?.SelectedItem, EstimateTab);
            if (!showStrip)
            {
                EstimateTreePanelBorder.Visibility = Visibility.Collapsed;
                EstimateTreeHoverStrip.Visibility = Visibility.Collapsed;
                EstimateTreeColumn.Width = new GridLength(0);
                EstimateTreeHoverColumn.Width = new GridLength(0);
                return;
            }

            bool show = isEstimateTreePinned || forceVisible;
            EstimateTreeHoverStrip.Visibility = Visibility.Visible;
            EstimateTreeHoverColumn.Width = new GridLength(28);
            EstimateTreePanelBorder.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            EstimateTreeColumn.Width = show ? new GridLength(320) : new GridLength(0);
        }

        private void EnsureProjectUiSettings()
        {
            if (currentObject != null)
                currentObject.UiSettings ??= new ProjectUiSettings();

            if (currentObject?.UiSettings != null && currentObject.UiSettings.ReminderSnoozeMinutes <= 0)
                currentObject.UiSettings.ReminderSnoozeMinutes = 15;
        }

        private void ApplyProjectUiSettings()
        {
            EnsureProjectUiSettings();

            if (currentObject?.UiSettings == null)
            {
                UpdateTreePanelState(forceVisible: false);
                return;
            }

            isTreePinned = currentObject.UiSettings.PinTreeByDefault && !currentObject.UiSettings.DisableTree;
            if (TreePinToggle != null)
                TreePinToggle.IsChecked = isTreePinned;

            UpdateTreePanelState(forceVisible: isTreePinned);
            RequestReminderRefresh(immediate: true);
        }
        private List<SummaryBlockInfo> BuildSummaryBlocks(string group)
        {
            var blocks = new List<SummaryBlockInfo>();

            if (currentObject == null || currentObject.BlocksCount <= 0)
                return blocks;

            var marks = LevelMarkHelper.GetMarksForGroup(currentObject, group);

            for (int i = 1; i <= currentObject.BlocksCount; i++)
            {
                blocks.Add(new SummaryBlockInfo
                {
                    Block = i,
                    Levels = marks.ToList()
                });
            }

            return blocks;
        }

        private Dictionary<int, Dictionary<string, double>> AllocateArrivalForBlocks(List<SummaryBlockInfo> blocks, MaterialDemand demand, double totalArrival)
        {
            var allocations = new Dictionary<int, Dictionary<string, double>>();
            if (blocks == null || blocks.Count == 0)
                return allocations;

            double remaining = NormalizeQuantityByUnit(totalArrival, demand?.Unit);
            foreach (var level in blocks.SelectMany(b => b.Levels).Distinct(StringComparer.CurrentCultureIgnoreCase))
            {
                foreach (var block in blocks)
                {
                    if (!block.Levels.Contains(level))
                        continue;

                    double plan = NormalizeQuantityByUnit(GetDemandValue(demand, block.Block, level), demand?.Unit);
                    double filled = Math.Min(plan, remaining);
                    if (IsDiscreteUnit(demand?.Unit))
                        filled = Math.Floor(filled);

                    if (!allocations.ContainsKey(block.Block))
                        allocations[block.Block] = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

                    allocations[block.Block][level] = filled;
                    remaining = NormalizeQuantityByUnit(remaining - filled, demand?.Unit);
                    if (remaining <= 0)
                        return allocations;
                }
            }

            return allocations;
        }

        private Dictionary<int, Dictionary<string, double>> AllocateArrival(MaterialDemand demand, double totalArrival)
        {
            var allocations = new Dictionary<int, Dictionary<string, double>>();

            if (summaryBlocks == null || summaryBlocks.Count == 0)
                return allocations;

            double remaining = NormalizeQuantityByUnit(totalArrival, demand?.Unit);

            foreach (var level in summaryBlocks.SelectMany(b => b.Levels).Distinct(StringComparer.CurrentCultureIgnoreCase))
            {
                foreach (var block in summaryBlocks)
                {
                    if (!block.Levels.Contains(level))
                        continue;

                    double plan = NormalizeQuantityByUnit(GetDemandValue(demand, block.Block, level), demand?.Unit);
                    double filled = Math.Min(plan, remaining);
                    if (IsDiscreteUnit(demand?.Unit))
                        filled = Math.Floor(filled);

                    if (!allocations.ContainsKey(block.Block))
                        allocations[block.Block] = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

                    allocations[block.Block][level] = filled;
                    remaining = NormalizeQuantityByUnit(remaining - filled, demand?.Unit);

                    if (remaining <= 0)
                        return allocations;
                }
            }

            return allocations;
        }

        private string BuildDemandKey(string group, string material) => $"{group}::{material}";

        private MaterialDemand GetOrCreateDemand(string demandKey, string unit)
        {
            if (!currentObject.Demand.TryGetValue(demandKey, out var demand))
            {
                demand = new MaterialDemand
                {
                    Unit = unit,
                    Levels = new Dictionary<int, Dictionary<string, double>>(),
                    MountedLevels = new Dictionary<int, Dictionary<string, double>>(),
                    Floors = new Dictionary<int, Dictionary<int, double>>(),
                    MountedFloors = new Dictionary<int, Dictionary<int, double>>()
                };

                currentObject.Demand[demandKey] = demand;
            }

            if (string.IsNullOrWhiteSpace(demand.Unit))
                demand.Unit = unit;
            demand.Levels ??= new Dictionary<int, Dictionary<string, double>>();
            demand.MountedLevels ??= new Dictionary<int, Dictionary<string, double>>();
            demand.Floors ??= new Dictionary<int, Dictionary<int, double>>();
            demand.MountedFloors ??= new Dictionary<int, Dictionary<int, double>>();

            return demand;
        }

        private void JvkLayout_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (e.WidthChanged && JvkTab?.IsSelected == true)
                RenderJvk();
        }

        private double GetDemandValue(MaterialDemand demand, int block, string level)
        {
            if (demand.Levels != null
                && demand.Levels.TryGetValue(block, out var levels)
                && levels.TryGetValue(level, out var value))
                return value;

            return 0;
        }
        private double GetMountedValue(MaterialDemand demand, int block, string level)
        {
            if (demand.MountedLevels != null
                && demand.MountedLevels.TryGetValue(block, out var levels)
                && levels.TryGetValue(level, out var value))
                return value;

            return 0;
        }

        private class SummaryBlockInfo
        {
            public int Block { get; set; }
            public List<string> Levels { get; set; } = new();
        }

        private class SummaryColumnInfo
        {
            public int ColumnIndex { get; set; }
            public int Block { get; set; }
            public string Level { get; set; }
            public bool IsBlockTotal { get; set; }
        }

        private class DemandCellTag
        {
            public string DemandKey { get; set; }
            public int Block { get; set; }
            public string Level { get; set; }
            public string Unit { get; set; }
            public bool IsMountedMode { get; set; }
        }

        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();

            if (undoStack.Count == 0)
                return;

            redoStack.Push(CloneState());
            var prev = undoStack.Pop();
            RestoreState(prev);
            UpdateUndoRedoButtons();
        }

        private void Redo_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();

            if (redoStack.Count == 0)
                return;

            undoStack.Push(CloneState());
            var next = redoStack.Pop();
            RestoreState(next);
            UpdateUndoRedoButtons();
        }
        void AddCell(Grid g, int r, int c, string text, int rowspan = 1, bool wrap = false, Brush bg = null, TextAlignment align = TextAlignment.Left, FontWeight? fontWeight = null, int colspan = 1, bool noWrap = false, double? minWidth = null)
        {
            var tb = new TextBlock
            {
                Text = text,
                Margin = new Thickness(6, 4, 6, 4),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = noWrap ? TextWrapping.NoWrap : TextWrapping.Wrap,
                TextTrimming = TextTrimming.None
            };
            if (fontWeight.HasValue)
                tb.FontWeight = fontWeight.Value;

            var border = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(180, 187, 198)),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Background = bg,
                MinHeight = 30
            };
            if (minWidth.HasValue)
                border.MinWidth = minWidth.Value;

            border.Child = tb;

            Grid.SetRow(border, r);
            Grid.SetColumn(border, c);

            if (rowspan > 1)
                Grid.SetRowSpan(border, rowspan);

            if (colspan > 1)
                Grid.SetColumnSpan(border, colspan);


            g.Children.Add(border);
        }
        void AddSummaryLevelHeaderCell(Grid g, int r, int c, string group, int levelIndex, string text, Brush bg)
        {
            var box = new TextBox
            {
                Text = text,
                Margin = new Thickness(4, 2, 4, 2),
                Padding = new Thickness(2),
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                FontWeight = FontWeights.SemiBold,
                Tag = $"{group}|{levelIndex}"
            };
            box.LostFocus += SummaryLevelHeader_LostFocus;

            var border = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(180, 187, 198)),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Background = bg,
                MinHeight = 30,
                Child = box
            };

            Grid.SetRow(border, r);
            Grid.SetColumn(border, c);
            g.Children.Add(border);
        }
        void AddDiagonalSummaryCell(Grid g, int r, int c, double topValue, double arrived, string demandKey, int block, string level, string unit, Brush bg, double minWidth, bool editableTop, bool isMountedMode)
        {
            var container = new Grid
            {
                SnapsToDevicePixels = true,
                UseLayoutRounding = true
            };

            var line = new WpfPath
            {
                Data = Geometry.Parse("M0,1 L1,0"),
                Stroke = new SolidColorBrush(Color.FromRgb(209, 213, 219)),
                StrokeThickness = 1,
                Stretch = Stretch.Fill,
                SnapsToDevicePixels = true,
                IsHitTestVisible = false
            };



            container.Children.Add(line);

            var topBox = new TextBox
            {
                Text = FormatNumber(topValue),
                Style = null,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Margin = new Thickness(2, 1, 2, 1),
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                MinWidth = 22,
                FontSize = 11,
                IsReadOnly = !editableTop || isLocked,
                IsEnabled = editableTop && !isLocked,
                Tag = new DemandCellTag
                {
                    DemandKey = demandKey,
                    Block = block,
                    Level = level,
                    Unit = unit,
                    IsMountedMode = isMountedMode
                }
            };

            if (editableTop)
                topBox.LostFocus += SummaryCell_LostFocus;

            var arrivedText = new TextBlock
            {
                Text = FormatNumber(arrived),
                Margin = new Thickness(2, 1, 2, 1),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Foreground = new SolidColorBrush(Color.FromRgb(55, 65, 81)),
                FontSize = 11
            };

            container.Children.Add(topBox);
            container.Children.Add(arrivedText);

            var border = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(180, 187, 198)),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Background = bg ?? Brushes.White,
                MinHeight = 30,
                Child = container
            };
            if (minWidth > 0)
                border.MinWidth = minWidth;
            Grid.SetRow(border, r);
            Grid.SetColumn(border, c);

            g.Children.Add(border);
        }

        private void SummaryCell_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is not TextBox tb || tb.Tag is not DemandCellTag tag)
                return;

            if (isLocked)
                return;


            var text = tb.Text?.Trim() ?? string.Empty;
            double value = NormalizeQuantityByUnit(ParseNumber(text), tag.Unit);

            var demand = GetOrCreateDemand(tag.DemandKey, tag.Unit);

            var target = tag.IsMountedMode ? demand.MountedLevels : demand.Levels;
            target ??= new Dictionary<int, Dictionary<string, double>>();

            if (tag.IsMountedMode)
                demand.MountedLevels = target;
            else
                demand.Levels = target;

            if (!target.ContainsKey(tag.Block))
                target[tag.Block] = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

            target[tag.Block][tag.Level] = value;

            RefreshSummaryTable();
        }

        private void SummaryLevelHeader_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is not TextBox tb || currentObject == null || tb.Tag is not string tagText)
                return;

            var parts = tagText.Split('|');
            if (parts.Length != 2 || !int.TryParse(parts[1], out var levelIndex))
                return;

            var group = parts[0];
            var newValue = tb.Text?.Trim();
            if (string.IsNullOrWhiteSpace(newValue))
                return;

            var marks = LevelMarkHelper.GetMarksForGroup(currentObject, group);
            if (levelIndex < 0 || levelIndex >= marks.Count)
                return;

            var oldValue = marks[levelIndex];
            if (string.Equals(oldValue, newValue, StringComparison.CurrentCultureIgnoreCase))
                return;

            marks[levelIndex] = newValue;
            currentObject.SummaryMarksByGroup[group] = marks;

            foreach (var demandPair in currentObject.Demand.Where(x => x.Key.StartsWith(group + "::", StringComparison.CurrentCultureIgnoreCase)))
            {
                RenameDemandLevel(demandPair.Value.Levels, oldValue, newValue);
                RenameDemandLevel(demandPair.Value.MountedLevels, oldValue, newValue);
            }

            RefreshProductionJournalLookups();
            RefreshProductionRemainingInfo();
            RefreshSummaryTable();
            SaveState();
        }

        private static void RenameDemandLevel(Dictionary<int, Dictionary<string, double>> map, string oldValue, string newValue)
        {
            if (map == null)
                return;

            foreach (var block in map.Values)
            {
                if (!block.TryGetValue(oldValue, out var amount))
                    continue;

                block.Remove(oldValue);
                if (!block.ContainsKey(newValue))
                    block[newValue] = 0;
                block[newValue] += amount;
            }
        }

        private double ParseNumber(string text)
        {
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out var value))
                return value;

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                return value;

            return 0;
        }

        private string FormatNumber(double value)
        {
            if (Math.Abs(value % 1) < 0.0001)
                return value.ToString("0", CultureInfo.CurrentCulture);

            return value.ToString("0.##", CultureInfo.CurrentCulture);
        }


        Color GetSoftColor(string ttn)
        {
            if (string.IsNullOrEmpty(ttn))
                ttn = "NO_TTN";

            int h = ttn.GetHashCode();

            byte r = (byte)(80 + (h & 0x7F));
            byte g = (byte)(80 + ((h >> 7) & 0x7F));
            byte b = (byte)(80 + ((h >> 14) & 0x7F));

            // 45 = прозрачность ~18%
            return Color.FromArgb(45, r, g, b);
        }

        private void MergeButton_Click(object sender, RoutedEventArgs e)
        {
            mergeEnabled = !mergeEnabled;

            MergeButton.Content = mergeEnabled ? "⇆ Объединено" : "⇆ Объединить";

            ApplyAllFilters();
        }

        private void RenderJvk()
        {
            JvkPanel.Children.Clear();

            if (!filteredJournal.Any())
            {
                if (JvkHeaderBorder != null)
                    JvkHeaderBorder.Visibility = Visibility.Collapsed;
                return;
            }

            if (JvkHeaderBorder != null)
                JvkHeaderBorder.Visibility = Visibility.Visible;

            // ===== авто размер колонок =====
            int maxName = filteredJournal.Max(j => j.MaterialName?.Length ?? 0);
            int maxPassport = filteredJournal.Max(j => j.Passport?.Length ?? 0);
            int maxSupplier = filteredJournal.Max(j => j.Supplier?.Length ?? 0);
            int maxTtn = filteredJournal.Max(j => j.Ttn?.Length ?? 0);

            int colTtn = Math.Max(140, maxTtn * 7);
            int colName = Math.Max(260, maxName * 7);
            int colStb = 90;
            int colUnit = 70;
            int colQty = 90;
            int colSupplier = Math.Max(220, maxSupplier * 7);
            int colPassport = Math.Max(260, maxPassport * 7);

            int maxTotalWidth = (int)Math.Max(900, JvkHeaderBorder?.ActualWidth > 0 ? JvkHeaderBorder.ActualWidth : ActualWidth - 180);
            int total = colTtn + colName + colStb + colUnit + colQty + colSupplier + colPassport;

            if (total > maxTotalWidth)
            {
                double overflow = total - maxTotalWidth;

                void shrink(ref int c, double factor)
                {
                    int reduce = (int)(overflow * factor);
                    c -= reduce;
                    if (c < 100) c = 100;
                }

                shrink(ref colName, 0.45);
                shrink(ref colPassport, 0.25);
                shrink(ref colSupplier, 0.20);
                shrink(ref colTtn, 0.10);
            }
            else if (total < maxTotalWidth)
            {
                var extra = maxTotalWidth - total;
                colName += (int)(extra * 0.45);
                colPassport += (int)(extra * 0.25);
                colSupplier += (int)(extra * 0.20);
                colTtn += (int)(extra * 0.10);
            }

            UpdateJvkHeaderColumns(colTtn, colName, colStb, colUnit, colQty, colSupplier, colPassport);

            var structured = filteredJournal
                .Where(j => j.Category == "Основные")
                .GroupBy(j => j.Date.Date)
                .OrderByDescending(g => g.Key);


            if (mergeEnabled)
            {
                var merged = structured
                    .Select(day => new
                    {
                        Date = day.Key,
                        Groups = day.GroupBy(x => x.MaterialGroup)
                            .Select(g =>
                            {
                                var ttns = string.Join(", ",
                                    g.Select(x => x.Ttn)
                                    .Where(x => !string.IsNullOrWhiteSpace(x))
                                    .Distinct());

                                var items = g.GroupBy(x => x.MaterialName)
                                    .Select(nn => new
                                    {
                                        Name = nn.Key,
                                        Qty = nn.Sum(x => x.Quantity),
                                        Unit = nn.First().Unit,
                                        Stb = string.Join(", ",
                                            nn.Select(x => x.Stb)
                                            .Where(x => !string.IsNullOrWhiteSpace(x))
                                            .Distinct()),
                                        Supplier = string.Join(", ",
                                            nn.Select(x => x.Supplier)
                                            .Where(x => !string.IsNullOrWhiteSpace(x))
                                            .Distinct()),
                                        Passport = string.Join(", ",
                                            nn.Select(x => x.Passport)
                                            .Where(x => !string.IsNullOrWhiteSpace(x))
                                            .Distinct())
                                    })
                                    .ToList();

                                return new { Group = g.Key, Ttn = ttns, Items = items };
                            })
                    })
                    .ToList();

                RenderMerged(merged, colTtn, colName, colStb, colUnit, colQty, colSupplier, colPassport);
                return;
            }



            foreach (var day in structured)
            {
                // Лёгкая горизонтальная разделительная линия между днями
                var daySeparator = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(180, 187, 198)), // тот же тон что в таблице
                    BorderThickness = new Thickness(0, 1, 0, 0),
                    Margin = new Thickness(0, 12, 0, 8) // чуть воздуха
                };

                JvkPanel.Children.Add(daySeparator);

                var dateHeader = new TextBlock
                {
                    Text = day.Key.ToString("dd.MM.yyyy"),
                    FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(0, 0, 0, 6),

                    FontSize = 15
                };

                JvkPanel.Children.Add(dateHeader);

                var ttnGroups = day.GroupBy(x => new { x.Ttn, x.MaterialGroup });

                foreach (var ttn in ttnGroups)
                {
                    if (mergeEnabled)
                    {
                        foreach (var grp in structured)
                        {
                            // рендерим дату
                            // рендерим grp.Groups как агрегированный грид
                        }
                    }
                    else
                    {
                        // старый вывод
                    }

                    var items = ttn.ToList();
                    int rows = items.Count;
                    bool stbSame = true;

                    for (int i = 1; i < items.Count; i++)
                    {
                        if (items[i].Stb != items[0].Stb)
                            stbSame = false;
                    }

                    string mergedStb = stbSame ? items[0].Stb : null;

                    bool unitSame = true;
                    bool supplierSame = true;

                    for (int i = 1; i < items.Count; i++)
                    {
                        if (items[i].Unit != items[0].Unit)
                            unitSame = false;

                        if (items[i].Supplier != items[0].Supplier)
                            supplierSame = false;
                    }

                    string mergedUnit = unitSame ? items[0].Unit : null;
                    string mergedSupplier = supplierSame ? items[0].Supplier : null;


                    var grid = new Grid { Margin = new Thickness(0, 0, 0, 4) };
                    var bg = new SolidColorBrush(GetSoftColor(ttn.Key.Ttn));

                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colTtn) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colName) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colStb) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colUnit) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colQty) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colSupplier) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colPassport) });

                    for (int i = 0; i < rows; i++)
                        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });





                    for (int r = 0; r < rows; r++)
                    {
                        var x = items[r];

                        string ttnVal = string.IsNullOrWhiteSpace(x.Ttn) ? "—" : x.Ttn;
                        string name = string.IsNullOrWhiteSpace(x.MaterialName) ? "—" : x.MaterialName;
                        string stb = string.IsNullOrWhiteSpace(x.Stb) ? "—" : x.Stb;
                        string unit = string.IsNullOrWhiteSpace(x.Unit) ? "—" : x.Unit;
                        string supplier = string.IsNullOrWhiteSpace(x.Supplier) ? "—" : x.Supplier;
                        string passport = string.IsNullOrWhiteSpace(x.Passport) ? "—" : x.Passport;
                        string qty = x.Quantity > 0 ? x.Quantity.ToString() : "—";

                        AddCell(grid, r, 0, ttnVal, bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 1, name, bg: bg, noWrap: true);
                        AddCell(grid, r, 2, stb, bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 3, unit, bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 4, qty, bg: bg, align: TextAlignment.Right);
                        AddCell(grid, r, 5, supplier, bg: bg, noWrap: true);
                        AddCell(grid, r, 6, passport, bg: bg, noWrap: true);
                    }

                    columnWidths["Ttn"] = colTtn;
                    columnWidths["Name"] = colName;
                    columnWidths["Stb"] = colStb;
                    columnWidths["Unit"] = colUnit;
                    columnWidths["Qty"] = colQty;
                    columnWidths["Supplier"] = colSupplier;
                    columnWidths["Passport"] = colPassport;

                    JvkPanel.Children.Add(grid);

                }
            }
        }
        
        void RenderMerged(
            IEnumerable<dynamic> merged,
            int colTtn, int colName, int colStb, int colUnit, int colQty, int colSupplier, int colPassport)
        {
            UpdateJvkHeaderColumns(colTtn, colName, colStb, colUnit, colQty, colSupplier, colPassport);

            foreach (var day in merged)
            {
                // ====== ДАТА ======
                var dateHeader = new TextBlock
                {
                    Text = day.Date.ToString("dd.MM.yyyy"),
                    FontWeight = FontWeights.SemiBold,
                    FontSize = 15,
                    Margin = new Thickness(0, 12, 0, 6)
                };

                JvkPanel.Children.Add(dateHeader);

                // ====== ТАБЛИЦА ДНЯ ======
                var dayGrid = new Grid
                {
                    Margin = new Thickness(0, 0, 0, 6)
                };

                dayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colTtn) });
                dayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colName) });
                dayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colStb) });
                dayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colUnit) });
                dayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colQty) });
                dayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colSupplier) });
                dayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colPassport) });

                int rowIndex = 0;
                foreach (var grp in day.Groups)
                {
                    var items = ((IEnumerable<dynamic>)grp.Items).ToList();
                    int start = rowIndex;
                    int rows = items.Count;
                    // === АГРЕГАЦИЯ СТБ ===
                    var stbRaw = items
                        .Select(x => Normalize(x.Stb))
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct()
                        .ToList();

                    string stbMerged = stbRaw.Count == 0 ? "—"
                                     : stbRaw.Count == 1 ? stbRaw[0]
                                     : string.Join(", ", stbRaw);


                    // === АГРЕГАЦИЯ UNIT ===
                    var unitRaw = items
                        .Select(x => Normalize(x.Unit))
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct()
                        .ToList();

                    string unitMerged = unitRaw.Count == 0 ? "—"
                                      : unitRaw.Count == 1 ? unitRaw[0]
                                      : string.Join(", ", unitRaw);


                    // === АГРЕГАЦИЯ SUPPLIER ===
                    var supplierRaw = items
                        .Select(x => Normalize(x.Supplier))
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct()
                        .ToList();

                    string supplierMerged = supplierRaw.Count == 0 ? "—"
                                          : supplierRaw.Count == 1 ? supplierRaw[0]
                                          : string.Join(", ", supplierRaw);

                    bool stbSame = true;

                    for (int i = 1; i < items.Count; i++)
                    {
                        if (items[i].Stb != items[0].Stb)
                            stbSame = false;
                    }

                    string mergedStb = stbSame ? items[0].Stb : null;

                    var bg = new SolidColorBrush(GetSoftColor(grp.Ttn ?? ""));

                    // UNIT + SUPPLIER анализ
                    bool unitSame = true;
                    bool supplierSame = true;

                    for (int i = 1; i < items.Count; i++)
                    {
                        if (items[i].Unit != items[0].Unit)
                            unitSame = false;

                        if (items[i].Supplier != items[0].Supplier)
                            supplierSame = false;
                    }


                    string mergedUnit = unitSame ? items[0].Unit : null;
                    string mergedSupplier = supplierSame ? items[0].Supplier : null;

                    // ===== ТТН ОДИН РАЗ =====
                    AddCell(dayGrid, rowIndex, 0, grp.Ttn ?? "", rowspan: rows, bg: bg, align: TextAlignment.Center);


                    // === АГРЕГАЦИЯ ПАСПОРТОВ ===
                    var passportsRaw = items
                        .Select(x => (x.Passport ?? "").Trim())
                        .ToList();

                    var nonEmpty = passportsRaw
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct()
                        .ToList();

                    string passportMerged;

                    if (nonEmpty.Count == 0)
                        passportMerged = "—";
                    else if (nonEmpty.Count == 1)
                        passportMerged = nonEmpty[0];
                    else
                        passportMerged = string.Join(", ", nonEmpty);


                    foreach (var x in items)
                    {
                        dayGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                        string name = string.IsNullOrWhiteSpace(x.Name) ? "—" : x.Name;
                        string stb = string.IsNullOrWhiteSpace(x.Stb) ? "—" : x.Stb;
                        string unit = string.IsNullOrWhiteSpace(x.Unit) ? "—" : x.Unit;
                        string supplier = string.IsNullOrWhiteSpace(x.Supplier) ? "—" : x.Supplier;
                        string passport = string.IsNullOrWhiteSpace(x.Passport) ? "—" : x.Passport;
                        string qty = x.Qty > 0 ? x.Qty.ToString() : "—";

                        AddCell(dayGrid, rowIndex, 1, name, bg: bg, noWrap: true);
                        
                        AddCell(dayGrid, rowIndex, 4, qty, bg: bg, align: TextAlignment.Right);
        


                        rowIndex++;
                    }
                    AddCell(dayGrid, start, 2, stbMerged, rowspan: rows, bg: bg, align: TextAlignment.Center);
                    AddCell(dayGrid, start, 3, unitMerged, rowspan: rows, bg: bg, align: TextAlignment.Center);
                    AddCell(dayGrid, start, 5, supplierMerged, rowspan: rows, bg: bg, noWrap: true);
                    AddCell(dayGrid, start, 6, passportMerged, rowspan: rows, bg: bg, noWrap: true);


                    // пустой отступ между группами
                    dayGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(6) });
                    rowIndex++;
                }



                JvkPanel.Children.Add(dayGrid);
            }
        }


        private void UpdateJvkHeaderColumns(
            int colTtn, int colName, int colStb, int colUnit, int colQty, int colSupplier, int colPassport)
        {
            if (JvkHeaderGrid == null)
                return;

            if (JvkHeaderGrid.ColumnDefinitions.Count < 7)
                return;

            JvkHeaderGrid.ColumnDefinitions[0].Width = new GridLength(colTtn);
            JvkHeaderGrid.ColumnDefinitions[1].Width = new GridLength(colName);
            JvkHeaderGrid.ColumnDefinitions[2].Width = new GridLength(colStb);
            JvkHeaderGrid.ColumnDefinitions[3].Width = new GridLength(colUnit);
            JvkHeaderGrid.ColumnDefinitions[4].Width = new GridLength(colQty);
            JvkHeaderGrid.ColumnDefinitions[5].Width = new GridLength(colSupplier);
            JvkHeaderGrid.ColumnDefinitions[6].Width = new GridLength(colPassport);
        }

        public List<JournalRecord> GetJournal()
        {
            return journal;
        }
        private void ArrivalGroups_Toggle(object sender, MouseButtonEventArgs e)
        {
            var item = ((FrameworkElement)e.OriginalSource).DataContext as string;
            if (item == null) return;

        }

        private void ArrivalNames_Toggle(object sender, MouseButtonEventArgs e)
        {
            var item = ((FrameworkElement)e.OriginalSource).DataContext as string;
            if (item == null) return;



            ApplyAllFilters();
        }
        private HashSet<string> selectedArrivalTypes = new();
        private HashSet<string> selectedArrivalNames = new();

        private void RefreshArrivalTypes()
        {
            ArrivalTypesPanel.Children.Clear();

            var groups = journal
                .Select(j => j.MaterialGroup)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct()
                .OrderBy(x => x);

            foreach (var g in groups)
            {
                var chip = new ToggleButton
                {
                    Content = g,
                    Tag = g,
                    Style = (Style)FindResource("ChipToggle")
                };

                chip.IsChecked = selectedArrivalTypes.Contains(g);

                chip.Checked += (_, _) =>
                {
                    if (arrivalMatrixMode)
                        selectedArrivalTypes.Clear();
                    selectedArrivalTypes.Add(g);
                    if (arrivalMatrixMode)
                    {
                        selectedArrivalNames.Clear();
                        RefreshArrivalTypes();
                    }
                    RefreshArrivalNames();
                    ApplyAllFilters();
                };
                chip.Unchecked += (_, _) =>
                {
                    if (selectedArrivalTypes.Contains(g))
                        selectedArrivalTypes.Remove(g);
                    if (arrivalMatrixMode)
                        selectedArrivalNames.Clear();
                    RefreshArrivalNames();
                    ApplyAllFilters();
                };

                ArrivalTypesPanel.Children.Add(chip);
            }
        }
        private void RefreshArrivalNames()
        {
            ArrivalNamesPanel.Children.Clear();

            var names = journal
                .Where(j => selectedArrivalTypes.Count == 0 || selectedArrivalTypes.Contains(j.MaterialGroup))
                .Select(j => j.MaterialName)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct()
                .OrderBy(x => x);

            foreach (var n in names)
            {
                var chip = new ToggleButton
                {
                    Content = n,
                    Tag = n,
                    Style = (Style)FindResource("ChipToggle")
                };

                chip.IsChecked = selectedArrivalNames.Contains(n);

                chip.Checked += (_, _) =>
                {
                    selectedArrivalNames.Add(n);
                    ApplyAllFilters();
                };
                chip.Unchecked += (_, _) =>
                {
                    selectedArrivalNames.Remove(n);
                    ApplyAllFilters();
                };

                ArrivalNamesPanel.Children.Add(chip);
            }
        }
        private void ArrivalSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyAllFilters();
        }
        private void ExportArrival_Click(object sender, RoutedEventArgs e)
        {
            if (!filteredJournal.Any())
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = "Приход.xlsx"
            };

            if (dlg.ShowDialog() != true)
                return;

            using (var wb = new XLWorkbook())
            {
                ExportArrival(wb);
                wb.SaveAs(dlg.FileName);
            }


            MessageBox.Show("Экспорт завершён");
        }
        void ExportArrival(IXLWorkbook wb)
        {
            // получаем уникальные группы
            var groups = filteredJournal
                .Where(j => !string.IsNullOrWhiteSpace(j.MaterialGroup))
                .Select(j => j.MaterialGroup)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            foreach (var group in groups)
            {
                // создаём лист с именем группы
                var ws = wb.Worksheets.Add(group);

                int row = 1;

                // заголовок
                ws.Cell(row, 1).Value = "Дата";
                ws.Cell(row, 2).Value = "Тип";
                ws.Cell(row, 3).Value = "Наименование";
                ws.Cell(row, 4).Value = "Ед.";
                ws.Cell(row, 5).Value = "Кол-во";
                ws.Cell(row, 6).Value = "ТТН";
                ws.Cell(row, 7).Value = "Поставщик";
                ws.Cell(row, 8).Value = "Паспорт";

                ws.Range(row, 1, row, 8).Style.Font.Bold = true;
                ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = XLColor.FromHtml("#E9EEF6");
                row++;

                // строки только этого типа
                var data = filteredJournal
                    .Where(j => j.MaterialGroup == group)
                    .OrderByDescending(j => j.Date);

                foreach (var rec in data)
                {
                    ws.Cell(row, 1).Value = rec.Date.ToString("dd.MM.yyyy");
                    ws.Cell(row, 2).Value = rec.MaterialGroup;
                    ws.Cell(row, 3).Value = rec.MaterialName;
                    ws.Cell(row, 4).Value = rec.Unit;
                    ws.Cell(row, 5).Value = rec.Quantity;
                    ws.Cell(row, 6).Value = rec.Ttn;
                    ws.Cell(row, 7).Value = rec.Supplier;
                    ws.Cell(row, 8).Value = rec.Passport;
                    row++;
                }

                ws.Columns().AdjustToContents();
                ws.Range(1, 1, row - 1, 8).SetAutoFilter();
            }
        }




    }

}
