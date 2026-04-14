using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Documents;
using System.Windows.Interop;
using System.Windows.Navigation;
using System.Windows.Threading;
using WpfPath = System.Windows.Shapes.Path;
using System.Text.RegularExpressions;
using IOPath = System.IO.Path;

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
        private DateTime? lastSuccessfulSaveLocalTime;
        private string lastOperationStatusText = "Готово";

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
        private readonly ObservableCollection<string> otStatusFilters = new();
        private readonly ObservableCollection<string> otSpecialtyFilters = new();
        private readonly ObservableCollection<string> otBrigadeFilters = new();
        private string selectedOtStatusFilter = "Все";
        private string selectedOtSpecialtyFilter = "Все";
        private string selectedOtBrigadeFilter = "Все";
        private bool suppressOtFilterSelectionChange;
        private bool isTreePinned;
        private bool isOtToolsPinned;
        private bool isTimesheetToolsPinned;
        private bool isProductionToolsPinned;
        private bool isInspectionToolsPinned;
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
        private readonly ObservableCollection<string> productionAutoFillProfileNames = new();
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
        private readonly ObservableCollection<DocumentTreeNode> openPdfTabs = new();
        private readonly ObservableCollection<DocumentTreeNode> openEstimateTabs = new();
        private DocumentTreeNode activePdfTab;
        private DocumentTreeNode activeEstimateTab;
        private DocumentTreeNode secondaryPdfTab;
        private DocumentTreeNode secondaryEstimateTab;
        private DocumentTreeNode pdfDraggedTab;
        private DocumentTreeNode estimateDraggedTab;
        private bool pdfSplitEnabled;
        private bool estimateSplitEnabled;
        private Process pdfEditorProcess;
        private IntPtr pdfEditorWindowHandle = IntPtr.Zero;
        private int pdfEmbeddedWindowX = int.MinValue;
        private int pdfEmbeddedWindowY = int.MinValue;
        private int pdfEmbeddedWindowWidth = -1;
        private int pdfEmbeddedWindowHeight = -1;
        private string pdfEmbeddedFilePath = string.Empty;
        private string preferredPdfEditorPath = string.Empty;
        private bool useExternalPdfEditor;
        private Window pdfDetachedWindow;
        private WebBrowser pdfDetachedBrowser;
        private DocumentTreeNode pdfDetachedNode;
        private Point pdfTabDragStart;
        private Point estimateTabDragStart;
        private string pdfPreviewCurrentPath = string.Empty;
        private string estimatePreviewCurrentPath = string.Empty;
        private string pdfSecondaryPreviewCurrentPath = string.Empty;
        private string estimateSecondaryPreviewCurrentPath = string.Empty;
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
        private readonly DispatcherTimer estimateLayoutGuardTimer;
        private readonly DispatcherTimer autoSaveTimer;
        private readonly DispatcherTimer gridPreferenceSaveDebounceTimer;
        private DateTime? reminderSnoozedUntil;
        private bool timesheetNeedsRebuild;
        private bool reminderRefreshRequested;
        private bool timesheetRebuildRequested;
        private bool timesheetRebuildForceRequested;
        private bool timesheetOtSyncDirty = true;
        private bool isSyncingTimesheetToOt;
        private const int MaxTimesheetMissingDocs = 3;
        private ReminderOverlayWindow reminderOverlayWindow;
        private object estimateExcelApplication;
        private object estimateExcelWorkbook;
        private Process estimateSpreadsheetProcess;
        private IntPtr estimateExcelWindowHandle = IntPtr.Zero;
        private int estimateEmbeddedWindowX = int.MinValue;
        private int estimateEmbeddedWindowY = int.MinValue;
        private int estimateEmbeddedWindowWidth = -1;
        private int estimateEmbeddedWindowHeight = -1;
        private string estimateEmbeddedFilePath = string.Empty;
        private Process estimateSpreadsheetProcessSecondary;
        private IntPtr estimateExcelWindowHandleSecondary = IntPtr.Zero;
        private int estimateEmbeddedWindowXSecondary = int.MinValue;
        private int estimateEmbeddedWindowYSecondary = int.MinValue;
        private int estimateEmbeddedWindowWidthSecondary = -1;
        private int estimateEmbeddedWindowHeightSecondary = -1;
        private string estimateEmbeddedFilePathSecondary = string.Empty;
        private bool estimateDetached;
        private bool estimateSecondaryDetached;
        private ExternalSpreadsheetInstance activeExternalSpreadsheetInstance;
        private ExternalSpreadsheetInstance activeExternalSpreadsheetInstanceSecondary;
        private readonly Dictionary<string, ExternalSpreadsheetInstance> externalSpreadsheetInstances = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ExternalSpreadsheetInstance> externalSpreadsheetInstancesSecondary = new(StringComparer.OrdinalIgnoreCase);
        private string preferredSpreadsheetEditorPath = string.Empty;
        private bool useExternalSpreadsheetEditor;
        private bool previewWarmupStarted;
        private readonly Dictionary<string, string> documentHashPathCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> lastStorageIntegrityIssues = new();
        private readonly ObservableCollection<OperationLogEntry> operationLogEntries = new();
        private string lastSavedStateSnapshot = string.Empty;
        private bool closeConfirmed;
        private FileStream projectLockStream;
        private string projectLockFilePath = string.Empty;
        private string sessionMarkerFilePath = string.Empty;
        private bool previousSessionCrashed;
        private const int MaxAutoBackupFiles = 20;
        private const int MaxChangeLogEntries = 5000;
        private const int MaxSummaryMatrixCacheEntries = 16;
        private const int EmbeddedPdfTopTrim = 0;
        private bool timesheetInitialized;
        private bool productionJournalInitialized;
        private bool inspectionJournalInitialized;
        private bool productionLookupsDirty = true;
        private bool inspectionLookupsDirty = true;
        private bool productionStateDirty = true;
        private bool inspectionStateDirty = true;
        private static readonly HttpClient UpdateHttpClient = new();
        private bool isRefreshingProductionProfileSelection;
        private int summaryDataVersion = 1;
        private readonly Dictionary<string, SummaryMatrixCacheEntry> summaryMatrixCache = new(StringComparer.Ordinal);
        private readonly Dictionary<string, TimeSpan> tabOpenDiagnostics = new(StringComparer.CurrentCultureIgnoreCase);
        private readonly Dictionary<string, List<string>> tabReminderMessages = new(StringComparer.CurrentCultureIgnoreCase);
        private readonly DispatcherTimer arrivalFilterDebounceTimer;
        private readonly DispatcherTimer otSearchDebounceTimer;
        private readonly DispatcherTimer summaryRefreshDebounceTimer;
        private bool arrivalFilterRefreshRequested;
        private bool otSearchRefreshRequested;
        private bool summaryRefreshRequested;
        private int summaryRefreshRequestVersion;
        private int processingOverlayDepth;
        private readonly DispatcherTimer processingOverlayDelayTimer;
        private string processingOverlayPendingText = "Идет обработка...";
        private bool isApplyingColumnPreferences;
        private bool pendingGridPreferenceSave;
        private bool isOpeningCommandDialog;
        private readonly ObservableCollection<string> arrivalFilterTemplateNames = new();
        private bool suppressArrivalTemplateSelectionChange;
        private static readonly Regex DigitsInputRegex = new(@"^\d+$", RegexOptions.Compiled);
        private const string ActiveTabModeKey = "__active_tab";

        private const string GridPrefArrival = "tab_arrival_legacy_grid";
        private const string GridPrefOt = "tab_ot_grid";
        private const string GridPrefTimesheet = "tab_timesheet_grid";
        private const string GridPrefProduction = "tab_production_grid";
        private const string GridPrefInspection = "tab_inspection_grid";

        private sealed class SummaryMatrixCacheEntry
        {
            public List<SummaryMatrixGroupData> Groups { get; set; } = new();
        }

        private sealed class SummaryMatrixGroupData
        {
            public string GroupName { get; set; } = string.Empty;
            public List<SummaryMatrixRowData> Rows { get; set; } = new();
        }

        private sealed class SummaryMatrixRowData
        {
            public string MaterialName { get; set; } = string.Empty;
            public string Unit { get; set; } = string.Empty;
            public string Position { get; set; } = string.Empty;
            public double TotalArrival { get; set; }
        }

        private sealed class ExternalSpreadsheetInstance
        {
            public string FilePath { get; set; } = string.Empty;
            public Process Process { get; set; }
            public IntPtr Handle { get; set; }
        }

        private sealed class ColumnManagerRow : INotifyPropertyChanged
        {
            private bool isVisible = true;
            private string widthText = string.Empty;
            private int order;

            public string Header { get; set; } = string.Empty;
            public bool IsVisible
            {
                get => isVisible;
                set
                {
                    if (isVisible == value)
                        return;

                    isVisible = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsVisible)));
                }
            }

            public string WidthText
            {
                get => widthText;
                set
                {
                    if (string.Equals(widthText, value, StringComparison.CurrentCulture))
                        return;

                    widthText = value ?? string.Empty;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(WidthText)));
                }
            }

            public int Order
            {
                get => order;
                set
                {
                    if (order == value)
                        return;

                    order = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Order)));
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private sealed class OtInstructionReferenceRow : INotifyPropertyChanged
        {
            private string profession = string.Empty;
            private string instructionNumbers = string.Empty;

            public string Profession
            {
                get => profession;
                set
                {
                    var normalized = value?.Trim() ?? string.Empty;
                    if (string.Equals(profession, normalized, StringComparison.CurrentCulture))
                        return;
                    profession = normalized;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Profession)));
                }
            }

            public string InstructionNumbers
            {
                get => instructionNumbers;
                set
                {
                    var normalized = value?.Trim() ?? string.Empty;
                    if (string.Equals(instructionNumbers, normalized, StringComparison.CurrentCulture))
                        return;
                    instructionNumbers = normalized;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(InstructionNumbers)));
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private sealed class ProductionDeviationReferenceRow : INotifyPropertyChanged
        {
            private string materialType = string.Empty;
            private string deviation = string.Empty;

            public string MaterialType
            {
                get => materialType;
                set
                {
                    var normalized = value?.Trim() ?? string.Empty;
                    if (string.Equals(materialType, normalized, StringComparison.CurrentCulture))
                        return;
                    materialType = normalized;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MaterialType)));
                }
            }

            public string Deviation
            {
                get => deviation;
                set
                {
                    var normalized = value?.Trim() ?? string.Empty;
                    if (string.Equals(deviation, normalized, StringComparison.CurrentCulture))
                        return;
                    deviation = normalized;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Deviation)));
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private sealed class GlobalSearchResult
        {
            public string TabHeader { get; set; } = string.Empty;
            public string Title { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
            public Action NavigateAction { get; set; }
        }

        private sealed class CommandPaletteAction : INotifyPropertyChanged
        {
            private string shortcut = string.Empty;

            public string Id { get; set; } = string.Empty;
            public string DefaultShortcut { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string Shortcut
            {
                get => shortcut;
                set
                {
                    var normalized = NormalizeShortcutText(value);
                    if (string.Equals(shortcut, normalized, StringComparison.CurrentCultureIgnoreCase))
                        return;

                    shortcut = normalized;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Shortcut)));
                }
            }

            public string Hint { get; set; } = string.Empty;
            public Action ExecuteAction { get; set; } = () => { };
            public event PropertyChangedEventHandler PropertyChanged;
        }

        private enum SaveTrigger
        {
            Manual,
            Auto,
            System
        }

        private const int GWL_STYLE = -16;
        private const int GWL_EXSTYLE = -20;
        private const int GWLP_HWNDPARENT = -8;
        private const long WS_CHILD = 0x40000000L;
        private const long WS_CAPTION = 0x00C00000L;
        private const long WS_DLGFRAME = 0x00400000L;
        private const long WS_THICKFRAME = 0x00040000L;
        private const long WS_MINIMIZEBOX = 0x00020000L;
        private const long WS_MAXIMIZEBOX = 0x00010000L;
        private const long WS_SYSMENU = 0x00080000L;
        private const long WS_POPUP = unchecked((int)0x80000000);
        private const long WS_EX_APPWINDOW = 0x00040000L;
        private const long WS_EX_TOOLWINDOW = 0x00000080L;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const uint SWP_NOOWNERZORDER = 0x0200;
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;
        private const int EmbeddedExcelTopTrim = 0;
        private const uint WM_CANCELMODE = 0x001F;
        private const byte VK_SHIFT = 0x10;
        private const byte VK_CONTROL = 0x11;
        private const byte VK_MENU = 0x12;
        private const byte VK_ESCAPE = 0x1B;
        private const byte VK_LSHIFT = 0xA0;
        private const byte VK_RSHIFT = 0xA1;
        private const byte VK_LCONTROL = 0xA2;
        private const byte VK_RCONTROL = 0xA3;
        private const byte VK_LMENU = 0xA4;
        private const byte VK_RMENU = 0xA5;
        private const uint KEYEVENTF_KEYUP = 0x0002;

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

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

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

        private sealed class SummaryComparisonRow
        {
            public string Тип { get; set; } = string.Empty;
            public string Наименование { get; set; } = string.Empty;
            public string Ед { get; set; } = string.Empty;
            public double План { get; set; }
            public double Пришло { get; set; }
            public double Смонтировано { get; set; }
            public double Остаток { get; set; }
        }

        private sealed class SummaryBalanceEditorRow : INotifyPropertyChanged
        {
            private string reason = string.Empty;
            public string Category { get; set; } = string.Empty;
            public string Group { get; set; } = string.Empty;
            public string Material { get; set; } = string.Empty;
            public string Unit { get; set; } = string.Empty;
            public double Quantity { get; set; }
            public bool IsOverage { get; set; }
            public string Scenario => IsOverage ? "Излишек" : "Дефицит";

            public string Reason
            {
                get => reason;
                set
                {
                    if (string.Equals(reason, value, StringComparison.CurrentCulture))
                        return;
                    reason = value ?? string.Empty;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Reason)));
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private sealed class DocumentLibraryReportRow
        {
            public string Library { get; set; } = string.Empty;
            public string NodeName { get; set; } = string.Empty;
            public string NodeType { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
        }

        private sealed class DocumentTreeSearchRow
        {
            public string Library { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string NodeType { get; set; } = string.Empty;
            public string Path { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public DocumentTreeNode Node { get; set; }
        }

        private sealed class DocumentStorageManifestEntry
        {
            public string RelativePath { get; set; } = string.Empty;
            public string Hash { get; set; } = string.Empty;
            public long Size { get; set; }
        }

        private sealed class DocumentIntegrityIssueRow
        {
            public string NodeName { get; set; } = string.Empty;
            public string Path { get; set; } = string.Empty;
            public string Issue { get; set; } = string.Empty;
        }

        private sealed class DocumentStorageMaintenanceResult
        {
            public int NodesVisited { get; set; }
            public int PathsRecovered { get; set; }
            public int MetadataUpdated { get; set; }
            public int LinksRepointed { get; set; }
            public int HashIndexSize { get; set; }
            public int DuplicateFilesRemoved { get; set; }
            public int OrphanFilesRemoved { get; set; }
            public List<string> Errors { get; } = new();
        }

        private sealed class OperationLogEntry
        {
            public DateTime TimestampLocal { get; set; }
            public string Kind { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string Details { get; set; } = string.Empty;
        }

        private sealed class DiagnosticsMetricRow
        {
            public string Metric { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
        }

        private sealed class BoolToScenarioConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
                => value is bool flag && flag ? "Излишек" : "Дефицит";

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
                => Binding.DoNothing;
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
        public ObservableCollection<string> ProductionAutoFillProfileNames => productionAutoFillProfileNames;
        public ObservableCollection<string> InspectionJournalNames => inspectionJournalNames;
        public ObservableCollection<string> InspectionNames => inspectionNames;
        public MainWindow()
        {
            InitializeComponent();
            if (OtJournalGrid != null)
            {
                OtJournalGrid.RowHeight = double.NaN;
                OtJournalGrid.MinRowHeight = 42;
            }
            if (PdfPreviewBrowser != null)
                PdfPreviewBrowser.LoadCompleted += DocumentPreviewBrowser_LoadCompleted;
            if (PdfPreviewBrowserSecondary != null)
                PdfPreviewBrowserSecondary.LoadCompleted += DocumentPreviewBrowser_LoadCompleted;
            if (EstimatePreviewBrowser != null)
                EstimatePreviewBrowser.LoadCompleted += DocumentPreviewBrowser_LoadCompleted;
            currentSaveFileName = ResolveDefaultSavePath();
            InitializeSpreadsheetEditorPreference();
            InitializePdfEditorPreference();
            SizeChanged += MainWindow_SizeChanged;
            LocationChanged += MainWindow_LocationChanged;
            StateChanged += MainWindow_StateChanged;
            Activated += MainWindow_Activated;
            Deactivated += MainWindow_Deactivated;
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

            estimateLayoutGuardTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(650)
            };
            estimateLayoutGuardTimer.Tick += EstimateLayoutGuardTimer_Tick;
            estimateLayoutGuardTimer.Start();
            arrivalFilterDebounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(220)
            };
            arrivalFilterDebounceTimer.Tick += ArrivalFilterDebounceTimer_Tick;
            otSearchDebounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(180)
            };
            otSearchDebounceTimer.Tick += OtSearchDebounceTimer_Tick;
            summaryRefreshDebounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(220)
            };
            summaryRefreshDebounceTimer.Tick += SummaryRefreshDebounceTimer_Tick;
            processingOverlayDelayTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(120)
            };
            processingOverlayDelayTimer.Tick += ProcessingOverlayDelayTimer_Tick;
            autoSaveTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(5)
            };
            autoSaveTimer.Tick += AutoSaveTimer_Tick;
            autoSaveTimer.Start();
            gridPreferenceSaveDebounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(900)
            };
            gridPreferenceSaveDebounceTimer.Tick += GridPreferenceSaveDebounceTimer_Tick;
            // ===== БЛОКИРОВКА ВКЛЮЧЕНА ПО УМОЛЧАНИЮ =====
            isLocked = true;

            if (!TryAcquireProjectLock(currentSaveFileName))
            {
                closeConfirmed = true;
                Dispatcher.BeginInvoke(new Action(Close), DispatcherPriority.ApplicationIdle);
                return;
            }

            previousSessionCrashed = CheckIfPreviousSessionCrashed(currentSaveFileName, out var markerPath);
            TryRecoverFromCrashIfNeeded(previousSessionCrashed);
            WriteSessionMarker(markerPath);

            LoadState();
            InitializeOtJournal();
            filteredJournal = journal.ToList();

            ArrivalPanel.ArrivalAdded += OnArrivalAdded;

            PushUndo();
            UpdateUndoRedoButtons();

            if (currentObject != null)
                ArrivalPanel.SetObject(currentObject, journal);
            RefreshArrivalTypes();
            RefreshArrivalNames();
            if (ArrivalTemplateBox != null)
                ArrivalTemplateBox.ItemsSource = arrivalFilterTemplateNames;
            RefreshArrivalFilterTemplates();
            RefreshDocumentLibraries();
            UpdateArrivalViewMode();

            RefreshTreePreserveState();
            ApplyProjectUiSettings();
            EnsureSelectedTabInitialized();
            AttachGridPreferenceTracking();
            lastSavedStateSnapshot = BuildCurrentStateSnapshotJson();

        }

        private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateReminderOverlayPlacement();
            ScheduleEstimateEmbeddedLayout();
            ScheduleEstimateEmbeddedLayoutSecondary();
            SchedulePdfEmbeddedLayout();
        }

        private void MainWindow_LocationChanged(object sender, EventArgs e)
        {
            UpdateReminderOverlayPlacement();
            ScheduleEstimateEmbeddedLayout();
            ScheduleEstimateEmbeddedLayoutSecondary();
            SchedulePdfEmbeddedLayout();
        }

        private void EstimateLayoutGuardTimer_Tick(object sender, EventArgs e)
        {
            if (!ReferenceEquals(MainTabs?.SelectedItem, EstimateTab))
            {
                if (ReferenceEquals(MainTabs?.SelectedItem, PdfTab) && pdfEditorWindowHandle != IntPtr.Zero)
                    LayoutEmbeddedPdfWindow();
                return;
            }

            if (estimateExcelWindowHandle != IntPtr.Zero && EstimateExcelHost?.Visibility == Visibility.Visible)
                LayoutEmbeddedEstimateWindow();

            if (estimateExcelWindowHandleSecondary != IntPtr.Zero && estimateSplitEnabled)
                LayoutEmbeddedEstimateWindowSecondary();

            if (ReferenceEquals(MainTabs?.SelectedItem, PdfTab) && pdfEditorWindowHandle != IntPtr.Zero)
                LayoutEmbeddedPdfWindow();
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

            var persistentPath = GetPersistentDefaultSavePath();
            if (File.Exists(persistentPath))
                return persistentPath;

            var candidates = new List<string>();
            AddCandidate(candidates, GetLegacyPersistentDefaultSavePath());
            AddCandidate(candidates, System.IO.Path.Combine(Environment.CurrentDirectory, DefaultSaveFileName));
            AddCandidate(candidates, System.IO.Path.Combine(AppContext.BaseDirectory, DefaultSaveFileName));

            var scanDir = new DirectoryInfo(AppContext.BaseDirectory);
            for (var i = 0; i < 6 && scanDir != null; i++, scanDir = scanDir.Parent)
            {
                AddCandidate(candidates, System.IO.Path.Combine(scanDir.FullName, DefaultSaveFileName));
            }

            var existing = candidates.FirstOrDefault(File.Exists);
            if (string.IsNullOrWhiteSpace(existing))
                return persistentPath;

            try
            {
                var persistentFolder = System.IO.Path.GetDirectoryName(persistentPath);
                if (!string.IsNullOrWhiteSpace(persistentFolder))
                    Directory.CreateDirectory(persistentFolder);

                File.Copy(existing, persistentPath, overwrite: false);
                CopyDirectorySafe(BuildStorageRootPath(existing), BuildStorageRootPath(persistentPath));
                CopyDirectorySafe(BuildLegacyStorageRootPath(existing), BuildStorageRootPath(persistentPath));
                return persistentPath;
            }
            catch
            {
                return existing;
            }
        }

        private static string GetDefaultDataRootPath()
        {
            var appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var root = System.IO.Path.Combine(appDataFolder, "ConstructionControl", "Data");
            Directory.CreateDirectory(root);
            return root;
        }

        private static string NormalizeDataRootPath(string dataRootPath)
        {
            var candidate = string.IsNullOrWhiteSpace(dataRootPath)
                ? GetDefaultDataRootPath()
                : Environment.ExpandEnvironmentVariables(dataRootPath.Trim());

            try
            {
                var fullPath = System.IO.Path.GetFullPath(candidate);
                Directory.CreateDirectory(fullPath);
                return fullPath;
            }
            catch
            {
                return GetDefaultDataRootPath();
            }
        }

        private string GetConfiguredDataRootPath()
        {
            var rawPath = currentObject?.UiSettings?.DataRootDirectory;
            return NormalizeDataRootPath(rawPath);
        }

        private static string GetPersistentDefaultSavePath()
        {
            var dataRoot = GetDefaultDataRootPath();
            return System.IO.Path.Combine(dataRoot, DefaultSaveFileName);
        }

        private static string GetLegacyPersistentDefaultSavePath()
        {
            var appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var appFolder = System.IO.Path.Combine(appDataFolder, "ConstructionControl");
            Directory.CreateDirectory(appFolder);
            return System.IO.Path.Combine(appFolder, DefaultSaveFileName);
        }

        private static void CopyDirectorySafe(string sourceRoot, string targetRoot)
        {
            if (string.IsNullOrWhiteSpace(sourceRoot) || string.IsNullOrWhiteSpace(targetRoot))
                return;

            if (!Directory.Exists(sourceRoot))
                return;

            Directory.CreateDirectory(targetRoot);
            foreach (var sourceFile in Directory.EnumerateFiles(sourceRoot, "*", SearchOption.AllDirectories))
            {
                var relative = System.IO.Path.GetRelativePath(sourceRoot, sourceFile);
                var targetFile = System.IO.Path.Combine(targetRoot, relative);
                var targetFolder = System.IO.Path.GetDirectoryName(targetFile);
                if (!string.IsNullOrWhiteSpace(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                File.Copy(sourceFile, targetFile, overwrite: true);
            }
        }

        private static string BuildProjectInstanceKey(string saveFileName)
        {
            var normalized = string.IsNullOrWhiteSpace(saveFileName)
                ? DefaultSaveFileName
                : System.IO.Path.GetFullPath(saveFileName).Trim().ToLowerInvariant();

            using var sha = SHA256.Create();
            var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(normalized));
            return string.Concat(bytes.Select(x => x.ToString("x2"))).Substring(0, 24);
        }

        private static string BuildStorageRootPath(string saveFileName)
        {
            if (string.IsNullOrWhiteSpace(saveFileName))
                return string.Empty;

            var dataRoot = GetDefaultDataRootPath();
            var root = System.IO.Path.Combine(dataRoot, "storage", BuildProjectInstanceKey(saveFileName));
            Directory.CreateDirectory(root);
            return root;
        }

        private static string BuildLegacyStorageRootPath(string saveFileName)
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

        private string BuildStorageRootPathForCurrentSettings(string saveFileName, bool createIfMissing)
        {
            if (string.IsNullOrWhiteSpace(saveFileName))
                return string.Empty;

            var dataRoot = GetConfiguredDataRootPath();
            var storageRoot = System.IO.Path.Combine(dataRoot, "storage", BuildProjectInstanceKey(saveFileName));
            if (createIfMissing)
                Directory.CreateDirectory(storageRoot);

            return storageRoot;
        }

        private string BuildRuntimeRootPathForCurrentSettings(string saveFileName, bool createIfMissing)
        {
            if (string.IsNullOrWhiteSpace(saveFileName))
                return string.Empty;

            var dataRoot = GetConfiguredDataRootPath();
            var runtimeRoot = System.IO.Path.Combine(dataRoot, "runtime", BuildProjectInstanceKey(saveFileName));
            if (createIfMissing)
                Directory.CreateDirectory(runtimeRoot);

            return runtimeRoot;
        }

        private string GetProjectRuntimeDirectory(string saveFileName)
            => BuildRuntimeRootPathForCurrentSettings(saveFileName, createIfMissing: true);

        private string GetAutoBackupDirectory(string saveFileName)
        {
            var projectRuntimeDir = GetProjectRuntimeDirectory(saveFileName);
            var backupDir = System.IO.Path.Combine(projectRuntimeDir, "autosave");
            Directory.CreateDirectory(backupDir);
            return backupDir;
        }

        private bool TryAcquireProjectLock(string saveFileName, bool showMessageOnFailure = true)
        {
            if (string.IsNullOrWhiteSpace(saveFileName))
                return false;

            var fullSavePath = System.IO.Path.GetFullPath(saveFileName);
            var lockPath = $"{fullSavePath}.lock";

            if (string.Equals(projectLockFilePath, lockPath, StringComparison.OrdinalIgnoreCase) && projectLockStream != null)
                return true;

            ReleaseProjectLock();

            try
            {
                var lockDirectory = System.IO.Path.GetDirectoryName(lockPath);
                if (!string.IsNullOrWhiteSpace(lockDirectory))
                    Directory.CreateDirectory(lockDirectory);

                projectLockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                projectLockFilePath = lockPath;

                projectLockStream.SetLength(0);
                using var writer = new StreamWriter(projectLockStream, Encoding.UTF8, 1024, leaveOpen: true);
                writer.WriteLine($"pid={Environment.ProcessId}");
                writer.WriteLine($"startedUtc={DateTime.UtcNow:O}");
                writer.Flush();
                projectLockStream.Flush(true);
                projectLockStream.Position = 0;
                return true;
            }
            catch
            {
                ReleaseProjectLock();
                if (showMessageOnFailure)
                {
                    MessageBox.Show(
                        "Этот файл проекта уже открыт в другом экземпляре программы. Закройте другой экземпляр и попробуйте снова.",
                        "Файл занят",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                return false;
            }
        }

        private void ReleaseProjectLock()
        {
            if (projectLockStream != null)
            {
                try
                {
                    projectLockStream.Dispose();
                }
                catch
                {
                    // ignore
                }
                projectLockStream = null;
            }

            if (!string.IsNullOrWhiteSpace(projectLockFilePath))
            {
                try
                {
                    if (File.Exists(projectLockFilePath))
                        File.Delete(projectLockFilePath);
                }
                catch
                {
                    // ignore
                }
            }

            projectLockFilePath = string.Empty;
        }

        private bool CheckIfPreviousSessionCrashed(string saveFileName, out string markerPath)
        {
            markerPath = System.IO.Path.Combine(GetProjectRuntimeDirectory(saveFileName), "session.marker");
            return File.Exists(markerPath);
        }

        private void WriteSessionMarker(string markerPath)
        {
            if (string.IsNullOrWhiteSpace(markerPath))
                return;

            var markerDirectory = System.IO.Path.GetDirectoryName(markerPath);
            if (!string.IsNullOrWhiteSpace(markerDirectory))
                Directory.CreateDirectory(markerDirectory);

            File.WriteAllText(markerPath, $"pid={Environment.ProcessId}{Environment.NewLine}startedUtc={DateTime.UtcNow:O}");
            sessionMarkerFilePath = markerPath;
        }

        private void RemoveSessionMarker()
        {
            if (string.IsNullOrWhiteSpace(sessionMarkerFilePath))
                return;

            try
            {
                if (File.Exists(sessionMarkerFilePath))
                    File.Delete(sessionMarkerFilePath);
            }
            catch
            {
                // ignore
            }
        }

        private void TryRecoverFromCrashIfNeeded(bool wasPreviousSessionCrashed)
        {
            if (!wasPreviousSessionCrashed)
                return;

            var latestBackup = GetLatestAutoBackupFile(currentSaveFileName);
            if (string.IsNullOrWhiteSpace(latestBackup) || !File.Exists(latestBackup))
                return;

            var hasCurrentState = TryReadStateSavedAtUtc(currentSaveFileName, out var currentSavedAtUtc);
            var hasBackupState = TryReadStateSavedAtUtc(latestBackup, out var backupSavedAtUtc);
            if (hasCurrentState && hasBackupState && currentSavedAtUtc >= backupSavedAtUtc)
                return;

            var backupTimeLocal = File.GetLastWriteTime(latestBackup);
            var answer = MessageBox.Show(
                $"Обнаружено некорректное завершение предыдущего сеанса.{Environment.NewLine}Найдено автосохранение от {backupTimeLocal:dd.MM.yyyy HH:mm:ss}.{Environment.NewLine}{Environment.NewLine}Восстановить его?",
                "Восстановление после сбоя",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (answer != MessageBoxResult.Yes)
                return;

            try
            {
                var backupJson = File.ReadAllText(latestBackup);
                if (!TryDeserializeAppState(backupJson, out var _, out _))
                    throw new InvalidDataException("Файл автосохранения повреждён.");

                if (File.Exists(currentSaveFileName))
                {
                    var rescuePath = $"{currentSaveFileName}.before_recovery_{DateTime.Now:yyyyMMdd_HHmmss}.bak";
                    File.Copy(currentSaveFileName, rescuePath, overwrite: true);
                }

                SaveStateJsonTransactional(backupJson, currentSaveFileName);
                MessageBox.Show("Данные восстановлены из автосохранения.", "Восстановление", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось восстановить автосохранение.{Environment.NewLine}{ex.Message}",
                    "Ошибка восстановления",
                    MessageBoxButton.OK,
                MessageBoxImage.Warning);
            }
        }

        private bool TryReadStateSavedAtUtc(string statePath, out DateTime savedAtUtc)
        {
            savedAtUtc = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(statePath) || !File.Exists(statePath))
                return false;

            try
            {
                var json = File.ReadAllText(statePath);
                if (!TryDeserializeAppState(json, out var state, out _)
                    || state == null
                    || state.SavedAtUtc == default)
                {
                    return false;
                }

                savedAtUtc = DateTime.SpecifyKind(state.SavedAtUtc, DateTimeKind.Utc);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void InitializeSpreadsheetEditorPreference()
        {
            preferredSpreadsheetEditorPath = ExternalToolPaths.ResolveSpreadsheetEditorPath(
                currentObject?.UiSettings?.PreferredSpreadsheetEditorPath);
            useExternalSpreadsheetEditor = !string.IsNullOrWhiteSpace(preferredSpreadsheetEditorPath);
        }

        private void InitializePdfEditorPreference()
        {
            preferredPdfEditorPath = ExternalToolPaths.ResolvePdfEditorPath(
                currentObject?.UiSettings?.PreferredPdfEditorPath);
            useExternalPdfEditor = !string.IsNullOrWhiteSpace(preferredPdfEditorPath);
        }

        private void StartPreviewWarmupAsync()
        {
            if (IsSafeStartupEnabled())
                return;

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
            if (useExternalSpreadsheetEditor)
                return;

            try
            {
                _ = EnsureEstimateExcelApplication();
            }
            catch
            {
                // Ignore Excel warmup errors to keep app startup resilient.
            }
        }

        private bool IsEstimateSpreadsheetProcessAlive()
        {
            if (estimateSpreadsheetProcess == null)
                return false;

            try
            {
                return !estimateSpreadsheetProcess.HasExited;
            }
            catch
            {
                return false;
            }
        }

        private bool IsEstimateSpreadsheetProcessAliveSecondary()
        {
            if (estimateSpreadsheetProcessSecondary == null)
                return false;

            try
            {
                return !estimateSpreadsheetProcessSecondary.HasExited;
            }
            catch
            {
                return false;
            }
        }

        private static IntPtr WaitForMainWindowHandle(Process process, int timeoutMs)
        {
            if (process == null)
                return IntPtr.Zero;

            var startedAt = Environment.TickCount;
            while (Environment.TickCount - startedAt < timeoutMs)
            {
                try
                {
                    process.Refresh();
                    var handle = process.MainWindowHandle;
                    if (handle != IntPtr.Zero)
                        return handle;

                    handle = FindTopLevelWindowForProcess(process.Id);
                    if (handle != IntPtr.Zero)
                        return handle;

                    if (process.HasExited)
                        return IntPtr.Zero;
                }
                catch
                {
                    return IntPtr.Zero;
                }

                Thread.Sleep(80);
            }

            return IntPtr.Zero;
        }

        private static IntPtr FindTopLevelWindowForProcess(int processId)
        {
            IntPtr found = IntPtr.Zero;
            EnumWindows((hWnd, lParam) =>
            {
                GetWindowThreadProcessId(hWnd, out var ownerPid);
                if (ownerPid != (uint)processId)
                    return true;

                if (!IsWindowVisible(hWnd))
                    return true;

                found = hWnd;
                return false;
            }, IntPtr.Zero);

            return found;
        }

        private void ShowEmbeddedEstimateInExternalEditor(string filePath)
        {
            if (string.IsNullOrWhiteSpace(preferredSpreadsheetEditorPath) || !File.Exists(preferredSpreadsheetEditorPath))
                throw new InvalidOperationException("PlanMaker не найден. Установите SoftMaker или укажите путь к PlanMaker.exe в настройках.");

            var normalizedPath = NormalizeDocumentPathKey(filePath);
            if (TryGetExternalSpreadsheetInstance(externalSpreadsheetInstances, normalizedPath, out var cached))
            {
                HideActiveExternalSpreadsheetInstance(isSecondary: false);
                ApplyExternalSpreadsheetInstance(cached, isSecondary: false);
                EstimateExcelHost.Visibility = Visibility.Collapsed;
                EstimatePreviewBrowser.Visibility = Visibility.Collapsed;
                EstimatePreviewPlaceholder.Visibility = Visibility.Collapsed;
                ScheduleEstimateEmbeddedLayout();
                return;
            }

            HideActiveExternalSpreadsheetInstance(isSecondary: false);

            Process process;
            try
            {
                process = Process.Start(new ProcessStartInfo
                {
                    FileName = preferredSpreadsheetEditorPath,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal
                }) ?? throw new InvalidOperationException("Не удалось запустить PlanMaker.");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Не удалось запустить PlanMaker: {ex.Message}");
            }

            try
            {
                process.WaitForInputIdle(5000);
            }
            catch
            {
                // Some builds of external editors may not support this call.
            }

            var handle = WaitForMainWindowHandle(process, timeoutMs: 25000);
            if (handle == IntPtr.Zero)
            {
                try
                {
                    if (!process.HasExited)
                        process.Kill();
                }
                catch
                {
                    // Ignore process shutdown failures.
                }

                throw new InvalidOperationException("PlanMaker не предоставил окно для предпросмотра.");
            }

            estimateSpreadsheetProcess = process;
            estimateEmbeddedFilePath = normalizedPath;
            estimateExcelWindowHandle = handle;
            var instance = new ExternalSpreadsheetInstance
            {
                FilePath = estimateEmbeddedFilePath,
                Process = process,
                Handle = handle
            };
            externalSpreadsheetInstances[estimateEmbeddedFilePath] = instance;
            activeExternalSpreadsheetInstance = instance;

            try
            {
                ShowWindow(estimateExcelWindowHandle, SW_HIDE);
            }
            catch
            {
                // Ignore early visibility errors before the window is positioned.
            }

            ConfigureFloatingEstimateWindow(estimateExcelWindowHandle);
            ResetEstimateEmbeddedLayoutCache();
            EstimateExcelHost.Visibility = Visibility.Collapsed;
            EstimatePreviewBrowser.Visibility = Visibility.Collapsed;
            EstimatePreviewPlaceholder.Visibility = Visibility.Collapsed;
            ScheduleEstimateEmbeddedLayout();
        }

        private void ShowEmbeddedEstimateInExternalEditorSecondary(string filePath)
        {
            if (string.IsNullOrWhiteSpace(preferredSpreadsheetEditorPath) || !File.Exists(preferredSpreadsheetEditorPath))
                throw new InvalidOperationException("PlanMaker не найден. Установите SoftMaker или укажите путь к PlanMaker.exe в настройках.");

            var normalizedPath = NormalizeDocumentPathKey(filePath);
            if (TryGetExternalSpreadsheetInstance(externalSpreadsheetInstancesSecondary, normalizedPath, out var cached))
            {
                HideActiveExternalSpreadsheetInstance(isSecondary: true);
                ApplyExternalSpreadsheetInstance(cached, isSecondary: true);
                ScheduleEstimateEmbeddedLayoutSecondary();
                return;
            }

            HideActiveExternalSpreadsheetInstance(isSecondary: true);

            Process process;
            try
            {
                process = Process.Start(new ProcessStartInfo
                {
                    FileName = preferredSpreadsheetEditorPath,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal
                }) ?? throw new InvalidOperationException("Не удалось запустить PlanMaker.");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Не удалось запустить PlanMaker: {ex.Message}");
            }

            try
            {
                process.WaitForInputIdle(5000);
            }
            catch
            {
                // ignore
            }

            var handle = WaitForMainWindowHandle(process, timeoutMs: 25000);
            if (handle == IntPtr.Zero)
            {
                try
                {
                    if (!process.HasExited)
                        process.Kill();
                }
                catch
                {
                    // ignore
                }

                throw new InvalidOperationException("PlanMaker не предоставил окно для предпросмотра.");
            }

            estimateSpreadsheetProcessSecondary = process;
            estimateEmbeddedFilePathSecondary = normalizedPath;
            estimateExcelWindowHandleSecondary = handle;
            var instance = new ExternalSpreadsheetInstance
            {
                FilePath = estimateEmbeddedFilePathSecondary,
                Process = process,
                Handle = handle
            };
            externalSpreadsheetInstancesSecondary[estimateEmbeddedFilePathSecondary] = instance;
            activeExternalSpreadsheetInstanceSecondary = instance;

            try
            {
                ShowWindow(estimateExcelWindowHandleSecondary, SW_HIDE);
            }
            catch
            {
                // Ignore early visibility errors.
            }

            ConfigureFloatingEstimateWindow(estimateExcelWindowHandleSecondary);
            ResetEstimateEmbeddedLayoutCacheSecondary();
            ScheduleEstimateEmbeddedLayoutSecondary();
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
            if (useExternalSpreadsheetEditor)
            {
                CloseAllExternalSpreadsheetInstances(externalSpreadsheetInstances, ref activeExternalSpreadsheetInstance);
                estimateSpreadsheetProcess = null;
                estimateExcelWindowHandle = IntPtr.Zero;
                estimateEmbeddedFilePath = string.Empty;
                ResetEstimateEmbeddedLayoutCache();
                return;
            }

            if (estimateSpreadsheetProcess != null)
            {
                try
                {
                    if (!estimateSpreadsheetProcess.HasExited)
                    {
                        estimateSpreadsheetProcess.CloseMainWindow();
                        if (!estimateSpreadsheetProcess.WaitForExit(1500))
                            estimateSpreadsheetProcess.Kill();
                    }
                }
                catch
                {
                    // Ignore external spreadsheet shutdown errors.
                }
                finally
                {
                    try
                    {
                        estimateSpreadsheetProcess.Dispose();
                    }
                    catch
                    {
                        // Ignore process dispose errors.
                    }

                    estimateSpreadsheetProcess = null;
                    estimateExcelWindowHandle = IntPtr.Zero;
                    ResetEstimateEmbeddedLayoutCache();
                    estimateEmbeddedFilePath = string.Empty;
                }
            }

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
                ResetEstimateEmbeddedLayoutCache();
                estimateEmbeddedFilePath = string.Empty;
            }
        }

        private void CloseEstimateWorkbookSecondary(bool saveChanges)
        {
            if (useExternalSpreadsheetEditor)
            {
                CloseAllExternalSpreadsheetInstances(externalSpreadsheetInstancesSecondary, ref activeExternalSpreadsheetInstanceSecondary);
                estimateSpreadsheetProcessSecondary = null;
                estimateExcelWindowHandleSecondary = IntPtr.Zero;
                estimateEmbeddedFilePathSecondary = string.Empty;
                ResetEstimateEmbeddedLayoutCacheSecondary();
                return;
            }

            if (estimateSpreadsheetProcessSecondary != null)
            {
                try
                {
                    if (!estimateSpreadsheetProcessSecondary.HasExited)
                    {
                        estimateSpreadsheetProcessSecondary.CloseMainWindow();
                        if (!estimateSpreadsheetProcessSecondary.WaitForExit(1500))
                            estimateSpreadsheetProcessSecondary.Kill();
                    }
                }
                catch
                {
                    // Ignore shutdown errors.
                }
                finally
                {
                    try
                    {
                        estimateSpreadsheetProcessSecondary.Dispose();
                    }
                    catch
                    {
                        // Ignore dispose errors.
                    }

                    estimateSpreadsheetProcessSecondary = null;
                    estimateExcelWindowHandleSecondary = IntPtr.Zero;
                    ResetEstimateEmbeddedLayoutCacheSecondary();
                    estimateEmbeddedFilePathSecondary = string.Empty;
                }
            }
        }

        private void DisposeEstimateExcelApplication()
        {
            CloseEstimateWorkbook(saveChanges: true);
            CloseEstimateWorkbookSecondary(saveChanges: true);

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
            if (EstimateExcelHost == null)
                return;
        }

        private IntPtr EnsureEstimatePreviewHostHandle()
        {
            if (EstimateExcelHost == null || EstimatePreviewContainer == null || EstimatePreviewBrowser == null || EstimatePreviewPlaceholder == null)
                throw new InvalidOperationException("Область предпросмотра сметы не готова.");

            EstimatePreviewContainer.Visibility = Visibility.Visible;
            EstimateExcelHost.Visibility = Visibility.Visible;
            EstimatePreviewBrowser.Visibility = Visibility.Collapsed;
            EstimatePreviewPlaceholder.Visibility = Visibility.Collapsed;

            EstimatePreviewContainer.UpdateLayout();
            EstimateExcelHost.UpdateLayout();

            var handle = EstimateExcelHost.HostHandle;
            if (handle != IntPtr.Zero)
                return handle;

            Dispatcher.Invoke(() =>
            {
                EstimatePreviewContainer.UpdateLayout();
                EstimateExcelHost.UpdateLayout();
            }, DispatcherPriority.Loaded);

            handle = EstimateExcelHost.HostHandle;
            if (handle == IntPtr.Zero)
                throw new InvalidOperationException("Не удалось подготовить встроенную область для сметы.");

            return handle;
        }

        private void MainWindow_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                reminderOverlayWindow?.Hide();
                HideEstimateEmbeddedPreview();
                HideEstimateEmbeddedSecondaryPreview();
                return;
            }

            UpdateReminderOverlayPlacement();

            if (ReferenceEquals(MainTabs?.SelectedItem, EstimateTab))
            {
                ScheduleEstimateEmbeddedLayout();
                ScheduleEstimateEmbeddedLayoutSecondary();
            }
        }

        private void MainWindow_Activated(object sender, EventArgs e)
        {
            UpdateReminderOverlayPlacement();

            if (ReferenceEquals(MainTabs?.SelectedItem, EstimateTab))
            {
                ScheduleEstimateEmbeddedLayout();
                ScheduleEstimateEmbeddedLayoutSecondary();
            }
        }

        private void MainWindow_Deactivated(object sender, EventArgs e)
        {
            if (useExternalSpreadsheetEditor && ReferenceEquals(MainTabs?.SelectedItem, EstimateTab))
                HideEstimateEmbeddedPreview();
            if (useExternalSpreadsheetEditor && ReferenceEquals(MainTabs?.SelectedItem, EstimateTab))
                HideEstimateEmbeddedSecondaryPreview();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            reminderRefreshTimer?.Stop();
            reminderRefreshDebounceTimer?.Stop();
            timesheetRebuildDebounceTimer?.Stop();
            estimateLayoutGuardTimer?.Stop();
            arrivalFilterDebounceTimer?.Stop();
            otSearchDebounceTimer?.Stop();
            summaryRefreshDebounceTimer?.Stop();
            processingOverlayDelayTimer?.Stop();
            autoSaveTimer?.Stop();
            gridPreferenceSaveDebounceTimer?.Stop();
            HideReminderOverlayWindow();
            StopEstimateEmbeddedPreview();
            StopEstimateEmbeddedSecondaryPreview();
            DisposeEstimateExcelApplication();
            ClosePdfExternalProcess();
            RemoveSessionMarker();
            ReleaseProjectLock();
            if (reminderOverlayWindow != null)
            {
                reminderOverlayWindow.Close();
                reminderOverlayWindow = null;
            }
        }

        private async Task CheckForUpdatesAsync(bool showUpToDateMessage)
        {
            var settings = currentObject?.UiSettings;
            var feedUrl = settings?.UpdateFeedUrl?.Trim();
            if (string.IsNullOrWhiteSpace(feedUrl))
            {
                if (showUpToDateMessage)
                    MessageBox.Show("URL обновлений не задан. Укажите его в настройках.", "Проверка обновлений", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                using var response = await UpdateHttpClient.GetAsync(feedUrl);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException($"Сервер вернул {response.StatusCode}.");

                var json = await response.Content.ReadAsStringAsync();
                var node = JsonNode.Parse(json);
                var versionText = node?["version"]?.ToString() ?? string.Empty;
                var url = node?["url"]?.ToString() ?? string.Empty;
                var zipUrl = node?["zipUrl"]?.ToString() ?? string.Empty;
                var notes = node?["notes"]?.ToString() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(versionText))
                    throw new InvalidOperationException("В файле обновлений нет версии.");

                if (!Version.TryParse(versionText, out var latestVersion))
                    throw new InvalidOperationException("Неверный формат версии обновления.");

                var currentVersion = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0, 0);

                if (latestVersion > currentVersion)
                {
                    var message = $"Доступна новая версия {latestVersion}.\nТекущая версия: {currentVersion}.";
                    if (!string.IsNullOrWhiteSpace(notes))
                        message += $"\n\nИзменения:\n{notes}";

                    var hasZip = !string.IsNullOrWhiteSpace(zipUrl) && zipUrl.EndsWith(".zip", StringComparison.OrdinalIgnoreCase);
                    var prompt = hasZip
                        ? message + "\n\nСкачать и установить обновление?"
                        : message + "\n\nОткрыть страницу загрузки?";

                    var result = MessageBox.Show(prompt, "Обновление доступно", MessageBoxButton.YesNo, MessageBoxImage.Information);
                    if (result == MessageBoxResult.Yes)
                    {
                        if (hasZip)
                        {
                            await DownloadAndApplyUpdateAsync(zipUrl);
                        }
                        else if (!string.IsNullOrWhiteSpace(url))
                        {
                            try
                            {
                                Process.Start(new ProcessStartInfo
                                {
                                    FileName = url,
                                    UseShellExecute = true
                                });
                            }
                            catch
                            {
                                MessageBox.Show("Не удалось открыть ссылку обновления.", "Обновление", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }
                    }

                    return;
                }

                if (showUpToDateMessage)
                {
                    MessageBox.Show("Установлена последняя версия.", "Проверка обновлений", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                if (showUpToDateMessage)
                    MessageBox.Show($"Не удалось проверить обновления.\n{ex.Message}", "Проверка обновлений", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async void CheckUpdatesMenu_Click(object sender, RoutedEventArgs e)
            => await CheckForUpdatesAsync(showUpToDateMessage: true);

        private async Task DownloadAndApplyUpdateAsync(string zipUrl)
        {
            if (string.IsNullOrWhiteSpace(zipUrl))
                return;

            var updaterExe = IOPath.Combine(AppContext.BaseDirectory, "Updater.exe");
            if (!File.Exists(updaterExe))
            {
                MessageBox.Show("Updater.exe не найден. Пересоберите приложение.", "Обновление", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var tempRoot = IOPath.Combine(IOPath.GetTempPath(), "MasterPRO_Update");
            var zipPath = IOPath.Combine(tempRoot, "update.zip");
            var extractPath = IOPath.Combine(tempRoot, "unpacked");

            try
            {
                Directory.CreateDirectory(tempRoot);
                if (Directory.Exists(extractPath))
                    Directory.Delete(extractPath, true);

                using (var response = await UpdateHttpClient.GetAsync(zipUrl))
                {
                    response.EnsureSuccessStatusCode();
                    await using var fs = new System.IO.FileStream(zipPath, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
                    await response.Content.CopyToAsync(fs);
                }

                System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, extractPath, overwriteFiles: true);

                var exeName = Environment.ProcessPath != null
                    ? IOPath.GetFileName(Environment.ProcessPath)
                    : "ConstructionControl.exe";
                var startInfo = new ProcessStartInfo
                {
                    FileName = updaterExe,
                    Arguments = $"--source \"{extractPath}\" --target \"{AppContext.BaseDirectory}\" --exe \"{exeName}\" --pid {Environment.ProcessId}",
                    UseShellExecute = true,
                    WorkingDirectory = AppContext.BaseDirectory
                };

                Process.Start(startInfo);
                closeConfirmed = true;
                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось установить обновление.\n{ex.Message}", "Обновление", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ReminderRefreshTimer_Tick(object sender, EventArgs e)
        {
            RequestReminderRefresh();
        }

        private void AutoSaveTimer_Tick(object sender, EventArgs e)
        {
            TryAutoSaveState();
        }

        private void AttachGridPreferenceTracking()
        {
            AttachGridPreferenceTracking(ArrivalLegacyGrid);
            AttachGridPreferenceTracking(OtJournalGrid);
            AttachGridPreferenceTracking(TimesheetGrid);
            AttachGridPreferenceTracking(ProductionJournalGrid);
            AttachGridPreferenceTracking(InspectionJournalGrid);
        }

        private void AttachGridPreferenceTracking(DataGrid grid)
        {
            if (grid == null)
                return;

            grid.ColumnReordered -= TrackGridPreferenceChanged;
            grid.ColumnReordered += TrackGridPreferenceChanged;
            grid.ColumnDisplayIndexChanged -= TrackGridPreferenceChanged;
            grid.ColumnDisplayIndexChanged += TrackGridPreferenceChanged;
        }

        private void TrackGridPreferenceChanged(object sender, DataGridColumnEventArgs e)
        {
            if (isApplyingColumnPreferences || sender is not DataGrid grid || currentObject == null)
                return;

            SaveGridColumnPreferences(grid);
            pendingGridPreferenceSave = true;
            gridPreferenceSaveDebounceTimer.Stop();
            gridPreferenceSaveDebounceTimer.Start();
        }

        private void GridPreferenceSaveDebounceTimer_Tick(object sender, EventArgs e)
        {
            gridPreferenceSaveDebounceTimer.Stop();
            if (!pendingGridPreferenceSave || currentObject == null)
                return;

            pendingGridPreferenceSave = false;
            SaveState(SaveTrigger.System);
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

        private bool IsSafeStartupEnabled()
            => currentObject?.UiSettings?.SafeStartupMode == true;

        private void EnsureSelectedTabInitialized()
        {
            if (MainTabs?.SelectedItem is TabItem selectedTab)
                EnsureTabInitialized(selectedTab);
        }

        private void EnsureTabInitialized(TabItem tabItem)
        {
            if (tabItem == null)
                return;

            if (ReferenceEquals(tabItem, TimesheetTab))
            {
                if (!timesheetInitialized)
                    InitializeTimesheet();
                return;
            }

            if (ReferenceEquals(tabItem, ProductionTab))
            {
                if (!productionJournalInitialized)
                    InitializeProductionJournal();
                else if (productionStateDirty)
                    RefreshProductionJournalState();
                else
                    RefreshProductionJournalLookups();
                return;
            }

            if (ReferenceEquals(tabItem, InspectionTab))
            {
                if (!inspectionJournalInitialized)
                    InitializeInspectionJournal();
                else if (inspectionStateDirty)
                    RefreshInspectionJournalState();
                else
                    RefreshInspectionLookups();
            }
        }

        private void ArrivalFilterDebounceTimer_Tick(object sender, EventArgs e)
        {
            arrivalFilterDebounceTimer.Stop();
            if (!arrivalFilterRefreshRequested)
                return;

            arrivalFilterRefreshRequested = false;
            ApplyAllFilters();
        }

        private void OtSearchDebounceTimer_Tick(object sender, EventArgs e)
        {
            otSearchDebounceTimer.Stop();
            if (!otSearchRefreshRequested)
                return;

            otSearchRefreshRequested = false;
            TryRefreshOtJournalView(scheduleRetryIfBusy: true);
        }

        private void SummaryRefreshDebounceTimer_Tick(object sender, EventArgs e)
        {
            summaryRefreshDebounceTimer.Stop();
            if (!summaryRefreshRequested)
                return;

            summaryRefreshRequested = false;
            StartSummaryRefreshRequest();
        }

        private void ProcessingOverlayDelayTimer_Tick(object sender, EventArgs e)
        {
            processingOverlayDelayTimer.Stop();
            if (processingOverlayDepth <= 0)
                return;

            if (ProcessingOverlayText != null)
                ProcessingOverlayText.Text = processingOverlayPendingText;

            if (ProcessingOverlay != null)
                ProcessingOverlay.Visibility = Visibility.Visible;
        }

        private void RequestArrivalFilterRefresh(bool immediate = false)
        {
            if (arrivalFilterDebounceTimer == null)
            {
                ApplyAllFilters();
                return;
            }

            if (immediate)
            {
                arrivalFilterRefreshRequested = false;
                arrivalFilterDebounceTimer.Stop();
                ApplyAllFilters();
                return;
            }

            arrivalFilterRefreshRequested = true;
            arrivalFilterDebounceTimer.Stop();
            arrivalFilterDebounceTimer.Start();
        }

        private void RequestOtSearchRefresh(bool immediate = false)
        {
            if (otSearchDebounceTimer == null)
            {
                if (!TryRefreshOtJournalView(scheduleRetryIfBusy: false))
                    Dispatcher.BeginInvoke(new Action(() => RequestOtSearchRefresh()), DispatcherPriority.Background);
                return;
            }

            if (immediate)
            {
                otSearchRefreshRequested = false;
                otSearchDebounceTimer.Stop();
                TryRefreshOtJournalView(scheduleRetryIfBusy: true);
                return;
            }

            otSearchRefreshRequested = true;
            otSearchDebounceTimer.Stop();
            otSearchDebounceTimer.Start();
        }

        private bool TryRefreshOtJournalView(bool scheduleRetryIfBusy)
        {
            var view = CollectionViewSource.GetDefaultView(OtJournalGrid?.ItemsSource);
            if (view == null)
                return true;

            if (view is IEditableCollectionView editableView && (editableView.IsAddingNew || editableView.IsEditingItem))
            {
                if (scheduleRetryIfBusy && otSearchDebounceTimer != null)
                {
                    otSearchRefreshRequested = true;
                    otSearchDebounceTimer.Stop();
                    otSearchDebounceTimer.Start();
                }

                return false;
            }

            view.Refresh();
            return true;
        }

        private void RequestSummaryRefresh(bool immediate = false)
        {
            if (summaryRefreshDebounceTimer == null)
            {
                StartSummaryRefreshRequest();
                return;
            }

            if (immediate)
            {
                summaryRefreshRequested = false;
                summaryRefreshDebounceTimer.Stop();
                StartSummaryRefreshRequest();
                return;
            }

            summaryRefreshRequested = true;
            summaryRefreshDebounceTimer.Stop();
            summaryRefreshDebounceTimer.Start();
        }

        private void StartSummaryRefreshRequest()
        {
            var requestId = Interlocked.Increment(ref summaryRefreshRequestVersion);
            _ = RefreshSummaryTableAsync(requestId);
        }

        private void MarkSummaryDataDirty()
        {
            summaryDataVersion++;
            summaryMatrixCache.Clear();
            productionLookupsDirty = true;
            inspectionLookupsDirty = true;
            productionStateDirty = true;
            inspectionStateDirty = true;
        }

        private string BuildSummaryMatrixCacheKey(IEnumerable<string> visibleGroups)
        {
            var groupsKey = string.Join("|", (visibleGroups ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase));

            var subTypeKey = string.IsNullOrWhiteSpace(summarySelectedSubType)
                ? "all"
                : summarySelectedSubType.Trim().ToLowerInvariant();

            return $"{summaryDataVersion}|{(summaryMountedMode ? "mounted" : "plan")}|{subTypeKey}|{groupsKey}";
        }

        private IDisposable BeginProcessingScope(string text)
        {
            processingOverlayDepth++;
            processingOverlayPendingText = string.IsNullOrWhiteSpace(text) ? "Идет обработка..." : text;
            AddOperationLogEntry("Фоновая операция", "Старт", processingOverlayPendingText);
            UpdateStatusBar();

            if (processingOverlayDelayTimer == null)
            {
                if (ProcessingOverlayText != null)
                    ProcessingOverlayText.Text = processingOverlayPendingText;
                if (ProcessingOverlay != null)
                    ProcessingOverlay.Visibility = Visibility.Visible;
                return new ProcessingScope(this);
            }

            if (!processingOverlayDelayTimer.IsEnabled)
            {
                processingOverlayDelayTimer.Stop();
                processingOverlayDelayTimer.Start();
            }

            return new ProcessingScope(this);
        }

        private void FinishProcessingScope()
        {
            if (processingOverlayDepth > 0)
                processingOverlayDepth--;

            if (processingOverlayDepth > 0)
            {
                UpdateStatusBar();
                return;
            }

            processingOverlayDelayTimer?.Stop();
            if (ProcessingOverlay != null)
                ProcessingOverlay.Visibility = Visibility.Collapsed;
            AddOperationLogEntry("Фоновая операция", "Завершено", processingOverlayPendingText);
            UpdateStatusBar();
        }

        private sealed class ProcessingScope : IDisposable
        {
            private MainWindow owner;

            public ProcessingScope(MainWindow owner)
            {
                this.owner = owner;
            }

            public void Dispose()
            {
                var currentOwner = Interlocked.Exchange(ref owner, null);
                currentOwner?.FinishProcessingScope();
            }
        }

        private void UpdateTabOpenDiagnostics(string tabHeader, TimeSpan elapsed)
        {
            if (string.IsNullOrWhiteSpace(tabHeader))
                return;

            tabOpenDiagnostics[tabHeader.Trim()] = elapsed;
            UpdateTabOpenDiagnosticsText();
        }

        private void UpdateTabOpenDiagnosticsText()
        {
            if (TabOpenDiagnosticsText == null)
                return;

            if (tabOpenDiagnostics.Count == 0)
            {
                TabOpenDiagnosticsText.Text = "Открытие вкладок: нет данных";
                return;
            }

            var top = tabOpenDiagnostics
                .OrderByDescending(x => x.Value.TotalMilliseconds)
                .Take(3)
                .Select(x => $"{x.Key}: {x.Value.TotalMilliseconds:0} мс");

            TabOpenDiagnosticsText.Text = $"Открытие вкладок: {string.Join(" | ", top)}";
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (initialUiPrepared)
                return;

            if (PdfTabStrip != null)
                PdfTabStrip.ItemsSource = openPdfTabs;
            if (EstimateTabStrip != null)
                EstimateTabStrip.ItemsSource = openEstimateTabs;

            initialUiPrepared = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                RequestArrivalFilterRefresh(immediate: true);
                Activate();
                RequestSummaryRefresh(immediate: true);
                if (!IsSafeStartupEnabled())
                    RequestReminderRefresh(immediate: true);
                UpdateReminderOverlayPlacement();
                EnsureSelectedTabInitialized();
                UpdateStatusBar();
            }), DispatcherPriority.Background);

            StartPreviewWarmupAsync();

            if (currentObject?.UiSettings?.CheckUpdatesOnStart == true
                && !string.IsNullOrWhiteSpace(currentObject.UiSettings.UpdateFeedUrl))
            {
                _ = CheckForUpdatesAsync(showUpToDateMessage: false);
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    var screen = SystemParameters.WorkArea;
                    if (Left < screen.Left - 50 || Top < screen.Top - 50 || Left > screen.Right - 50 || Top > screen.Bottom - 50)
                    {
                        WindowStartupLocation = WindowStartupLocation.CenterScreen;
                        Left = screen.Left + (screen.Width - Width) / 2;
                        Top = screen.Top + (screen.Height - Height) / 2;
                    }

                    if (!IsVisible)
                        Show();
                    if (WindowState == WindowState.Minimized)
                        WindowState = WindowState.Normal;
                    Activate();
                }
                catch { }
            }), DispatcherPriority.ApplicationIdle);
        }

        private void EnsureDocumentLibraries()
        {
            if (currentObject == null)
                return;

            currentObject.PdfDocuments ??= new List<DocumentTreeNode>();
            currentObject.EstimateDocuments ??= new List<DocumentTreeNode>();
            documentHashPathCache.Clear();
            NormalizeDocumentPaths(currentObject.PdfDocuments, isPdfLibrary: true);
            NormalizeDocumentPaths(currentObject.EstimateDocuments, isPdfLibrary: false);
            RebuildDocumentHashPathCache();
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
                    {
                        node.StoredRelativePath = relative;
                    }
                    else if (TryCopyDocumentToStorage(node.FilePath, isPdfLibrary, out var copiedPath, out var copiedRelativePath, out var copiedHash, out var copiedSize))
                    {
                        node.FilePath = copiedPath;
                        node.StoredRelativePath = copiedRelativePath;
                        node.ContentHash = copiedHash;
                        node.FileSizeBytes = copiedSize;
                        node.HashVerifiedAtUtc = DateTime.UtcNow;
                    }
                }

                var resolved = ResolveDocumentPath(node);
                if (!string.IsNullOrWhiteSpace(resolved))
                {
                    node.FilePath = resolved;
                    EnsureDocumentNodeFileMetadata(node, resolved);
                }

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

                return System.IO.Path.GetRelativePath(fullFolderPath, fullFilePath);
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
            {
                EnsureDocumentNodeFileMetadata(node, node.FilePath);
                return node.FilePath;
            }

            if (!string.IsNullOrWhiteSpace(node.StoredRelativePath))
            {
                var baseFolder = GetProjectStorageRoot(createIfMissing: false);
                if (!string.IsNullOrWhiteSpace(baseFolder))
                {
                    var candidate = System.IO.Path.Combine(baseFolder, node.StoredRelativePath);
                    if (File.Exists(candidate))
                    {
                        EnsureDocumentNodeFileMetadata(node, candidate);
                        return candidate;
                    }
                }
            }

            if (TryRecoverDocumentPath(node, out var recoveredPath))
            {
                node.FilePath = recoveredPath;
                EnsureDocumentNodeFileMetadata(node, recoveredPath);
                return recoveredPath;
            }

            return node.FilePath ?? string.Empty;
        }

        private bool TryRecoverDocumentPath(DocumentTreeNode node, out string recoveredPath)
        {
            recoveredPath = string.Empty;
            if (node == null || node.IsFolder)
                return false;

            if (!string.IsNullOrWhiteSpace(node.ContentHash) && TryGetStoredPathByHash(node.ContentHash, out var byHashPath))
            {
                recoveredPath = byHashPath;
                return true;
            }

            var fileName = string.Empty;
            try
            {
                if (!string.IsNullOrWhiteSpace(node.FilePath))
                    fileName = System.IO.Path.GetFileName(node.FilePath);

                if (string.IsNullOrWhiteSpace(fileName) && !string.IsNullOrWhiteSpace(node.StoredRelativePath))
                    fileName = System.IO.Path.GetFileName(node.StoredRelativePath);
            }
            catch
            {
                fileName = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(fileName))
                return false;

            foreach (var root in GetDocumentSearchRoots(node.FilePath))
            {
                if (TryFindFileByName(root, fileName, out var foundPath))
                {
                    recoveredPath = foundPath;
                    return true;
                }
            }

            return false;
        }

        private IEnumerable<string> GetDocumentSearchRoots(string originalPath)
        {
            var roots = new List<string>();
            var storageRoot = GetProjectStorageRoot(createIfMissing: false);
            if (!string.IsNullOrWhiteSpace(storageRoot) && Directory.Exists(storageRoot))
                roots.Add(storageRoot);

            if (!string.IsNullOrWhiteSpace(currentSaveFileName))
            {
                var saveFolder = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(currentSaveFileName));
                if (!string.IsNullOrWhiteSpace(saveFolder) && Directory.Exists(saveFolder))
                    roots.Add(saveFolder);
            }

            if (!string.IsNullOrWhiteSpace(originalPath))
            {
                var originalFolder = System.IO.Path.GetDirectoryName(originalPath);
                if (!string.IsNullOrWhiteSpace(originalFolder) && Directory.Exists(originalFolder))
                    roots.Add(originalFolder);
            }

            return roots
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => System.IO.Path.GetFullPath(x))
                .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static bool TryFindFileByName(string root, string fileName, out string foundPath)
        {
            foundPath = string.Empty;
            if (string.IsNullOrWhiteSpace(root) || string.IsNullOrWhiteSpace(fileName) || !Directory.Exists(root))
                return false;

            try
            {
                foreach (var file in Directory.EnumerateFiles(root, fileName, SearchOption.AllDirectories))
                {
                    foundPath = file;
                    return true;
                }
            }
            catch
            {
                // ignore search errors in inaccessible folders
            }

            return false;
        }

        private string GetProjectStorageRoot(bool createIfMissing)
        {
            if (string.IsNullOrWhiteSpace(currentSaveFileName))
                return string.Empty;

            var primaryRoot = BuildStorageRootPathForCurrentSettings(currentSaveFileName, createIfMissing);
            if (string.IsNullOrWhiteSpace(primaryRoot))
                return string.Empty;

            if (!createIfMissing && !Directory.Exists(primaryRoot))
            {
                var legacyRoot = BuildLegacyStorageRootPath(currentSaveFileName);
                if (!string.IsNullOrWhiteSpace(legacyRoot) && Directory.Exists(legacyRoot))
                    return legacyRoot;
            }

            return primaryRoot;
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
            => TryCopyDocumentToStorage(sourceFilePath, isPdfLibrary, out storedAbsolutePath, out storedRelativePath, out _, out _);

        private bool TryCopyDocumentToStorage(string sourceFilePath, bool isPdfLibrary, out string storedAbsolutePath, out string storedRelativePath, out string contentHash, out long fileSize)
        {
            storedAbsolutePath = sourceFilePath;
            storedRelativePath = string.Empty;
            contentHash = string.Empty;
            fileSize = 0;

            if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
                return false;

            if (!TryComputeFileHash(sourceFilePath, out contentHash, out fileSize))
                return false;

            if (!string.IsNullOrWhiteSpace(contentHash) && TryGetStoredPathByHash(contentHash, out var existingPath))
            {
                storedAbsolutePath = existingPath;
                var baseFolder = GetProjectStorageRoot(createIfMissing: false);
                if (!string.IsNullOrWhiteSpace(baseFolder))
                    storedRelativePath = System.IO.Path.GetRelativePath(baseFolder, existingPath);
                return true;
            }

            var folder = GetDocumentStorageFolder(isPdfLibrary, createIfMissing: true);
            if (string.IsNullOrWhiteSpace(folder))
                return false;

            try
            {
                var extension = System.IO.Path.GetExtension(sourceFilePath);
                var hashPrefix = string.IsNullOrWhiteSpace(contentHash)
                    ? Guid.NewGuid().ToString("N")
                    : contentHash[..Math.Min(contentHash.Length, 24)];
                var uniqueName = $"{hashPrefix}{extension}";
                var targetPath = System.IO.Path.Combine(folder, uniqueName);

                if (!File.Exists(targetPath))
                {
                    File.Copy(sourceFilePath, targetPath, overwrite: false);
                }
                else
                {
                    if (!TryComputeFileHash(targetPath, out var existingHash, out _)
                        || !string.Equals(existingHash, contentHash, StringComparison.OrdinalIgnoreCase))
                    {
                        uniqueName = $"{hashPrefix}_{Guid.NewGuid().ToString("N")[..6]}{extension}";
                        targetPath = System.IO.Path.Combine(folder, uniqueName);
                        File.Copy(sourceFilePath, targetPath, overwrite: false);
                    }
                }

                storedAbsolutePath = targetPath;
                var baseFolder = GetProjectStorageRoot(createIfMissing: true);
                storedRelativePath = System.IO.Path.GetRelativePath(baseFolder, targetPath);
                if (!string.IsNullOrWhiteSpace(contentHash))
                    documentHashPathCache[contentHash] = storedAbsolutePath;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryComputeFileHash(string filePath, out string hashHex, out long fileSize)
        {
            hashHex = string.Empty;
            fileSize = 0;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                fileSize = stream.Length;
                using var sha = SHA256.Create();
                var hash = sha.ComputeHash(stream);
                hashHex = Convert.ToHexString(hash);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool TryGetStoredPathByHash(string hash, out string storedPath)
        {
            storedPath = string.Empty;
            if (string.IsNullOrWhiteSpace(hash))
                return false;

            if (documentHashPathCache.TryGetValue(hash, out var cached) && File.Exists(cached))
            {
                storedPath = cached;
                return true;
            }

            var storageRoot = GetProjectStorageRoot(createIfMissing: false);
            if (string.IsNullOrWhiteSpace(storageRoot) || !Directory.Exists(storageRoot))
                return false;

            try
            {
                foreach (var file in Directory.EnumerateFiles(storageRoot, "*", SearchOption.AllDirectories))
                {
                    if (!TryComputeFileHash(file, out var fileHash, out _))
                        continue;

                    if (!string.Equals(fileHash, hash, StringComparison.OrdinalIgnoreCase))
                        continue;

                    documentHashPathCache[hash] = file;
                    storedPath = file;
                    return true;
                }
            }
            catch
            {
                // ignore inaccessible files
            }

            return false;
        }

        private void RebuildDocumentHashPathCache()
        {
            documentHashPathCache.Clear();
            void Collect(IEnumerable<DocumentTreeNode> nodes)
            {
                if (nodes == null)
                    return;

                foreach (var node in nodes)
                {
                    if (node == null)
                        continue;

                    if (node.IsFolder)
                    {
                        Collect(node.Children);
                        continue;
                    }

                    var resolved = ResolveDocumentPath(node);
                    if (string.IsNullOrWhiteSpace(resolved) || !File.Exists(resolved))
                        continue;

                    EnsureDocumentNodeFileMetadata(node, resolved);
                    if (!string.IsNullOrWhiteSpace(node.ContentHash))
                        documentHashPathCache[node.ContentHash] = resolved;
                }
            }

            Collect(currentObject?.PdfDocuments);
            Collect(currentObject?.EstimateDocuments);
        }

        private void EnsureDocumentNodeFileMetadata(DocumentTreeNode node, string filePath)
        {
            if (node == null || node.IsFolder || string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return;

            var info = new FileInfo(filePath);
            var shouldRehash = string.IsNullOrWhiteSpace(node.ContentHash)
                || !node.FileSizeBytes.HasValue
                || node.FileSizeBytes.Value != info.Length;

            if (shouldRehash && TryComputeFileHash(filePath, out var hash, out var fileSize))
            {
                node.ContentHash = hash;
                node.FileSizeBytes = fileSize;
                node.HashVerifiedAtUtc = DateTime.UtcNow;
                if (!string.IsNullOrWhiteSpace(hash))
                    documentHashPathCache[hash] = filePath;
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
            SetTabButtonState(SummaryTabButton, ReferenceEquals(MainTabs?.SelectedItem, SummaryTab), "Сводка");
            SetTabButtonState(JvkTabButton, ReferenceEquals(MainTabs?.SelectedItem, JvkTab), "ЖВК");
            SetTabButtonState(ArrivalTabButton, ReferenceEquals(MainTabs?.SelectedItem, ArrivalTab), "Приход");
            SetTabButtonState(OtTabButton, ReferenceEquals(MainTabs?.SelectedItem, OtTab), "ОТ");
            SetTabButtonState(TimesheetTabButton, ReferenceEquals(MainTabs?.SelectedItem, TimesheetTab), "Табель");
            SetTabButtonState(ProductionTabButton, ReferenceEquals(MainTabs?.SelectedItem, ProductionTab), "ПР");
            SetTabButtonState(InspectionTabButton, ReferenceEquals(MainTabs?.SelectedItem, InspectionTab), "Осмотры");
            SetTabButtonState(PdfPinnedTabButton, ReferenceEquals(MainTabs?.SelectedItem, PdfTab), "ПДФ");
            SetTabButtonState(EstimatePinnedTabButton, ReferenceEquals(MainTabs?.SelectedItem, EstimateTab), "Сметы");
        }

        private void SetTabButtonState(Button button, bool isActive, string tabHeader)
        {
            if (button == null)
                return;

            var hasReminder = tabReminderMessages.TryGetValue(tabHeader, out var reminderItems) && reminderItems?.Count > 0;
            var reminderToolTip = hasReminder
                ? $"Требуется действие:\n• {string.Join("\n• ", reminderItems.Take(4))}"
                : null;
            var highlightReminder = ShouldHighlightReminderTabs() && hasReminder;

            if (isActive)
            {
                button.Background = (Brush)new BrushConverter().ConvertFromString(highlightReminder ? "#FDE68A" : "#DBEAFE");
                button.BorderBrush = (Brush)new BrushConverter().ConvertFromString(highlightReminder ? "#F59E0B" : "#3B82F6");
                button.Foreground = (Brush)new BrushConverter().ConvertFromString("#111827");
            }
            else if (highlightReminder)
            {
                button.Background = (Brush)new BrushConverter().ConvertFromString("#FEF3C7");
                button.BorderBrush = (Brush)new BrushConverter().ConvertFromString("#F59E0B");
                button.Foreground = (Brush)new BrushConverter().ConvertFromString("#92400E");
            }
            else
            {
                button.Background = (Brush)new BrushConverter().ConvertFromString("#F9FAFB");
                button.BorderBrush = (Brush)new BrushConverter().ConvertFromString("#E5E7EB");
                button.Foreground = (Brush)new BrushConverter().ConvertFromString(isActive ? "#111827" : "#6B7280");
            }

            button.ToolTip = reminderToolTip;
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

        private void PdfActionsMenuButton_Click(object sender, RoutedEventArgs e)
            => OpenInlineActionsMenu(sender as FrameworkElement);

        private void EstimateActionsMenuButton_Click(object sender, RoutedEventArgs e)
            => OpenInlineActionsMenu(sender as FrameworkElement);

        private static void OpenInlineActionsMenu(FrameworkElement sourceElement)
        {
            if (sourceElement?.ContextMenu == null)
                return;

            sourceElement.ContextMenu.PlacementTarget = sourceElement;
            sourceElement.ContextMenu.Placement = PlacementMode.Bottom;
            sourceElement.ContextMenu.IsOpen = true;
        }

        private void UpdatePdfSelectionInfo()
        {
            if (selectedPdfNode != null && !selectedPdfNode.IsFolder)
                activePdfTab = selectedPdfNode;
            var active = activePdfTab ?? selectedPdfNode;
            UpdateDocumentSelectionInfo(active, PdfSelectedNameText, PdfSelectedPathText, PdfSelectedTypeText);
            UpdatePdfPreview(active);
            UpdatePdfSecondaryPreview();
        }

        private void UpdatePdfPreview(DocumentTreeNode node)
        {
            if (PdfPreviewContainer == null || PdfPreviewPlaceholder == null || PdfPreviewStatusText == null)
                return;

            if (!useExternalPdfEditor)
            {
                var activePath = NormalizeDocumentPathKey(ResolveDocumentPath(node));
                if (node == null
                    || !string.Equals(pdfPreviewCurrentPath, activePath, StringComparison.CurrentCultureIgnoreCase))
                {
                    UpdateDocumentPreview(node, PdfInfoPanel, PdfPreviewContainer, PdfPreviewBrowser, PdfPreviewPlaceholder, PdfPreviewStatusText);
                    pdfPreviewCurrentPath = activePath;
                }

                return;
            }

            if (!ReferenceEquals(MainTabs?.SelectedItem, PdfTab))
            {
                HidePdfEmbeddedPreview();
                return;
            }

            if (node == null)
            {
                HidePdfEmbeddedPreview();
                PdfPreviewContainer.Visibility = Visibility.Visible;
                pdfPreviewCurrentPath = string.Empty;
                ShowDocumentPreviewPlaceholder(PdfPreviewBrowser, PdfPreviewPlaceholder, PdfPreviewStatusText, "Выберите PDF-файл в дереве слева, и он откроется здесь.");
                return;
            }

            if (node.IsFolder)
            {
                HidePdfEmbeddedPreview();
                PdfPreviewContainer.Visibility = Visibility.Visible;
                pdfPreviewCurrentPath = string.Empty;
                ShowDocumentPreviewPlaceholder(PdfPreviewBrowser, PdfPreviewPlaceholder, PdfPreviewStatusText, "Для папки предпросмотр не показывается. Выберите конкретный файл.");
                return;
            }

            var resolvedPath = ResolveDocumentPath(node);
            if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath))
            {
                HidePdfEmbeddedPreview();
                PdfPreviewContainer.Visibility = Visibility.Visible;
                pdfPreviewCurrentPath = string.Empty;
                ShowDocumentPreviewPlaceholder(PdfPreviewBrowser, PdfPreviewPlaceholder, PdfPreviewStatusText, "Файл не найден по сохраненному пути.");
                return;
            }

            PdfPreviewContainer.Visibility = Visibility.Visible;
            if (PdfExternalHost != null)
                PdfExternalHost.Visibility = Visibility.Collapsed;
            PdfPreviewBrowser.Visibility = Visibility.Collapsed;
            PdfPreviewPlaceholder.Visibility = Visibility.Collapsed;

            var normalizedPath = NormalizeDocumentPathKey(resolvedPath);
            if (string.Equals(pdfPreviewCurrentPath, normalizedPath, StringComparison.CurrentCultureIgnoreCase))
            {
                SchedulePdfEmbeddedLayout();
                return;
            }

            pdfPreviewCurrentPath = normalizedPath;
            OpenPdfInExternalEditor(resolvedPath);
        }

        private void UpdateEstimateSelectionInfo()
        {
            if (activeEstimateTab == null && openEstimateTabs.Count > 0)
                activeEstimateTab = openEstimateTabs[0];
            var active = activeEstimateTab ?? selectedEstimateNode;
            UpdateDocumentSelectionInfo(active, EstimateSelectedNameText, EstimateSelectedPathText, EstimateSelectedTypeText);
            var activePath = NormalizeDocumentPathKey(ResolveDocumentPath(active));
            if (active == null
                || !string.Equals(estimatePreviewCurrentPath, activePath, StringComparison.CurrentCultureIgnoreCase))
            {
                UpdateEstimatePreview(active);
                estimatePreviewCurrentPath = activePath;
            }
            UpdateEstimateSecondaryPreview();
        }

        private void OpenPdfTab(DocumentTreeNode node)
        {
            if (node == null || node.IsFolder)
                return;

            activePdfTab = node;
            if (PdfTabStrip != null)
                PdfTabStrip.SelectedItem = node;
            UpdatePdfSelectionInfo();
        }

        private void OpenEstimateTab(DocumentTreeNode node)
        {
            if (node == null || node.IsFolder)
                return;

            var existing = FindOpenTab(openEstimateTabs, node);
            if (existing == null)
            {
                openEstimateTabs.Add(node);
                existing = node;
            }

            activeEstimateTab = existing;
            if (EstimateTabStrip != null)
                EstimateTabStrip.SelectedItem = existing;
            UpdateEstimateSelectionInfo();
        }

        private static DocumentTreeNode FindOpenTab(ObservableCollection<DocumentTreeNode> list, DocumentTreeNode node)
        {
            if (list == null || node == null)
                return null;

            var target = node.FilePath ?? string.Empty;
            foreach (var item in list)
            {
                if (ReferenceEquals(item, node))
                    return item;
                if (!string.IsNullOrWhiteSpace(target) && string.Equals(item?.FilePath, target, StringComparison.CurrentCultureIgnoreCase))
                    return item;
            }

            return null;
        }

        private static string NormalizeDocumentPathKey(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;

            try
            {
                return System.IO.Path.GetFullPath(path).Trim();
            }
            catch
            {
                return path.Trim();
            }
        }

        private void PdfTabStrip_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PdfTabStrip?.SelectedItem is DocumentTreeNode node)
            {
                activePdfTab = node;
                UpdatePdfSelectionInfo();
            }
        }

        private void EstimateTabStrip_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (EstimateTabStrip?.SelectedItem is DocumentTreeNode node)
            {
                activeEstimateTab = node;
                UpdateEstimateSelectionInfo();
            }
        }

        private void PdfTabClose_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is DocumentTreeNode node)
            {
                openPdfTabs.Remove(node);
                if (ReferenceEquals(activePdfTab, node))
                    activePdfTab = openPdfTabs.FirstOrDefault();
                if (ReferenceEquals(secondaryPdfTab, node))
                    secondaryPdfTab = null;
                if (PdfTabStrip != null)
                    PdfTabStrip.SelectedItem = activePdfTab;
                UpdatePdfSelectionInfo();
            }
        }

        private void EstimateTabClose_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is DocumentTreeNode node)
            {
                openEstimateTabs.Remove(node);
                if (ReferenceEquals(activeEstimateTab, node))
                    activeEstimateTab = openEstimateTabs.FirstOrDefault();
                if (ReferenceEquals(secondaryEstimateTab, node))
                    secondaryEstimateTab = null;
                if (EstimateTabStrip != null)
                    EstimateTabStrip.SelectedItem = activeEstimateTab;
                UpdateEstimateSelectionInfo();

                if (useExternalSpreadsheetEditor && !string.IsNullOrWhiteSpace(node.FilePath))
                {
                    CloseExternalSpreadsheetInstance(externalSpreadsheetInstances, NormalizeDocumentPathKey(node.FilePath), ref activeExternalSpreadsheetInstance);
                    CloseExternalSpreadsheetInstance(externalSpreadsheetInstancesSecondary, NormalizeDocumentPathKey(node.FilePath), ref activeExternalSpreadsheetInstanceSecondary);
                }
            }
        }

        private void PdfSplitToggle_Checked(object sender, RoutedEventArgs e)
        {
        }

        private void PdfSplitToggle_Unchecked(object sender, RoutedEventArgs e)
        {
        }

        private void EstimateSplitToggle_Checked(object sender, RoutedEventArgs e)
        {
        }

        private void EstimateSplitToggle_Unchecked(object sender, RoutedEventArgs e)
        {
        }

        private void UpdatePdfSecondaryPreview()
        {
            if (PdfSplitBar == null || PdfPreviewContainerSecondary == null || PdfPreviewBrowserSecondary == null || PdfPreviewPlaceholderSecondary == null || PdfPreviewStatusTextSecondary == null)
                return;

            pdfSplitEnabled = false;
            secondaryPdfTab = null;
            UpdatePdfSplitColumns();
            PdfSplitBar.Visibility = Visibility.Collapsed;
            PdfPreviewContainerSecondary.Visibility = Visibility.Collapsed;
            if (PdfSplitDropOverlay != null)
                PdfSplitDropOverlay.Visibility = Visibility.Collapsed;
        }

        private void UpdatePdfSplitColumns()
        {
            if (PdfSplitPrimaryColumn == null || PdfSplitDividerColumn == null || PdfSplitSecondaryColumn == null)
                return;

            if (pdfSplitEnabled)
            {
                PdfSplitPrimaryColumn.Width = new GridLength(1, GridUnitType.Star);
                PdfSplitDividerColumn.Width = new GridLength(6, GridUnitType.Pixel);
                PdfSplitSecondaryColumn.Width = new GridLength(1, GridUnitType.Star);
            }
            else
            {
                PdfSplitPrimaryColumn.Width = new GridLength(1, GridUnitType.Star);
                PdfSplitDividerColumn.Width = new GridLength(0);
                PdfSplitSecondaryColumn.Width = new GridLength(0);
            }
        }

        private void UpdateEstimateSecondaryPreview()
        {
            if (EstimateSplitBar == null || EstimatePreviewContainerSecondary == null || EstimatePreviewStatusTextSecondary == null)
                return;

            if (secondaryEstimateTab == null)
                estimateSplitEnabled = false;

            UpdateEstimateSplitColumns();

            if (!estimateSplitEnabled)
            {
                EstimateSplitBar.Visibility = Visibility.Collapsed;
                EstimatePreviewContainerSecondary.Visibility = Visibility.Collapsed;
                HideEstimateEmbeddedSecondaryPreview();
                if (EstimateSplitDropOverlay != null)
                    EstimateSplitDropOverlay.Visibility = Visibility.Collapsed;
                return;
            }

            EstimateSplitBar.Visibility = Visibility.Visible;
            EstimatePreviewContainerSecondary.Visibility = Visibility.Visible;

            if (secondaryEstimateTab == null)
            {
                EstimatePreviewStatusTextSecondary.Text = "Перетащите вкладку сюда для открытия рядом.";
                HideEstimateEmbeddedSecondaryPreview();
                return;
            }

            UpdateEstimateSecondaryPreviewContent(secondaryEstimateTab);
        }

        private void UpdateEstimateSplitColumns()
        {
            if (EstimateSplitPrimaryColumn == null || EstimateSplitDividerColumn == null || EstimateSplitSecondaryColumn == null)
                return;

            if (estimateSplitEnabled)
            {
                EstimateSplitPrimaryColumn.Width = new GridLength(1, GridUnitType.Star);
                EstimateSplitDividerColumn.Width = new GridLength(6, GridUnitType.Pixel);
                EstimateSplitSecondaryColumn.Width = new GridLength(1, GridUnitType.Star);
            }
            else
            {
                EstimateSplitPrimaryColumn.Width = new GridLength(1, GridUnitType.Star);
                EstimateSplitDividerColumn.Width = new GridLength(0);
                EstimateSplitSecondaryColumn.Width = new GridLength(0);
            }
        }

        private void PdfTabStrip_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            pdfTabDragStart = e.GetPosition(null);
            pdfDraggedTab = GetTabNodeFromOriginalSource(e.OriginalSource);
        }

        private void PdfTabStrip_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || pdfDraggedTab == null)
                return;

            var position = e.GetPosition(null);
            if (Math.Abs(position.X - pdfTabDragStart.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(position.Y - pdfTabDragStart.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            var result = DragDrop.DoDragDrop(PdfTabStrip, pdfDraggedTab, DragDropEffects.Move);
            if (result == DragDropEffects.None)
                TogglePdfDetach(pdfDraggedTab);
            pdfDraggedTab = null;
        }

        private void PdfTabStrip_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(DocumentTreeNode)) is not DocumentTreeNode node)
                return;

            var target = GetTabNodeFromOriginalSource(e.OriginalSource);
            if (target == null)
            {
                TogglePdfDetach(node);
                return;
            }

            ReorderTab(openPdfTabs, node, target);
        }

        private void PdfSecondaryPane_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(DocumentTreeNode)) is not DocumentTreeNode node)
                return;

            OpenPdfTab(node);
        }

        private void EstimateTabStrip_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            estimateTabDragStart = e.GetPosition(null);
            estimateDraggedTab = GetTabNodeFromOriginalSource(e.OriginalSource);
        }

        private void EstimateTabStrip_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || estimateDraggedTab == null)
                return;

            var position = e.GetPosition(null);
            if (Math.Abs(position.X - estimateTabDragStart.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(position.Y - estimateTabDragStart.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            var result = DragDrop.DoDragDrop(EstimateTabStrip, estimateDraggedTab, DragDropEffects.Move);
            if (result == DragDropEffects.None)
                ToggleEstimateDetach(estimateDraggedTab);
            estimateDraggedTab = null;
        }

        private void EstimateTabStrip_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(DocumentTreeNode)) is not DocumentTreeNode node)
                return;

            var target = GetTabNodeFromOriginalSource(e.OriginalSource);
            if (target == null)
            {
                ToggleEstimateDetach(node);
                return;
            }

            ReorderTab(openEstimateTabs, node, target);
        }

        private void EstimateSecondaryPane_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(DocumentTreeNode)) is not DocumentTreeNode node)
                return;

            if (FindOpenTab(openEstimateTabs, node) == null)
                openEstimateTabs.Add(node);
            secondaryEstimateTab = node;
            estimateSplitEnabled = true;
            UpdateEstimateSelectionInfo();
        }

        private void PdfSplitGrid_DragEnter(object sender, DragEventArgs e)
        {
            if (!TryGetDragDocumentNode(e, out _))
                return;

            e.Effects = DragDropEffects.Move;
            UpdatePdfSplitDropOverlay(e);
            e.Handled = true;
        }

        private void PdfSplitGrid_DragOver(object sender, DragEventArgs e)
        {
            if (!TryGetDragDocumentNode(e, out _))
                return;

            e.Effects = DragDropEffects.Move;
            UpdatePdfSplitDropOverlay(e);
            e.Handled = true;
        }

        private void PdfSplitGrid_DragLeave(object sender, DragEventArgs e)
        {
            HidePdfSplitDropOverlay();
        }

        private void PdfSplitGrid_Drop(object sender, DragEventArgs e)
        {
            if (!TryGetDragDocumentNode(e, out var node))
                return;

            if (node.IsFolder)
                return;

            HidePdfSplitDropOverlay();
            OpenPdfTab(node);

            e.Handled = true;
        }

        private void EstimateSplitGrid_DragEnter(object sender, DragEventArgs e)
        {
            if (!TryGetDragDocumentNode(e, out _))
                return;

            e.Effects = DragDropEffects.Move;
            UpdateEstimateSplitDropOverlay(e);
            e.Handled = true;
        }

        private void EstimateSplitGrid_DragOver(object sender, DragEventArgs e)
        {
            if (!TryGetDragDocumentNode(e, out _))
                return;

            e.Effects = DragDropEffects.Move;
            UpdateEstimateSplitDropOverlay(e);
            e.Handled = true;
        }

        private void EstimateSplitGrid_DragLeave(object sender, DragEventArgs e)
        {
            HideEstimateSplitDropOverlay();
        }

        private void EstimateSplitGrid_Drop(object sender, DragEventArgs e)
        {
            if (!TryGetDragDocumentNode(e, out var node))
                return;

            if (node.IsFolder)
                return;

            var side = GetSplitDropSide(EstimateSplitGrid, e);
            HideEstimateSplitDropOverlay();

            if (side == SplitDropSide.Right)
            {
                if (FindOpenTab(openEstimateTabs, node) == null)
                    openEstimateTabs.Add(node);
                secondaryEstimateTab = node;
                estimateSplitEnabled = true;
                UpdateEstimateSelectionInfo();
            }
            else
            {
                OpenEstimateTab(node);
            }

            e.Handled = true;
        }

        private enum SplitDropSide
        {
            Left,
            Right
        }

        private static bool TryGetDragDocumentNode(DragEventArgs e, out DocumentTreeNode node)
        {
            node = e.Data.GetData(typeof(DocumentTreeNode)) as DocumentTreeNode;
            return node != null;
        }

        private static SplitDropSide GetSplitDropSide(FrameworkElement target, DragEventArgs e)
        {
            if (target == null)
                return SplitDropSide.Left;

            var position = e.GetPosition(target);
            return position.X <= target.ActualWidth / 2
                ? SplitDropSide.Left
                : SplitDropSide.Right;
        }

        private void UpdatePdfSplitDropOverlay(DragEventArgs e)
        {
            if (PdfSplitDropOverlay == null || PdfSplitDropLeft == null || PdfSplitDropRight == null)
                return;

            PdfSplitDropOverlay.Visibility = Visibility.Visible;
            var side = GetSplitDropSide(PdfSplitGrid, e);
            PdfSplitDropLeft.Opacity = side == SplitDropSide.Left ? 0.55 : 0.15;
            PdfSplitDropRight.Opacity = side == SplitDropSide.Right ? 0.55 : 0.15;
        }

        private void HidePdfSplitDropOverlay()
        {
            if (PdfSplitDropOverlay != null)
                PdfSplitDropOverlay.Visibility = Visibility.Collapsed;
        }

        private void UpdateEstimateSplitDropOverlay(DragEventArgs e)
        {
            if (EstimateSplitDropOverlay == null || EstimateSplitDropLeft == null || EstimateSplitDropRight == null)
                return;

            EstimateSplitDropOverlay.Visibility = Visibility.Visible;
            var side = GetSplitDropSide(EstimateSplitGrid, e);
            EstimateSplitDropLeft.Opacity = side == SplitDropSide.Left ? 0.55 : 0.15;
            EstimateSplitDropRight.Opacity = side == SplitDropSide.Right ? 0.55 : 0.15;
        }

        private void HideEstimateSplitDropOverlay()
        {
            if (EstimateSplitDropOverlay != null)
                EstimateSplitDropOverlay.Visibility = Visibility.Collapsed;
        }

        private static DocumentTreeNode GetTabNodeFromOriginalSource(object source)
        {
            if (source is FrameworkElement fe && fe.DataContext is DocumentTreeNode direct)
                return direct;

            var current = source as DependencyObject;
            while (current != null)
            {
                if (current is FrameworkElement element && element.DataContext is DocumentTreeNode node)
                    return node;

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        private static void ReorderTab(ObservableCollection<DocumentTreeNode> list, DocumentTreeNode dragged, DocumentTreeNode target)
        {
            if (list == null || dragged == null)
                return;

            var oldIndex = list.IndexOf(dragged);
            if (oldIndex < 0)
                return;

            var newIndex = target == null ? list.Count - 1 : list.IndexOf(target);
            if (newIndex < 0)
                newIndex = list.Count - 1;
            if (newIndex == oldIndex)
                return;

            list.Move(oldIndex, newIndex);
        }

        private void PdfDetachButton_Click(object sender, RoutedEventArgs e)
        {
            var node = activePdfTab ?? selectedPdfNode;
            if (node == null || node.IsFolder)
                return;

            if (pdfDetachedWindow != null)
            {
                try { pdfDetachedWindow.Close(); } catch { }
                pdfDetachedWindow = null;
                pdfDetachedBrowser = null;
                pdfDetachedNode = null;
                return;
            }

            var path = ResolveDocumentPath(node);
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show("Файл не найден.");
                return;
            }

            pdfDetachedNode = node;
            pdfDetachedBrowser = new WebBrowser();
            pdfDetachedWindow = new Window
            {
                Title = node.Name,
                Owner = this,
                Width = 1200,
                Height = 800,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Content = pdfDetachedBrowser
            };
            pdfDetachedWindow.Closed += (_, _) =>
            {
                pdfDetachedWindow = null;
                pdfDetachedBrowser = null;
                pdfDetachedNode = null;
            };
            pdfDetachedBrowser.Navigate(path);
            pdfDetachedWindow.Show();
        }

        private void TogglePdfDetach(DocumentTreeNode node)
        {
            if (node == null || node.IsFolder)
                return;

            if (pdfDetachedWindow != null && ReferenceEquals(pdfDetachedNode, node))
            {
                try { pdfDetachedWindow.Close(); } catch { }
                pdfDetachedWindow = null;
                pdfDetachedBrowser = null;
                pdfDetachedNode = null;
                return;
            }

            PdfDetachButton_Click(this, new RoutedEventArgs());
        }

        private void EstimateDetachButton_Click(object sender, RoutedEventArgs e)
        {
            if (estimateExcelWindowHandle == IntPtr.Zero)
                return;

            estimateDetached = !estimateDetached;
            if (estimateDetached)
            {
                ConfigureDetachedEstimateWindow(estimateExcelWindowHandle);
                ShowWindow(estimateExcelWindowHandle, SW_SHOW);
            }
            else
            {
                ConfigureFloatingEstimateWindow(estimateExcelWindowHandle);
                ScheduleEstimateEmbeddedLayout();
            }
        }

        private void ToggleEstimateDetach(DocumentTreeNode node)
        {
            if (node == null || node.IsFolder)
                return;

            if (estimateExcelWindowHandle == IntPtr.Zero || !string.Equals(estimateEmbeddedFilePath, node.FilePath ?? string.Empty, StringComparison.CurrentCultureIgnoreCase))
            {
                OpenEstimateTab(node);
            }

            estimateDetached = !estimateDetached;
            if (estimateDetached)
            {
                ConfigureDetachedEstimateWindow(estimateExcelWindowHandle);
                ShowWindow(estimateExcelWindowHandle, SW_SHOW);
            }
            else
            {
                ConfigureFloatingEstimateWindow(estimateExcelWindowHandle);
                ScheduleEstimateEmbeddedLayout();
            }
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

            if (useExternalSpreadsheetEditor && estimateDetached)
            {
                return;
            }

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

            if (useExternalSpreadsheetEditor
                && estimateExcelWindowHandle != IntPtr.Zero
                && IsEstimateSpreadsheetProcessAlive()
                && string.Equals(estimateEmbeddedFilePath, resolvedPath, StringComparison.CurrentCultureIgnoreCase))
            {
                EstimateInfoPanel.Visibility = Visibility.Collapsed;
                EstimatePreviewContainer.Visibility = Visibility.Visible;
                if (EstimateExcelHost != null)
                    EstimateExcelHost.Visibility = Visibility.Collapsed;

                EstimatePreviewBrowser.Visibility = Visibility.Collapsed;
                EstimatePreviewPlaceholder.Visibility = Visibility.Collapsed;
                ScheduleEstimateEmbeddedLayout();
                return;
            }

            StopEstimateEmbeddedPreview();
            EstimateInfoPanel.Visibility = Visibility.Visible;
            EstimatePreviewContainer.Visibility = Visibility.Collapsed;

            try
            {
                EstimateInfoPanel.Visibility = Visibility.Collapsed;
                EstimatePreviewContainer.Visibility = Visibility.Visible;
                ShowEmbeddedEstimateWorkbook(node, resolvedPath);
            }
            catch (Exception ex)
            {
                StopEstimateEmbeddedPreview();
                EstimateInfoPanel.Visibility = Visibility.Visible;
                EstimatePreviewContainer.Visibility = Visibility.Collapsed;
                ShowEstimatePreviewPlaceholder($"Не удалось открыть смету во встроенном режиме: {ex.Message}");
            }
        }

        private void UpdateEstimateSecondaryPreviewContent(DocumentTreeNode node)
        {
            if (EstimatePreviewContainerSecondary == null || EstimatePreviewStatusTextSecondary == null)
                return;

            if (node == null || node.IsFolder)
            {
                EstimatePreviewStatusTextSecondary.Text = "Выберите файл сметы.";
                HideEstimateEmbeddedSecondaryPreview();
                return;
            }

            var resolvedPath = ResolveDocumentPath(node);
            if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath))
            {
                EstimatePreviewStatusTextSecondary.Text = "Файл не найден по сохраненному пути.";
                HideEstimateEmbeddedSecondaryPreview();
                return;
            }

            var extension = System.IO.Path.GetExtension(resolvedPath)?.ToLowerInvariant() ?? string.Empty;
            if (!IsEstimateExcelExtension(extension))
            {
                EstimatePreviewStatusTextSecondary.Text = "Формат не поддерживается для разделенного просмотра.";
                HideEstimateEmbeddedSecondaryPreview();
                return;
            }

            if (!useExternalSpreadsheetEditor)
            {
                EstimatePreviewStatusTextSecondary.Text = "Разделенный просмотр доступен только во внешнем режиме.";
                HideEstimateEmbeddedSecondaryPreview();
                return;
            }

            try
            {
                ShowEmbeddedEstimateInExternalEditorSecondary(resolvedPath);
                EstimatePreviewStatusTextSecondary.Text = string.Empty;
                ScheduleEstimateEmbeddedLayoutSecondary();
            }
            catch (Exception ex)
            {
                HideEstimateEmbeddedSecondaryPreview();
                EstimatePreviewStatusTextSecondary.Text = $"Не удалось открыть смету: {ex.Message}";
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
            if (previewContainer == null || browser == null || placeholder == null || statusText == null)
                return;

            if (node == null)
            {
                if (infoPanel != null)
                    infoPanel.Visibility = Visibility.Visible;
                previewContainer.Visibility = Visibility.Collapsed;
                ShowDocumentPreviewPlaceholder(browser, placeholder, statusText, "Выберите файл в дереве слева, и он откроется здесь.");
                return;
            }

            if (node.IsFolder)
            {
                if (infoPanel != null)
                    infoPanel.Visibility = Visibility.Visible;
                previewContainer.Visibility = Visibility.Collapsed;
                ShowDocumentPreviewPlaceholder(browser, placeholder, statusText, "Для папки предпросмотр не показывается. Выберите конкретный файл.");
                return;
            }

            var resolvedPath = ResolveDocumentPath(node);
            if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath))
            {
                if (infoPanel != null)
                    infoPanel.Visibility = Visibility.Visible;
                previewContainer.Visibility = Visibility.Collapsed;
                ShowDocumentPreviewPlaceholder(browser, placeholder, statusText, "Файл не найден по сохраненному пути.");
                return;
            }

            if (infoPanel != null)
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
                if (TryOpenDocumentInExternalApp(resolvedPath, out var fallbackError))
                {
                    ShowDocumentPreviewPlaceholder(
                        browser,
                        placeholder,
                        statusText,
                        "Встроенный предпросмотр недоступен. Файл открыт во внешнем приложении.");
                }
                else
                {
                    ShowDocumentPreviewPlaceholder(
                        browser,
                        placeholder,
                        statusText,
                        $"Не удалось открыть предпросмотр: {ex.Message}{Environment.NewLine}{fallbackError}");
                }
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

        private void DocumentPreviewBrowser_LoadCompleted(object sender, NavigationEventArgs e)
        {
            if (ReferenceEquals(sender, PdfPreviewBrowser))
            {
                RestoreBrowserPreviewState(selectedPdfNode, PdfPreviewBrowser);
                return;
            }

            if (ReferenceEquals(sender, PdfPreviewBrowserSecondary))
            {
                RestoreBrowserPreviewState(secondaryPdfTab, PdfPreviewBrowserSecondary);
                return;
            }

            if (ReferenceEquals(sender, EstimatePreviewBrowser))
                RestoreBrowserPreviewState(selectedEstimateNode, EstimatePreviewBrowser);
        }

        private static void SaveBrowserPreviewState(DocumentTreeNode node, WebBrowser browser)
        {
            if (node == null || browser == null || browser.Visibility != Visibility.Visible)
                return;

            try
            {
                dynamic doc = browser.Document;
                if (doc == null)
                    return;

                dynamic root = doc.documentElement;
                if (root == null)
                    return;

                node.PreviewScrollX = SafeToInt(root.scrollLeft);
                node.PreviewScrollY = SafeToInt(root.scrollTop);
            }
            catch
            {
                // Some embedded engines do not expose DOM scroll state.
            }
        }

        private static void RestoreBrowserPreviewState(DocumentTreeNode node, WebBrowser browser)
        {
            if (node == null || browser == null || browser.Visibility != Visibility.Visible)
                return;

            try
            {
                dynamic doc = browser.Document;
                if (doc == null)
                    return;

                dynamic window = doc.parentWindow;
                if (window != null)
                    window.scrollTo(node.PreviewScrollX, node.PreviewScrollY);
            }
            catch
            {
                // Ignore unsupported script host operations.
            }
        }

        private static int SafeToInt(object value, int fallback = 0)
        {
            if (value == null)
                return fallback;

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return fallback;
            }
        }

        private static double SafeToDouble(object value, double fallback = 0)
        {
            if (value == null)
                return fallback;

            try
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return fallback;
            }
        }

        private void ShowEmbeddedEstimateWorkbook(DocumentTreeNode node, string filePath)
        {
            InitializeEstimatePreviewHost();
            if (EstimateExcelHost == null)
                throw new InvalidOperationException("Область предпросмотра Excel не готова.");

            if (useExternalSpreadsheetEditor)
            {
                ShowEmbeddedEstimateInExternalEditor(filePath);
                return;
            }

            if (string.Equals(estimateEmbeddedFilePath, filePath, StringComparison.CurrentCultureIgnoreCase)
                && estimateExcelWindowHandle != IntPtr.Zero)
            {
                EstimateExcelHost.Visibility = Visibility.Visible;
                EstimatePreviewBrowser.Visibility = Visibility.Collapsed;
                EstimatePreviewPlaceholder.Visibility = Visibility.Collapsed;
                RestoreEstimatePreviewState(node);
                ScheduleEstimateEmbeddedLayout();
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
                estimateEmbeddedFilePath = NormalizeDocumentPathKey(filePath);
                estimateExcelWindowHandle = new IntPtr((int)excelApp.Hwnd);
                if (estimateExcelWindowHandle == IntPtr.Zero)
                    throw new InvalidOperationException("Excel не предоставил окно для встраивания.");

                ConfigureEstimateExcelLiteUi(excelApp);
                RestoreEstimatePreviewState(node);

                ConfigureEmbeddedWindow(estimateExcelWindowHandle, EnsureEstimatePreviewHostHandle());
                ResetEstimateEmbeddedLayoutCache();
                EstimateExcelHost.Visibility = Visibility.Visible;
                EstimatePreviewBrowser.Visibility = Visibility.Collapsed;
                EstimatePreviewPlaceholder.Visibility = Visibility.Collapsed;
                ScheduleEstimateEmbeddedLayout();
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (EstimateExcelHost?.Visibility == Visibility.Visible && estimateExcelWindowHandle != IntPtr.Zero)
                        ActivateEmbeddedEstimateWindow();
                }), DispatcherPriority.Background);
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

        private void SaveEstimatePreviewState()
        {
            var targetNode = activeEstimateTab ?? selectedEstimateNode;
            if (targetNode == null || estimateExcelApplication == null || estimateExcelWorkbook == null)
                return;

            try
            {
                dynamic excelApp = estimateExcelApplication;
                dynamic window = excelApp.ActiveWindow;
                if (window != null)
                {
                targetNode.PreviewZoom = SafeToDouble(window.Zoom, targetNode.PreviewZoom);
                targetNode.PreviewScrollY = Math.Max(0, SafeToInt(window.ScrollRow, targetNode.PreviewScrollY));
                targetNode.PreviewScrollX = Math.Max(0, SafeToInt(window.ScrollColumn, targetNode.PreviewScrollX));
                }

                dynamic activeSheet = excelApp.ActiveSheet;
                if (activeSheet != null)
                    targetNode.PreviewSheetName = activeSheet.Name?.ToString() ?? string.Empty;

                targetNode.HashVerifiedAtUtc ??= DateTime.UtcNow;
            }
            catch
            {
                // ignore state-read errors from COM
            }
        }

        private void RestoreEstimatePreviewState(DocumentTreeNode node)
        {
            if (node == null || estimateExcelApplication == null || estimateExcelWorkbook == null)
                return;

            try
            {
                dynamic excelApp = estimateExcelApplication;
                dynamic workbook = estimateExcelWorkbook;

                if (!string.IsNullOrWhiteSpace(node.PreviewSheetName))
                {
                    try
                    {
                        dynamic sheet = workbook.Worksheets[node.PreviewSheetName];
                        sheet?.Activate();
                    }
                    catch
                    {
                        // keep current sheet if stored one is missing
                    }
                }

                dynamic window = excelApp.ActiveWindow;
                if (window != null)
                {
                    var zoom = Math.Clamp((int)Math.Round(node.PreviewZoom <= 0 ? 100 : node.PreviewZoom), 10, 400);
                    window.Zoom = zoom;
                    if (node.PreviewScrollY > 0)
                        window.ScrollRow = node.PreviewScrollY;
                    if (node.PreviewScrollX > 0)
                        window.ScrollColumn = node.PreviewScrollX;
                }
            }
            catch
            {
                // ignore state-apply errors from COM
            }
        }

        private void StopEstimateEmbeddedPreview()
        {
            SaveEstimatePreviewState();
            if (EstimateExcelHost != null)
                EstimateExcelHost.Visibility = Visibility.Collapsed;

            if (useExternalSpreadsheetEditor)
            {
                HideActiveExternalSpreadsheetInstance(isSecondary: false);
                return;
            }

            CloseEstimateWorkbook(saveChanges: true);
            ResetEstimateEmbeddedLayoutCache();
        }

        private void StopEstimateEmbeddedSecondaryPreview()
        {
            if (useExternalSpreadsheetEditor)
            {
                HideActiveExternalSpreadsheetInstance(isSecondary: true);
                return;
            }

            CloseEstimateWorkbookSecondary(saveChanges: true);
            ResetEstimateEmbeddedLayoutCacheSecondary();
        }

        private void HideEstimateEmbeddedPreview()
        {
            SaveEstimatePreviewState();

            if (useExternalSpreadsheetEditor && estimateExcelWindowHandle != IntPtr.Zero)
            {
                if (EstimateExcelHost != null)
                    EstimateExcelHost.Visibility = Visibility.Collapsed;

                try
                {
                    if (!estimateDetached)
                        ShowWindow(estimateExcelWindowHandle, SW_HIDE);
                }
                catch
                {
                    // Ignore visibility errors for the detached estimate window.
                }
            }
        }

        private void HideEstimateEmbeddedSecondaryPreview()
        {
            if (useExternalSpreadsheetEditor && estimateExcelWindowHandleSecondary != IntPtr.Zero)
            {
                try
                {
                    if (!estimateSecondaryDetached)
                        ShowWindow(estimateExcelWindowHandleSecondary, SW_HIDE);
                }
                catch
                {
                    // Ignore visibility errors.
                }
            }
        }

        private void ApplyExternalSpreadsheetInstance(ExternalSpreadsheetInstance instance, bool isSecondary)
        {
            if (instance == null)
                return;

            if (isSecondary)
            {
                estimateSpreadsheetProcessSecondary = instance.Process;
                estimateExcelWindowHandleSecondary = instance.Handle;
                estimateEmbeddedFilePathSecondary = instance.FilePath;
                activeExternalSpreadsheetInstanceSecondary = instance;
            }
            else
            {
                estimateSpreadsheetProcess = instance.Process;
                estimateExcelWindowHandle = instance.Handle;
                estimateEmbeddedFilePath = instance.FilePath;
                activeExternalSpreadsheetInstance = instance;
            }

            ConfigureFloatingEstimateWindow(instance.Handle);
            if (isSecondary)
                ResetEstimateEmbeddedLayoutCacheSecondary();
            else
                ResetEstimateEmbeddedLayoutCache();
        }

        private bool TryGetExternalSpreadsheetInstance(Dictionary<string, ExternalSpreadsheetInstance> cache, string filePath, out ExternalSpreadsheetInstance instance)
        {
            instance = null;
            if (cache == null || string.IsNullOrWhiteSpace(filePath))
                return false;

            if (!cache.TryGetValue(filePath, out var cached))
                return false;

            if (cached == null || cached.Handle == IntPtr.Zero)
            {
                cache.Remove(filePath);
                return false;
            }

            try
            {
                cached.Process?.Refresh();
                if (cached.Process != null && cached.Process.HasExited)
                {
                    cache.Remove(filePath);
                    return false;
                }
            }
            catch
            {
                cache.Remove(filePath);
                return false;
            }

            instance = cached;
            return true;
        }

        private void HideActiveExternalSpreadsheetInstance(bool isSecondary)
        {
            var instance = isSecondary ? activeExternalSpreadsheetInstanceSecondary : activeExternalSpreadsheetInstance;
            if (instance == null || instance.Handle == IntPtr.Zero)
                return;

            try
            {
                if (!isSecondary && estimateDetached)
                    return;
                if (isSecondary && estimateSecondaryDetached)
                    return;

                ShowWindow(instance.Handle, SW_HIDE);
            }
            catch
            {
                // ignore
            }
        }

        private void CloseAllExternalSpreadsheetInstances(Dictionary<string, ExternalSpreadsheetInstance> cache, ref ExternalSpreadsheetInstance activeInstance)
        {
            if (cache == null || cache.Count == 0)
                return;

            foreach (var instance in cache.Values.ToList())
            {
                if (instance?.Process == null)
                    continue;

                try
                {
                    if (!instance.Process.HasExited)
                    {
                        instance.Process.CloseMainWindow();
                        if (!instance.Process.WaitForExit(1500))
                            instance.Process.Kill();
                    }
                }
                catch
                {
                    // ignore shutdown errors
                }
                finally
                {
                    try { instance.Process.Dispose(); } catch { }
                }
            }

            cache.Clear();
            activeInstance = null;
        }

        private void CloseExternalSpreadsheetInstance(Dictionary<string, ExternalSpreadsheetInstance> cache, string filePath, ref ExternalSpreadsheetInstance activeInstance)
        {
            if (cache == null || string.IsNullOrWhiteSpace(filePath))
                return;

            if (!cache.TryGetValue(filePath, out var instance) || instance?.Process == null)
                return;

            try
            {
                if (!instance.Process.HasExited)
                {
                    instance.Process.CloseMainWindow();
                    if (!instance.Process.WaitForExit(1500))
                        instance.Process.Kill();
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                try { instance.Process.Dispose(); } catch { }
            }

            cache.Remove(filePath);
            if (ReferenceEquals(activeInstance, instance))
                activeInstance = null;
        }

        private void ReleaseStuckModifierKeys()
        {
            static bool IsDown(int key) => (GetAsyncKeyState(key) & 0x8000) != 0;
            static void Release(byte key) => keybd_event(key, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);

            if (IsDown(VK_SHIFT) || IsDown(VK_LSHIFT) || IsDown(VK_RSHIFT))
            {
                Release(VK_LSHIFT);
                Release(VK_RSHIFT);
                Release(VK_SHIFT);
            }

            if (IsDown(VK_CONTROL) || IsDown(VK_LCONTROL) || IsDown(VK_RCONTROL))
            {
                Release(VK_LCONTROL);
                Release(VK_RCONTROL);
                Release(VK_CONTROL);
            }

            if (IsDown(VK_MENU) || IsDown(VK_LMENU) || IsDown(VK_RMENU))
            {
                Release(VK_LMENU);
                Release(VK_RMENU);
                Release(VK_MENU);
            }
        }

        private static void PressAndReleaseKey(byte key)
        {
            keybd_event(key, 0, 0, UIntPtr.Zero);
            keybd_event(key, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        private void ResetEmbeddedEstimateInputState()
        {
            try { ReleaseCapture(); } catch { }

            if (estimateExcelWindowHandle != IntPtr.Zero)
            {
                try { SendMessage(estimateExcelWindowHandle, WM_CANCELMODE, IntPtr.Zero, IntPtr.Zero); } catch { }
            }

            try
            {
                if (estimateExcelApplication != null)
                {
                    dynamic excelApp = estimateExcelApplication;
                    try { excelApp.CutCopyMode = false; } catch { }
                    try { excelApp.SendKeys("{ESC}", false); } catch { }
                }
            }
            catch
            {
                // Ignore compatibility issues for non-Excel editors.
            }

            try { PressAndReleaseKey(VK_ESCAPE); } catch { }
        }

        private void ActivateEmbeddedEstimateWindow()
        {
            if (estimateExcelWindowHandle == IntPtr.Zero)
                return;

            try { Mouse.Capture(null); } catch { }

            try
            {
                SetFocus(estimateExcelWindowHandle);
            }
            catch
            {
                // Ignore focus errors.
            }
        }

        private void EstimateExcelHost_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            ActivateEmbeddedEstimateWindow();
        }

        private void EstimateExcelHost_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            ActivateEmbeddedEstimateWindow();
        }

        private void ResetEstimateEmbeddedLayoutCache()
        {
            estimateEmbeddedWindowX = int.MinValue;
            estimateEmbeddedWindowY = int.MinValue;
            estimateEmbeddedWindowWidth = -1;
            estimateEmbeddedWindowHeight = -1;
        }

        private void ResetEstimateEmbeddedLayoutCacheSecondary()
        {
            estimateEmbeddedWindowXSecondary = int.MinValue;
            estimateEmbeddedWindowYSecondary = int.MinValue;
            estimateEmbeddedWindowWidthSecondary = -1;
            estimateEmbeddedWindowHeightSecondary = -1;
        }

        private void ScheduleEstimateEmbeddedLayout()
        {
            if (estimateExcelWindowHandle == IntPtr.Zero)
                return;

            LayoutEmbeddedEstimateWindow(force: true);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (estimateExcelWindowHandle != IntPtr.Zero)
                    LayoutEmbeddedEstimateWindow();
            }), DispatcherPriority.Render);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (estimateExcelWindowHandle != IntPtr.Zero)
                    LayoutEmbeddedEstimateWindow();
            }), DispatcherPriority.ApplicationIdle);
        }

        private void ScheduleEstimateEmbeddedLayoutSecondary()
        {
            if (estimateExcelWindowHandleSecondary == IntPtr.Zero)
                return;

            LayoutEmbeddedEstimateWindowSecondary(force: true);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (estimateExcelWindowHandleSecondary != IntPtr.Zero)
                    LayoutEmbeddedEstimateWindowSecondary();
            }), DispatcherPriority.Render);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (estimateExcelWindowHandleSecondary != IntPtr.Zero)
                    LayoutEmbeddedEstimateWindowSecondary();
            }), DispatcherPriority.ApplicationIdle);
        }

        private void LayoutEmbeddedEstimateWindow(bool force = false)
        {
            if (estimateExcelWindowHandle == IntPtr.Zero)
                return;

            if (estimateDetached)
                return;

            int width;
            int height;
            int x;
            int y;

            if (useExternalSpreadsheetEditor)
            {
                if (EstimatePreviewContainer == null)
                    return;

                if (!IsVisible
                    || WindowState == WindowState.Minimized
                    || !ReferenceEquals(MainTabs?.SelectedItem, EstimateTab)
                    || EstimatePreviewContainer.Visibility != Visibility.Visible)
                {
                    try
                    {
                        ShowWindow(estimateExcelWindowHandle, SW_HIDE);
                    }
                    catch
                    {
                        // Ignore visibility errors for the detached estimate window.
                    }

                    return;
                }

                EstimatePreviewContainer.UpdateLayout();
                var screenBounds = GetScreenBounds(EstimatePreviewContainer);
                width = screenBounds.Width;
                height = screenBounds.Height;
                if (width <= 0 || height <= 0)
                {
                    try
                    {
                        ShowWindow(estimateExcelWindowHandle, SW_HIDE);
                    }
                    catch
                    {
                        // Ignore visibility errors for the detached estimate window.
                    }

                    return;
                }

                x = screenBounds.X;
                y = screenBounds.Y;
            }
            else
            {
                if (EstimateExcelHost == null || EstimateExcelHost.HostHandle == IntPtr.Zero)
                    return;

                width = Math.Max(0, (int)Math.Round(EstimateExcelHost.ActualWidth));
                var hostHeight = Math.Max(0, (int)Math.Round(EstimateExcelHost.ActualHeight));
                var topTrim = hostHeight > 120 ? EmbeddedExcelTopTrim : 0;
                height = Math.Max(0, hostHeight + topTrim);
                x = 0;
                y = -topTrim;
            }

            if (!force
                && x == estimateEmbeddedWindowX
                && y == estimateEmbeddedWindowY
                && width == estimateEmbeddedWindowWidth
                && height == estimateEmbeddedWindowHeight)
            {
                return;
            }

            estimateEmbeddedWindowX = x;
            estimateEmbeddedWindowY = y;
            estimateEmbeddedWindowWidth = width;
            estimateEmbeddedWindowHeight = height;

            var flags = SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_SHOWWINDOW;
            if (useExternalSpreadsheetEditor)
                flags |= SWP_NOACTIVATE;
            if (force)
                flags |= SWP_FRAMECHANGED;
            SetWindowPos(
                estimateExcelWindowHandle,
                IntPtr.Zero,
                x,
                y,
                width,
                height,
                flags);

            if (!useExternalSpreadsheetEditor)
                ShowWindow(estimateExcelWindowHandle, SW_SHOW);
        }

        private void LayoutEmbeddedEstimateWindowSecondary(bool force = false)
        {
            if (estimateExcelWindowHandleSecondary == IntPtr.Zero)
                return;

            if (estimateSecondaryDetached)
                return;

            int width;
            int height;
            int x;
            int y;

            if (EstimatePreviewContainerSecondary == null)
                return;

            if (!IsVisible
                || WindowState == WindowState.Minimized
                || !ReferenceEquals(MainTabs?.SelectedItem, EstimateTab)
                || EstimatePreviewContainerSecondary.Visibility != Visibility.Visible
                || !estimateSplitEnabled)
            {
                try { ShowWindow(estimateExcelWindowHandleSecondary, SW_HIDE); } catch { }
                return;
            }

            EstimatePreviewContainerSecondary.UpdateLayout();
            var screenBounds = GetScreenBounds(EstimatePreviewContainerSecondary);
            width = screenBounds.Width;
            height = screenBounds.Height;
            x = screenBounds.X;
            y = screenBounds.Y;

            if (!force
                && x == estimateEmbeddedWindowXSecondary
                && y == estimateEmbeddedWindowYSecondary
                && width == estimateEmbeddedWindowWidthSecondary
                && height == estimateEmbeddedWindowHeightSecondary)
                return;

            estimateEmbeddedWindowXSecondary = x;
            estimateEmbeddedWindowYSecondary = y;
            estimateEmbeddedWindowWidthSecondary = width;
            estimateEmbeddedWindowHeightSecondary = height;

            SetWindowPos(
                estimateExcelWindowHandleSecondary,
                IntPtr.Zero,
                x,
                y,
                width,
                height,
                SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_NOACTIVATE);

            try { ShowWindow(estimateExcelWindowHandleSecondary, SW_SHOW); } catch { }
        }

        private void ConfigureFloatingEstimateWindow(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero)
                return;

            try
            {
                ShowWindow(windowHandle, SW_HIDE);
            }
            catch
            {
                // Ignore visibility errors while preparing the pseudo-embedded window.
            }

            var ownerHandle = new WindowInteropHelper(this).Handle;
            var style = GetWindowLongPtr(windowHandle, GWL_STYLE).ToInt64();
            style &= ~(WS_CAPTION | WS_DLGFRAME | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU | WS_CHILD);
            style |= WS_POPUP;
            SetWindowLongPtr(windowHandle, GWL_STYLE, new IntPtr(style));

            var exStyle = GetWindowLongPtr(windowHandle, GWL_EXSTYLE).ToInt64();
            exStyle &= ~WS_EX_APPWINDOW;
            exStyle |= WS_EX_TOOLWINDOW;
            SetWindowLongPtr(windowHandle, GWL_EXSTYLE, new IntPtr(exStyle));

            if (ownerHandle != IntPtr.Zero)
                SetWindowLongPtr(windowHandle, GWLP_HWNDPARENT, ownerHandle);

            SetWindowPos(
                windowHandle,
                IntPtr.Zero,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        }

        private void ConfigureDetachedEstimateWindow(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero)
                return;

            var style = GetWindowLongPtr(windowHandle, GWL_STYLE).ToInt64();
            style &= ~WS_CHILD;
            style &= ~WS_POPUP;
            style |= WS_CAPTION | WS_DLGFRAME | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU;
            SetWindowLongPtr(windowHandle, GWL_STYLE, new IntPtr(style));

            var exStyle = GetWindowLongPtr(windowHandle, GWL_EXSTYLE).ToInt64();
            exStyle |= WS_EX_APPWINDOW;
            exStyle &= ~WS_EX_TOOLWINDOW;
            SetWindowLongPtr(windowHandle, GWL_EXSTYLE, new IntPtr(exStyle));

            SetWindowLongPtr(windowHandle, GWLP_HWNDPARENT, IntPtr.Zero);
            SetWindowPos(
                windowHandle,
                IntPtr.Zero,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_FRAMECHANGED);
        }

        private static (int X, int Y, int Width, int Height) GetScreenBounds(FrameworkElement element)
        {
            if (element == null || !element.IsLoaded)
                return (0, 0, 0, 0);

            var topLeft = element.PointToScreen(new Point(0, 0));
            var bottomRight = element.PointToScreen(new Point(element.ActualWidth, element.ActualHeight));

            var left = (int)Math.Round(Math.Min(topLeft.X, bottomRight.X));
            var top = (int)Math.Round(Math.Min(topLeft.Y, bottomRight.Y));
            var right = (int)Math.Round(Math.Max(topLeft.X, bottomRight.X));
            var bottom = (int)Math.Round(Math.Max(topLeft.Y, bottomRight.Y));

            return (left, top, Math.Max(0, right - left), Math.Max(0, bottom - top));
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
            SaveBrowserPreviewState(selectedPdfNode, PdfPreviewBrowser);
            selectedPdfNode = e.NewValue as DocumentTreeNode;
            if (selectedPdfNode != null && !selectedPdfNode.IsFolder)
                OpenPdfTab(selectedPdfNode);
            else
                UpdatePdfSelectionInfo();
        }

        private void EstimateTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            SaveEstimatePreviewState();
            SaveBrowserPreviewState(selectedEstimateNode, EstimatePreviewBrowser);
            selectedEstimateNode = e.NewValue as DocumentTreeNode;
            if (selectedEstimateNode != null && !selectedEstimateNode.IsFolder)
                OpenEstimateTab(selectedEstimateNode);
            else
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
                if (!TryCopyDocumentToStorage(file, refreshPdf, out var storedPath, out var storedRelativePath, out var contentHash, out var fileSize)
                    || string.IsNullOrWhiteSpace(storedRelativePath))
                {
                    copyFailedFiles.Add(file);
                    continue;
                }

                targetCollection.Add(new DocumentTreeNode
                {
                    Name = System.IO.Path.GetFileNameWithoutExtension(file),
                    FilePath = storedPath,
                    StoredRelativePath = storedRelativePath,
                    ContentHash = contentHash,
                    FileSizeBytes = fileSize,
                    HashVerifiedAtUtc = DateTime.UtcNow,
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
                    "Часть файлов не удалось скопировать во внутреннее хранилище проекта. Эти файлы не были добавлены.",
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

            var isMassDelete = selectedNode.IsFolder || (selectedNode.Children?.Count ?? 0) > 0;
            if (isMassDelete)
            {
                if (!EnsureCanRunCriticalOperation("массовое удаление узла документов"))
                    return;
            }
            else if (!EnsureCanEditOperation("удаление документа"))
            {
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

        private void OpenPdfLibraryReport_Click(object sender, RoutedEventArgs e)
            => OpenDocumentLibraryReport("Отчет по PDF-документам", currentObject?.PdfDocuments);

        private void OpenEstimateLibraryReport_Click(object sender, RoutedEventArgs e)
            => OpenDocumentLibraryReport("Отчет по сметам", currentObject?.EstimateDocuments);

        private void OpenAllDocumentsReport_Click(object sender, RoutedEventArgs e)
        {
            var rows = new List<DocumentLibraryReportRow>();
            rows.AddRange(BuildDocumentLibraryRows(currentObject?.PdfDocuments, "PDF"));
            rows.AddRange(BuildDocumentLibraryRows(currentObject?.EstimateDocuments, "Сметы"));
            ShowDocumentLibraryReport("Сводный отчет по документам", rows);
        }

        private void SearchPdfNodes_Click(object sender, RoutedEventArgs e)
            => ShowDocumentTreeSearchDialog("Поиск по PDF", currentObject?.PdfDocuments, "PDF", node =>
            {
                selectedPdfNode = node;
                SelectMainTab(PdfTab);
                if (node != null && !node.IsFolder)
                    OpenPdfTab(node);
                else
                    UpdatePdfSelectionInfo();
            });

        private void SearchEstimateNodes_Click(object sender, RoutedEventArgs e)
            => ShowDocumentTreeSearchDialog("Поиск по сметам", currentObject?.EstimateDocuments, "Сметы", node =>
            {
                selectedEstimateNode = node;
                SelectMainTab(EstimateTab);
                if (node != null && !node.IsFolder)
                    OpenEstimateTab(node);
                else
                    UpdateEstimateSelectionInfo();
            });

        private void PdfTreeSearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                OpenDocumentTreeSearchFromInline(true);
                e.Handled = true;
            }
        }

        private void PdfTreeSearchButton_Click(object sender, RoutedEventArgs e)
            => OpenDocumentTreeSearchFromInline(true);

        private void PdfTreeSearchClear_Click(object sender, RoutedEventArgs e)
        {
            if (PdfTreeSearchBox != null)
                PdfTreeSearchBox.Text = string.Empty;
        }

        private void EstimateTreeSearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                OpenDocumentTreeSearchFromInline(false);
                e.Handled = true;
            }
        }

        private void EstimateTreeSearchButton_Click(object sender, RoutedEventArgs e)
            => OpenDocumentTreeSearchFromInline(false);

        private void EstimateTreeSearchClear_Click(object sender, RoutedEventArgs e)
        {
            if (EstimateTreeSearchBox != null)
                EstimateTreeSearchBox.Text = string.Empty;
        }

        private void OpenDocumentTreeSearchFromInline(bool isPdf)
        {
            var query = (isPdf ? PdfTreeSearchBox?.Text : EstimateTreeSearchBox?.Text)?.Trim() ?? string.Empty;
            var title = isPdf ? "Поиск по PDF" : "Поиск по сметам";
            var libraryName = isPdf ? "PDF" : "Сметы";
            var root = isPdf ? currentObject?.PdfDocuments : currentObject?.EstimateDocuments;

            ShowDocumentTreeSearchDialog(title, root, libraryName, node =>
            {
                if (isPdf)
                {
                    selectedPdfNode = node;
                    SelectMainTab(PdfTab);
                    if (node != null && !node.IsFolder)
                        OpenPdfTab(node);
                    else
                        UpdatePdfSelectionInfo();
                }
                else
                {
                    selectedEstimateNode = node;
                    SelectMainTab(EstimateTab);
                    if (node != null && !node.IsFolder)
                        OpenEstimateTab(node);
                    else
                        UpdateEstimateSelectionInfo();
                }
            }, query);
        }

        private void BatchRenamePdfNodes_Click(object sender, RoutedEventArgs e)
            => BatchRenameDocumentNodes(currentObject?.PdfDocuments, selectedPdfNode, true);

        private void BatchRenameEstimateNodes_Click(object sender, RoutedEventArgs e)
            => BatchRenameDocumentNodes(currentObject?.EstimateDocuments, selectedEstimateNode, false);

        private void ReorganizePdfNodes_Click(object sender, RoutedEventArgs e)
            => ReorganizeDocumentNodesByExtension(currentObject?.PdfDocuments, selectedPdfNode, true);

        private void ReorganizeEstimateNodes_Click(object sender, RoutedEventArgs e)
            => ReorganizeDocumentNodesByExtension(currentObject?.EstimateDocuments, selectedEstimateNode, false);

        private void CheckPdfDocumentsIntegrity_Click(object sender, RoutedEventArgs e)
            => ShowDocumentIntegrityReport("Проверка целостности PDF", currentObject?.PdfDocuments);

        private void CheckEstimateDocumentsIntegrity_Click(object sender, RoutedEventArgs e)
            => ShowDocumentIntegrityReport("Проверка целостности смет", currentObject?.EstimateDocuments);

        private void RunDocumentStorageMaintenance_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanRunCriticalOperation("обслуживание хранилища документов"))
                return;

            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            var confirm = MessageBox.Show(
                "Выполнить обслуживание хранилища документов?\n\nБудет выполнено:\n- восстановление путей по хэшу/имени;\n- пересчет хэшей и метаданных;\n- дедупликация ссылок на одинаковые файлы;\n- удаление неиспользуемых файлов из внутреннего хранилища.",
                "Обслуживание хранилища",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes)
                return;

            var result = RunDocumentStorageMaintenanceCore();
            SaveState(SaveTrigger.System);
            RefreshDocumentLibraries();

            var summary = $"Проверено файлов: {result.NodesVisited}\n" +
                          $"Восстановлено путей: {result.PathsRecovered}\n" +
                          $"Обновлено хэшей/размеров: {result.MetadataUpdated}\n" +
                          $"Ссылок перекинуто на каноничные файлы: {result.LinksRepointed}\n" +
                          $"Уникальных файлов по хэшу: {result.HashIndexSize}\n" +
                          $"Удалено дубликатов в хранилище: {result.DuplicateFilesRemoved}\n" +
                          $"Удалено неиспользуемых файлов: {result.OrphanFilesRemoved}";

            if (result.Errors.Count > 0)
            {
                summary += $"\n\nОшибок при обслуживании: {result.Errors.Count}\n" +
                           string.Join("\n", result.Errors.Take(8));
                if (result.Errors.Count > 8)
                    summary += "\n...";
            }

            MessageBox.Show(summary, "Обслуживание хранилища", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private DocumentStorageMaintenanceResult RunDocumentStorageMaintenanceCore()
        {
            var result = new DocumentStorageMaintenanceResult();
            EnsureDocumentLibraries();

            var fileNodes = EnumerateDocumentFileNodes(currentObject?.PdfDocuments)
                .Concat(EnumerateDocumentFileNodes(currentObject?.EstimateDocuments))
                .ToList();

            result.NodesVisited = fileNodes.Count;
            var referencedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var canonicalByHash = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var duplicatePhysicalCandidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var node in fileNodes)
            {
                var originalPath = node.FilePath ?? string.Empty;
                var resolved = ResolveDocumentPath(node);
                if (!string.IsNullOrWhiteSpace(resolved)
                    && !string.Equals(originalPath, resolved, StringComparison.OrdinalIgnoreCase))
                {
                    result.PathsRecovered++;
                }

                if (string.IsNullOrWhiteSpace(resolved) || !File.Exists(resolved))
                    continue;

                var fullResolved = System.IO.Path.GetFullPath(resolved);

                var oldHash = node.ContentHash ?? string.Empty;
                var oldSize = node.FileSizeBytes;
                EnsureDocumentNodeFileMetadata(node, fullResolved);
                if (!string.Equals(oldHash, node.ContentHash ?? string.Empty, StringComparison.OrdinalIgnoreCase)
                    || oldSize != node.FileSizeBytes)
                {
                    result.MetadataUpdated++;
                }

                var hash = node.ContentHash ?? string.Empty;
                if (string.IsNullOrWhiteSpace(hash))
                    continue;

                if (!canonicalByHash.TryGetValue(hash, out var canonicalPath))
                {
                    canonicalByHash[hash] = fullResolved;
                    continue;
                }

                var fullCanonical = System.IO.Path.GetFullPath(canonicalPath);
                if (string.Equals(fullCanonical, fullResolved, StringComparison.OrdinalIgnoreCase))
                    continue;

                node.FilePath = fullCanonical;
                var storageRoot = GetProjectStorageRoot(createIfMissing: false);
                if (!string.IsNullOrWhiteSpace(storageRoot))
                {
                    try
                    {
                        node.StoredRelativePath = System.IO.Path.GetRelativePath(storageRoot, fullCanonical);
                    }
                    catch
                    {
                        // keep existing relative path if conversion failed
                    }
                }

                duplicatePhysicalCandidates.Add(fullResolved);
                result.LinksRepointed++;
            }

            result.HashIndexSize = canonicalByHash.Count;
            referencedPaths.Clear();
            foreach (var node in fileNodes)
            {
                var path = ResolveDocumentPath(node);
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    continue;

                referencedPaths.Add(System.IO.Path.GetFullPath(path));
            }

            var storageRootPath = GetProjectStorageRoot(createIfMissing: false);
            var storageRootFullPath = string.IsNullOrWhiteSpace(storageRootPath)
                ? string.Empty
                : System.IO.Path.GetFullPath(storageRootPath);

            foreach (var duplicateFile in duplicatePhysicalCandidates)
            {
                if (string.IsNullOrWhiteSpace(storageRootFullPath)
                    || !IsPathUnderRoot(duplicateFile, storageRootFullPath)
                    || referencedPaths.Contains(duplicateFile)
                    || !File.Exists(duplicateFile))
                    continue;

                try
                {
                    File.Delete(duplicateFile);
                    result.DuplicateFilesRemoved++;
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"Не удалось удалить дубликат: {duplicateFile} ({ex.Message})");
                }
            }

            if (!string.IsNullOrWhiteSpace(storageRootFullPath) && Directory.Exists(storageRootFullPath))
            {
                foreach (var file in Directory.EnumerateFiles(storageRootFullPath, "*", SearchOption.AllDirectories))
                {
                    var full = System.IO.Path.GetFullPath(file);
                    if (referencedPaths.Contains(full))
                        continue;

                    try
                    {
                        File.Delete(full);
                        result.OrphanFilesRemoved++;
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add($"Не удалось удалить неиспользуемый файл: {full} ({ex.Message})");
                    }
                }

                foreach (var directory in Directory.EnumerateDirectories(storageRootFullPath, "*", SearchOption.AllDirectories)
                             .OrderByDescending(path => path.Length))
                {
                    try
                    {
                        if (!Directory.EnumerateFileSystemEntries(directory).Any())
                            Directory.Delete(directory);
                    }
                    catch
                    {
                        // best effort cleanup
                    }
                }
            }

            RebuildDocumentHashPathCache();
            return result;
        }

        private static bool IsPathUnderRoot(string path, string root)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(root))
                return false;

            try
            {
                var fullPath = System.IO.Path.GetFullPath(path);
                var fullRoot = System.IO.Path.GetFullPath(root);
                if (!fullRoot.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                    fullRoot += System.IO.Path.DirectorySeparatorChar;

                return fullPath.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static IEnumerable<DocumentTreeNode> EnumerateDocumentFileNodes(IEnumerable<DocumentTreeNode> roots)
        {
            if (roots == null)
                yield break;

            var stack = new Stack<DocumentTreeNode>(roots.Where(x => x != null).Reverse());
            while (stack.Count > 0)
            {
                var node = stack.Pop();
                if (node == null)
                    continue;

                if (!node.IsFolder)
                    yield return node;

                if (node.Children == null || node.Children.Count == 0)
                    continue;

                for (var i = node.Children.Count - 1; i >= 0; i--)
                {
                    var child = node.Children[i];
                    if (child != null)
                        stack.Push(child);
                }
            }
        }

        private void ShowDocumentTreeSearchDialog(string title, List<DocumentTreeNode> root, string libraryName, Action<DocumentTreeNode> navigate, string initialQuery = null)
        {
            if (root == null || root.Count == 0)
            {
                MessageBox.Show("Дерево документов пусто.");
                return;
            }

            var sourceRows = BuildDocumentTreeSearchRows(root, libraryName);
            if (sourceRows.Count == 0)
            {
                MessageBox.Show("В дереве нет узлов для поиска.");
                return;
            }

            var querySeed = initialQuery?.Trim() ?? string.Empty;
            var filteredSeed = FilterDocumentSearchRows(sourceRows, querySeed);
            var resultRows = new ObservableCollection<DocumentTreeSearchRow>(filteredSeed);
            var dialog = new Window
            {
                Title = title,
                Owner = this,
                Width = 1000,
                Height = 680,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var rootGrid = new Grid { Margin = new Thickness(14) };
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            dialog.Content = rootGrid;

            var searchBox = new TextBox
            {
                Tag = "Введите часть названия узла или файла",
                Margin = new Thickness(0, 0, 0, 8),
                Text = querySeed
            };
            rootGrid.Children.Add(searchBox);

            var grid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                ItemsSource = resultRows,
                ColumnWidth = DataGridLength.Auto
            };
            grid.Columns.Add(new DataGridTextColumn { Header = "Раздел", Binding = new Binding(nameof(DocumentTreeSearchRow.Library)), Width = 90 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Узел", Binding = new Binding(nameof(DocumentTreeSearchRow.Name)), Width = 260 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Тип", Binding = new Binding(nameof(DocumentTreeSearchRow.NodeType)), Width = 90 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Путь в дереве", Binding = new Binding(nameof(DocumentTreeSearchRow.Path)), Width = 300 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Путь к файлу", Binding = new Binding(nameof(DocumentTreeSearchRow.FilePath)), Width = 280 });
            DataGridSizingHelper.SetEnableSmartSizing(grid, true);
            Grid.SetRow(grid, 1);
            rootGrid.Children.Add(grid);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            Grid.SetRow(footer, 2);
            rootGrid.Children.Add(footer);

            var openButton = new Button { Content = "Открыть", MinWidth = 120, IsDefault = true };
            var cancelButton = new Button { Content = "Закрыть", MinWidth = 120, IsCancel = true, Margin = new Thickness(8, 0, 0, 0), Style = FindResource("SecondaryButton") as Style };
            footer.Children.Add(openButton);
            footer.Children.Add(cancelButton);

            void ApplyFilter()
            {
                var query = searchBox.Text?.Trim() ?? string.Empty;
                var filtered = FilterDocumentSearchRows(sourceRows, query);

                resultRows.Clear();
                foreach (var row in filtered)
                    resultRows.Add(row);
            }

            searchBox.TextChanged += (_, _) => ApplyFilter();

            void NavigateSelected()
            {
                if (grid.SelectedItem is not DocumentTreeSearchRow selected || selected.Node == null)
                    return;

                navigate?.Invoke(selected.Node);
                dialog.DialogResult = true;
            }

            openButton.Click += (_, _) => NavigateSelected();
            grid.MouseDoubleClick += (_, _) => NavigateSelected();
            ApplyFilter();
            dialog.ShowDialog();
        }

        private static List<DocumentTreeSearchRow> FilterDocumentSearchRows(IEnumerable<DocumentTreeSearchRow> sourceRows, string query)
        {
            if (sourceRows == null)
                return new List<DocumentTreeSearchRow>();

            var trimmed = query?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
                return sourceRows.ToList();

            return sourceRows.Where(x =>
                    (x.Name?.Contains(trimmed, StringComparison.CurrentCultureIgnoreCase) ?? false)
                    || (x.Path?.Contains(trimmed, StringComparison.CurrentCultureIgnoreCase) ?? false)
                    || (x.FilePath?.Contains(trimmed, StringComparison.CurrentCultureIgnoreCase) ?? false))
                .ToList();
        }

        private List<DocumentTreeSearchRow> BuildDocumentTreeSearchRows(IEnumerable<DocumentTreeNode> nodes, string libraryName)
        {
            var rows = new List<DocumentTreeSearchRow>();
            if (nodes == null)
                return rows;

            void Collect(IEnumerable<DocumentTreeNode> currentNodes, string pathPrefix)
            {
                if (currentNodes == null)
                    return;

                foreach (var node in currentNodes)
                {
                    if (node == null)
                        continue;

                    var nodeName = string.IsNullOrWhiteSpace(node.Name) ? "Без названия" : node.Name.Trim();
                    var nodePath = string.IsNullOrWhiteSpace(pathPrefix) ? nodeName : $"{pathPrefix} / {nodeName}";
                    var resolvedFilePath = node.IsFolder ? string.Empty : ResolveDocumentPath(node);
                    rows.Add(new DocumentTreeSearchRow
                    {
                        Library = libraryName,
                        Name = nodeName,
                        NodeType = node.IsFolder ? "Папка" : "Файл",
                        Path = nodePath,
                        FilePath = resolvedFilePath,
                        Node = node
                    });

                    if (node.Children?.Count > 0)
                        Collect(node.Children, nodePath);
                }
            }

            Collect(nodes, string.Empty);
            return rows;
        }

        private void BatchRenameDocumentNodes(List<DocumentTreeNode> root, DocumentTreeNode selectedNode, bool isPdfLibrary)
        {
            if (root == null || root.Count == 0)
            {
                MessageBox.Show("Дерево документов пусто.");
                return;
            }

            List<DocumentTreeNode> targetCollection;
            if (selectedNode?.IsFolder == true)
            {
                targetCollection = selectedNode.Children ?? new List<DocumentTreeNode>();
            }
            else if (selectedNode != null)
            {
                targetCollection = GetOwningDocumentCollection(root, selectedNode) ?? root;
            }
            else
            {
                targetCollection = root;
            }

            var fileNodes = targetCollection.Where(x => x != null && !x.IsFolder).ToList();
            if (fileNodes.Count == 0)
            {
                MessageBox.Show("Для пакетного переименования нет файлов в выбранном разделе.");
                return;
            }

            var prefix = Microsoft.VisualBasic.Interaction.InputBox(
                "Введите префикс имени для пакетного переименования:",
                "Пакетное переименование",
                isPdfLibrary ? "PDF_" : "Смета_")?.Trim();
            if (string.IsNullOrWhiteSpace(prefix))
                return;

            var startText = Microsoft.VisualBasic.Interaction.InputBox(
                "Введите стартовый номер:",
                "Пакетное переименование",
                "1")?.Trim();
            if (!int.TryParse(startText, out var start))
                start = 1;
            start = Math.Max(1, start);

            var index = start;
            foreach (var node in fileNodes.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase))
            {
                node.Name = $"{prefix}{index:D3}";
                index++;
            }

            SaveState();
            RefreshDocumentLibraries();
            MessageBox.Show($"Переименовано файлов: {fileNodes.Count}.");
        }

        private void ReorganizeDocumentNodesByExtension(List<DocumentTreeNode> root, DocumentTreeNode selectedNode, bool isPdfLibrary)
        {
            if (root == null || root.Count == 0)
            {
                MessageBox.Show("Дерево документов пусто.");
                return;
            }

            List<DocumentTreeNode> targetCollection = selectedNode?.IsFolder == true
                ? (selectedNode.Children ?? new List<DocumentTreeNode>())
                : root;

            var directFiles = targetCollection.Where(x => x != null && !x.IsFolder).ToList();
            if (directFiles.Count == 0)
            {
                MessageBox.Show("Нет файлов верхнего уровня для реорганизации.");
                return;
            }

            var moved = 0;
            foreach (var fileNode in directFiles)
            {
                var resolved = ResolveDocumentPath(fileNode);
                var ext = System.IO.Path.GetExtension(resolved)?.Trim('.').ToUpperInvariant();
                if (string.IsNullOrWhiteSpace(ext))
                    ext = "ДРУГОЕ";

                var folder = targetCollection.FirstOrDefault(x => x.IsFolder && string.Equals(x.Name, ext, StringComparison.CurrentCultureIgnoreCase));
                if (folder == null)
                {
                    folder = new DocumentTreeNode
                    {
                        Name = ext,
                        IsFolder = true,
                        Children = new List<DocumentTreeNode>()
                    };
                    targetCollection.Add(folder);
                }

                targetCollection.Remove(fileNode);
                folder.Children ??= new List<DocumentTreeNode>();
                folder.Children.Add(fileNode);
                moved++;
            }

            SaveState();
            RefreshDocumentLibraries();
            MessageBox.Show($"Реорганизация завершена. Перемещено файлов: {moved}.");
        }

        private void ShowDocumentIntegrityReport(string title, IEnumerable<DocumentTreeNode> root)
        {
            var issues = BuildDocumentIntegrityIssues(root).ToList();
            var dialog = new Window
            {
                Title = title,
                Owner = this,
                Width = 1080,
                Height = 680,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var layout = new Grid { Margin = new Thickness(14) };
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            dialog.Content = layout;

            layout.Children.Add(new TextBlock
            {
                Text = issues.Count == 0 ? "Проблем целостности не найдено." : $"Найдено проблем: {issues.Count}",
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            });

            var grid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                ItemsSource = issues,
                ColumnWidth = DataGridLength.Auto
            };
            grid.Columns.Add(new DataGridTextColumn { Header = "Узел", Binding = new Binding(nameof(DocumentIntegrityIssueRow.NodeName)), Width = 280 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Проблема", Binding = new Binding(nameof(DocumentIntegrityIssueRow.Issue)), Width = 320 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Путь", Binding = new Binding(nameof(DocumentIntegrityIssueRow.Path)), Width = 420 });
            DataGridSizingHelper.SetEnableSmartSizing(grid, true);
            Grid.SetRow(grid, 1);
            layout.Children.Add(grid);

            dialog.ShowDialog();
        }

        private IEnumerable<DocumentIntegrityIssueRow> BuildDocumentIntegrityIssues(IEnumerable<DocumentTreeNode> root)
        {
            var issues = new List<DocumentIntegrityIssueRow>();
            var hashOwner = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            void Collect(IEnumerable<DocumentTreeNode> nodes)
            {
                if (nodes == null)
                    return;

                foreach (var node in nodes)
                {
                    if (node == null)
                        continue;

                    if (node.IsFolder)
                    {
                        Collect(node.Children);
                        continue;
                    }

                    var resolved = ResolveDocumentPath(node);
                    if (string.IsNullOrWhiteSpace(resolved) || !File.Exists(resolved))
                    {
                        issues.Add(new DocumentIntegrityIssueRow
                        {
                            NodeName = node.Name ?? string.Empty,
                            Path = node.FilePath ?? string.Empty,
                            Issue = "Файл не найден"
                        });
                        continue;
                    }

                    if (!TryComputeFileHash(resolved, out var hash, out var size))
                    {
                        issues.Add(new DocumentIntegrityIssueRow
                        {
                            NodeName = node.Name ?? string.Empty,
                            Path = resolved,
                            Issue = "Не удалось вычислить хэш"
                        });
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(node.ContentHash)
                        && !string.Equals(node.ContentHash, hash, StringComparison.OrdinalIgnoreCase))
                    {
                        issues.Add(new DocumentIntegrityIssueRow
                        {
                            NodeName = node.Name ?? string.Empty,
                            Path = resolved,
                            Issue = "Хэш не совпадает с сохраненным"
                        });
                    }

                    if (node.FileSizeBytes.HasValue && node.FileSizeBytes.Value != size)
                    {
                        issues.Add(new DocumentIntegrityIssueRow
                        {
                            NodeName = node.Name ?? string.Empty,
                            Path = resolved,
                            Issue = "Размер отличается от сохраненного"
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(hash))
                    {
                        if (hashOwner.TryGetValue(hash, out var existing))
                        {
                            issues.Add(new DocumentIntegrityIssueRow
                            {
                                NodeName = node.Name ?? string.Empty,
                                Path = resolved,
                                Issue = $"Дубликат содержимого (как \"{existing}\")"
                            });
                        }
                        else
                        {
                            hashOwner[hash] = node.Name ?? string.Empty;
                        }
                    }
                }
            }

            Collect(root);
            issues.AddRange(lastStorageIntegrityIssues.Select(x => new DocumentIntegrityIssueRow
            {
                NodeName = "Импорт/Экспорт",
                Path = string.Empty,
                Issue = x
            }));
            return issues;
        }

        private void OpenDocumentLibraryReport(string title, List<DocumentTreeNode> root)
        {
            var rows = BuildDocumentLibraryRows(root, string.Empty);
            ShowDocumentLibraryReport(title, rows);
        }

        private List<DocumentLibraryReportRow> BuildDocumentLibraryRows(List<DocumentTreeNode> root, string libraryName)
        {
            var rows = new List<DocumentLibraryReportRow>();
            if (root == null || root.Count == 0)
                return rows;

            void Collect(IEnumerable<DocumentTreeNode> nodes)
            {
                if (nodes == null)
                    return;

                foreach (var node in nodes)
                {
                    if (node == null)
                        continue;

                    if (node.IsFolder)
                    {
                        Collect(node.Children);
                        continue;
                    }

                    var resolved = ResolveDocumentPath(node);
                    var exists = !string.IsNullOrWhiteSpace(resolved) && File.Exists(resolved);
                    rows.Add(new DocumentLibraryReportRow
                    {
                        Library = libraryName,
                        NodeName = node.Name ?? string.Empty,
                        NodeType = "Файл",
                        FilePath = resolved ?? string.Empty,
                        Status = exists ? "Доступен" : "Файл не найден"
                    });
                }
            }

            Collect(root);
            return rows;
        }

        private void ShowDocumentLibraryReport(string title, List<DocumentLibraryReportRow> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                MessageBox.Show("Дерево документов пусто.");
                return;
            }

            var total = rows.Count;
            var missing = rows.Count(x => string.Equals(x.Status, "Файл не найден", StringComparison.CurrentCultureIgnoreCase));
            var available = total - missing;
            var showLibraryColumn = rows.Any(x => !string.IsNullOrWhiteSpace(x.Library));

            var dialog = new Window
            {
                Title = title,
                Owner = this,
                Width = 1080,
                Height = 680,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var layout = new Grid { Margin = new Thickness(14) };
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            dialog.Content = layout;

            layout.Children.Add(new TextBlock
            {
                Text = $"Всего файлов: {total}; доступно: {available}; отсутствует: {missing}",
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            });

            var dg = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                ItemsSource = rows,
                ColumnWidth = DataGridLength.Auto
            };
            if (showLibraryColumn)
                dg.Columns.Add(new DataGridTextColumn { Header = "Раздел", Binding = new Binding(nameof(DocumentLibraryReportRow.Library)), Width = 90 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Узел", Binding = new Binding(nameof(DocumentLibraryReportRow.NodeName)), Width = 280 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Тип", Binding = new Binding(nameof(DocumentLibraryReportRow.NodeType)), Width = 80 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Статус", Binding = new Binding(nameof(DocumentLibraryReportRow.Status)), Width = 140 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Путь", Binding = new Binding(nameof(DocumentLibraryReportRow.FilePath)), Width = showLibraryColumn ? 450 : 520 });
            DataGridSizingHelper.SetEnableSmartSizing(dg, true);
            Grid.SetRow(dg, 1);
            layout.Children.Add(dg);

            dialog.ShowDialog();
        }

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

            if (!TryOpenDocumentInExternalApp(resolvedPath, out var error))
            {
                MessageBox.Show(
                    $"Не удалось открыть файл во внешнем приложении.{Environment.NewLine}{error}",
                    "Открытие файла",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private bool TryOpenDocumentInExternalApp(string filePath, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                errorMessage = "Файл не найден.";
                return false;
            }

            var extension = IOPath.GetExtension(filePath)?.Trim().ToLowerInvariant() ?? string.Empty;
            var preferredExecutablePath = GetPreferredExecutablePathForDocument(extension);
            if (!string.IsNullOrWhiteSpace(preferredExecutablePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = preferredExecutablePath,
                        Arguments = $"\"{filePath}\"",
                        UseShellExecute = false
                    });
                    return true;
                }
                catch (Exception ex)
                {
                    errorMessage = ex.Message;
                }
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
                return true;
            }
            catch (Exception ex)
            {
                if (!string.IsNullOrWhiteSpace(errorMessage))
                    errorMessage = errorMessage + Environment.NewLine + ex.Message;
                else
                    errorMessage = ex.Message;
                return false;
            }
        }

        private string GetPreferredExecutablePathForDocument(string extension)
        {
            if (string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase))
                return preferredPdfEditorPath;

            if (IsSpreadsheetDocumentExtension(extension))
                return preferredSpreadsheetEditorPath;

            return string.Empty;
        }

        private static bool IsSpreadsheetDocumentExtension(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
                return false;

            return extension is ".xlsx" or ".xlsm" or ".xls" or ".csv" or ".tsv" or ".ods";
        }

        private void EstimatePreviewContainer_SizeChanged(object sender, SizeChangedEventArgs e)
            => LayoutEmbeddedEstimateWindow();

        private void EstimatePreviewContainerSecondary_SizeChanged(object sender, SizeChangedEventArgs e)
            => LayoutEmbeddedEstimateWindowSecondary();

        private void PdfPreviewContainer_SizeChanged(object sender, SizeChangedEventArgs e)
            => LayoutEmbeddedPdfWindow();

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
            SetTabDisplayMode("Приход", "Таблица");
            UpdateArrivalViewMode();
            RequestArrivalFilterRefresh(immediate: true);
        }

        private void ArrivalMatrixViewButton_Click(object sender, RoutedEventArgs e)
        {
            arrivalMatrixMode = true;
            SetTabDisplayMode("Приход", "Матрица");
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
            RequestArrivalFilterRefresh(immediate: true);
        }
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source is not TabControl tab || tab.SelectedItem is not TabItem item)
                return;

            var previousTab = e.RemovedItems.OfType<TabItem>().FirstOrDefault();
            if (previousTab != null)
                SaveGridColumnPreferences(GetGridByTabHeader(previousTab.Header?.ToString()));

            var tabStopwatch = Stopwatch.StartNew();
            EnsureTabInitialized(item);
            SetTabDisplayMode(ActiveTabModeKey, item.Header?.ToString() ?? string.Empty);
            UpdateTabButtons();
            UpdateTreePanelState(forceVisible: isTreePinned);
            UpdatePdfTreePanelState(forceVisible: isPdfTreePinned);
            UpdateEstimateTreePanelState(forceVisible: isEstimateTreePinned);
            UpdateAllJournalToolsPanelStates();

            if (!ReferenceEquals(item, EstimateTab))
            {
                HideEstimateEmbeddedPreview();
                HideEstimateEmbeddedSecondaryPreview();
            }
            else
            {
                ScheduleEstimateEmbeddedLayout();
                ScheduleEstimateEmbeddedLayoutSecondary();
            }

            if (!ReferenceEquals(item, PdfTab))
            {
                HidePdfEmbeddedPreview();
            }
            else
            {
                UpdatePdfSelectionInfo();
                SchedulePdfEmbeddedLayout();
            }

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

            if (view != null)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (view is IEditableCollectionView editableView && (editableView.IsAddingNew || editableView.IsEditingItem))
                        return;

                    view.Refresh();
                }), DispatcherPriority.Background);
            }
            if (item.Header?.ToString() == "Табель")
            {
                if (timesheetNeedsRebuild || TimesheetGrid?.ItemsSource == null)
                    RebuildTimesheetView(force: true);
            }
            if (item.Header?.ToString() == "ПР")
            {
                RefreshProductionJournalLookups();
                if (productionStateDirty || ProductionJournalGrid?.ItemsSource == null)
                    RefreshProductionJournalState();
            }
            if (item.Header?.ToString() == "Осмотры")
            {
                RefreshInspectionLookups();
                if (inspectionStateDirty || InspectionJournalGrid?.ItemsSource == null)
                    RefreshInspectionJournalState();
            }
            if (ReferenceEquals(item, EstimateTab))
            {
                UpdateEstimateSelectionInfo();
                ScheduleEstimateEmbeddedLayout();
            }

            ApplyGridColumnPreferences(GetGridByTabHeader(item.Header?.ToString()));
            ScheduleAutoFitCurrentTabColumns(item);

            tabStopwatch.Stop();
            UpdateTabOpenDiagnostics(item.Header?.ToString() ?? string.Empty, tabStopwatch.Elapsed);
        }

        private void ScheduleAutoFitCurrentTabColumns(TabItem item)
        {
            if (item == null || currentObject?.UiSettings?.AutoFitCurrentTabColumns != true)
                return;

            var tabHeader = item.Header?.ToString();
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (!ReferenceEquals(MainTabs?.SelectedItem, item))
                    return;

                var grid = GetGridByTabHeader(tabHeader);
                if (grid == null || grid.Columns.Count == 0)
                    return;

                DataGridSizingHelper.SetEnableSmartSizing(grid, false);
                DataGridSizingHelper.SetEnableSmartSizing(grid, true);
            }), DispatcherPriority.Background);
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
            if (currentObject == null)
                return;

            currentObject.OtJournal ??= new List<OtJournalEntry>();
            EnsureReferenceMappingsStorage();
        }

        private void EnsureReferenceMappingsStorage()
        {
            if (currentObject == null)
                return;

            currentObject.OtInstructionNumbersByProfession ??= new Dictionary<string, string>();
            currentObject.ProductionDeviationsByType ??= new Dictionary<string, List<string>>();

            var normalizedOt = new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            foreach (var pair in currentObject.OtInstructionNumbersByProfession)
            {
                var profession = pair.Key?.Trim();
                var numbers = pair.Value?.Trim();
                if (string.IsNullOrWhiteSpace(profession) || string.IsNullOrWhiteSpace(numbers))
                    continue;

                normalizedOt[profession] = numbers;
            }

            var normalizedProduction = new Dictionary<string, List<string>>(StringComparer.CurrentCultureIgnoreCase);
            foreach (var pair in currentObject.ProductionDeviationsByType)
            {
                var materialType = pair.Key?.Trim();
                if (string.IsNullOrWhiteSpace(materialType))
                    continue;

                var deviations = (pair.Value ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                if (deviations.Count == 0)
                    continue;

                normalizedProduction[materialType] = deviations;
            }

            currentObject.OtInstructionNumbersByProfession = normalizedOt.ToDictionary(x => x.Key, x => x.Value);
            currentObject.ProductionDeviationsByType = normalizedProduction.ToDictionary(x => x.Key, x => x.Value);
        }

        private void BindOtJournal()
        {
            OtJournalGrid.ItemsSource = currentObject?.OtJournal;
            if (currentObject?.OtJournal == null)
            {
                OtJournalGrid.ItemsSource = null;
                return;
            }

            EnsureProjectUiSettings();
            selectedOtStatusFilter = string.IsNullOrWhiteSpace(currentObject.UiSettings?.OtStatusFilter) ? "Все" : currentObject.UiSettings.OtStatusFilter.Trim();
            selectedOtSpecialtyFilter = string.IsNullOrWhiteSpace(currentObject.UiSettings?.OtSpecialtyFilter) ? "Все" : currentObject.UiSettings.OtSpecialtyFilter.Trim();
            selectedOtBrigadeFilter = string.IsNullOrWhiteSpace(currentObject.UiSettings?.OtBrigadeFilter) ? "Все" : currentObject.UiSettings.OtBrigadeFilter.Trim();

            var view = CollectionViewSource.GetDefaultView(currentObject.OtJournal);
            view.Filter = OtJournalFilter;
            OtJournalGrid.ItemsSource = view;
            SubscribeOtJournalEntryEvents();
            RefreshBrigadierNames();
            RefreshSpecialties();
            RefreshProfessions();
            RefreshOtFilterOptions();
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

            if (!string.IsNullOrWhiteSpace(otSearchText)
                && !(row.FullName ?? string.Empty).Contains(otSearchText, StringComparison.CurrentCultureIgnoreCase))
                return false;

            if (!string.Equals(selectedOtStatusFilter, "Все", StringComparison.CurrentCultureIgnoreCase))
            {
                var status = GetOtStatusFilterLabel(row);
                if (!string.Equals(status, selectedOtStatusFilter, StringComparison.CurrentCultureIgnoreCase))
                    return false;
            }

            if (!string.Equals(selectedOtSpecialtyFilter, "Все", StringComparison.CurrentCultureIgnoreCase))
            {
                var specialty = string.IsNullOrWhiteSpace(row.Specialty) ? "Без специальности" : row.Specialty.Trim();
                if (!string.Equals(specialty, selectedOtSpecialtyFilter, StringComparison.CurrentCultureIgnoreCase))
                    return false;
            }

            if (!string.Equals(selectedOtBrigadeFilter, "Все", StringComparison.CurrentCultureIgnoreCase))
            {
                var brigade = GetOtBrigadeFilterLabel(row);
                if (!string.Equals(brigade, selectedOtBrigadeFilter, StringComparison.CurrentCultureIgnoreCase))
                    return false;
            }

            return true;
        }

        private string GetOtStatusFilterLabel(OtJournalEntry row)
        {
            if (row == null)
                return "Без статуса";
            if (row.IsDismissed)
                return "Снят с объекта";
            if (row.IsPrimaryInstruction && row.IsPendingRepeat)
                return "Требуется первичный";
            if (row.IsPendingRepeat)
                return "Требуется повторный";
            if (row.IsPrimaryInstruction && row.IsRepeatCompleted)
                return "Первичный пройден";
            if (row.IsScheduledRepeat)
                return "Запланирован";
            if (row.IsRepeatCompleted)
                return "Повторный пройден";
            return "Без статуса";
        }

        private string GetOtBrigadeFilterLabel(OtJournalEntry row)
        {
            if (row == null)
                return "Без бригады";
            if (row.IsBrigadier)
                return string.IsNullOrWhiteSpace(row.FullName) ? "Без бригады" : row.FullName.Trim();
            return string.IsNullOrWhiteSpace(row.BrigadierName) ? "Без бригады" : row.BrigadierName.Trim();
        }

        private void RefreshOtFilterOptions()
        {
            suppressOtFilterSelectionChange = true;
            try
            {
                otStatusFilters.Clear();
                foreach (var status in new[] { "Все", "Требуется первичный", "Требуется повторный", "Первичный пройден", "Запланирован", "Повторный пройден", "Снят с объекта", "Без статуса" })
                    otStatusFilters.Add(status);

                otSpecialtyFilters.Clear();
                otSpecialtyFilters.Add("Все");
                if (currentObject?.OtJournal != null)
                {
                    foreach (var specialty in currentObject.OtJournal
                                 .Select(x => string.IsNullOrWhiteSpace(x.Specialty) ? "Без специальности" : x.Specialty.Trim())
                                 .Distinct(StringComparer.CurrentCultureIgnoreCase)
                                 .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
                    {
                        otSpecialtyFilters.Add(specialty);
                    }
                }

                otBrigadeFilters.Clear();
                otBrigadeFilters.Add("Все");
                if (currentObject?.OtJournal != null)
                {
                    foreach (var brigade in currentObject.OtJournal
                                 .Select(GetOtBrigadeFilterLabel)
                                 .Distinct(StringComparer.CurrentCultureIgnoreCase)
                                 .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
                    {
                        otBrigadeFilters.Add(brigade);
                    }
                }

                if (OtStatusFilterBox != null)
                {
                    OtStatusFilterBox.ItemsSource = otStatusFilters;
                    OtStatusFilterBox.SelectedItem = otStatusFilters.Contains(selectedOtStatusFilter) ? selectedOtStatusFilter : "Все";
                }

                if (OtSpecialtyFilterBox != null)
                {
                    OtSpecialtyFilterBox.ItemsSource = otSpecialtyFilters;
                    OtSpecialtyFilterBox.SelectedItem = otSpecialtyFilters.Contains(selectedOtSpecialtyFilter) ? selectedOtSpecialtyFilter : "Все";
                }

                if (OtBrigadeFilterBox != null)
                {
                    OtBrigadeFilterBox.ItemsSource = otBrigadeFilters;
                    OtBrigadeFilterBox.SelectedItem = otBrigadeFilters.Contains(selectedOtBrigadeFilter) ? selectedOtBrigadeFilter : "Все";
                }
            }
            finally
            {
                suppressOtFilterSelectionChange = false;
            }
        }

        private void OtFilters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (suppressOtFilterSelectionChange)
                return;

            selectedOtStatusFilter = OtStatusFilterBox?.SelectedItem?.ToString() ?? "Все";
            selectedOtSpecialtyFilter = OtSpecialtyFilterBox?.SelectedItem?.ToString() ?? "Все";
            selectedOtBrigadeFilter = OtBrigadeFilterBox?.SelectedItem?.ToString() ?? "Все";

            EnsureProjectUiSettings();
            if (currentObject?.UiSettings != null)
            {
                currentObject.UiSettings.OtStatusFilter = selectedOtStatusFilter;
                currentObject.UiSettings.OtSpecialtyFilter = selectedOtSpecialtyFilter;
                currentObject.UiSettings.OtBrigadeFilter = selectedOtBrigadeFilter;
            }

            RequestOtSearchRefresh(immediate: true);
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

            if (e.PropertyName == nameof(OtJournalEntry.Specialty)
                || e.PropertyName == nameof(OtJournalEntry.BrigadierName)
                || e.PropertyName == nameof(OtJournalEntry.IsBrigadier)
                || e.PropertyName == nameof(OtJournalEntry.IsDismissed)
                || e.PropertyName == nameof(OtJournalEntry.IsPendingRepeat)
                || e.PropertyName == nameof(OtJournalEntry.IsScheduledRepeat)
                || e.PropertyName == nameof(OtJournalEntry.IsRepeatCompleted))
            {
                RefreshOtFilterOptions();
            }

            if (!isSyncingTimesheetToOt && IsOtPropertyAffectingTimesheet(e.PropertyName))
            {
                SyncOtEntryToTimesheet(row, refreshVisibleTimesheet: true);
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

            var values = Enumerable.Empty<string>();

            if (currentObject?.OtJournal != null)
            {
                values = values.Concat(currentObject.OtJournal
                    .Where(x => !string.IsNullOrWhiteSpace(x.Specialty))
                    .Select(x => x.Specialty.Trim()));
            }

            if (currentObject?.TimesheetPeople != null)
            {
                values = values.Concat(currentObject.TimesheetPeople
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.Specialty))
                    .Select(x => x.Specialty.Trim()));
            }

            foreach (var item in values
                         .Where(x => !string.IsNullOrWhiteSpace(x))
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

            EnsureReferenceMappingsStorage();
            var mappedNumbers = TryGetMappedInstructionNumbers(key);
            if (!string.IsNullOrWhiteSpace(mappedNumbers))
            {
                row.InstructionNumbers = mappedNumbers;
                return;
            }

            var template = currentObject.OtJournal
                .Where(x => !ReferenceEquals(x, row))
                .FirstOrDefault(x =>
                    !string.IsNullOrWhiteSpace(x.InstructionNumbers)
                    && (string.Equals(x.Profession?.Trim(), key, StringComparison.CurrentCultureIgnoreCase)
                        || string.Equals(x.Specialty?.Trim(), key, StringComparison.CurrentCultureIgnoreCase)));

            if (template != null)
                row.InstructionNumbers = template.InstructionNumbers;
        }

        private string TryGetMappedInstructionNumbers(string professionOrSpecialty)
        {
            if (currentObject?.OtInstructionNumbersByProfession == null || string.IsNullOrWhiteSpace(professionOrSpecialty))
                return string.Empty;

            var key = professionOrSpecialty.Trim();
            if (currentObject.OtInstructionNumbersByProfession.TryGetValue(key, out var exactValue)
                && !string.IsNullOrWhiteSpace(exactValue))
            {
                return exactValue.Trim();
            }

            var pair = currentObject.OtInstructionNumbersByProfession.FirstOrDefault(x =>
                string.Equals(x.Key?.Trim(), key, StringComparison.CurrentCultureIgnoreCase)
                && !string.IsNullOrWhiteSpace(x.Value));

            return pair.Equals(default(KeyValuePair<string, string>))
                ? string.Empty
                : pair.Value.Trim();
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

            var toAdd = OtRepeatRuleEngine.Apply(currentObject.OtJournal, DateTime.Today, BuildRepeatInstructionType);
            foreach (var row in toAdd)
                row.PropertyChanged += OtJournalEntry_PropertyChanged;

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

        private static string NormalizeReminderPresentationMode(string mode)
        {
            var normalized = (mode ?? string.Empty).Trim().ToLowerInvariant();
            return normalized switch
            {
                ReminderPresentationModes.Tabs => ReminderPresentationModes.Tabs,
                ReminderPresentationModes.Combined => ReminderPresentationModes.Combined,
                _ => ReminderPresentationModes.Overlay
            };
        }

        private string GetReminderPresentationMode()
        {
            var mode = currentObject?.UiSettings?.ReminderPresentationMode;
            return NormalizeReminderPresentationMode(mode);
        }

        private bool ShouldShowOverlayReminders()
        {
            var mode = GetReminderPresentationMode();
            return mode == ReminderPresentationModes.Overlay || mode == ReminderPresentationModes.Combined;
        }

        private bool ShouldHighlightReminderTabs()
        {
            var mode = GetReminderPresentationMode();
            return mode == ReminderPresentationModes.Tabs || mode == ReminderPresentationModes.Combined;
        }

        private static string ParseReminderSectionTabHeader(string header)
        {
            if (string.IsNullOrWhiteSpace(header))
                return string.Empty;

            const string prefix = "Вкладка ";
            var text = header.Trim();
            if (text.StartsWith(prefix, StringComparison.CurrentCultureIgnoreCase))
                text = text.Substring(prefix.Length).Trim();

            return text;
        }

        private void ClearTabReminderMessages()
        {
            tabReminderMessages.Clear();
            UpdateTabButtons();
        }

        private void UpdateTabReminderMessagesFromSections(IEnumerable<ReminderSectionViewModel> sections)
        {
            tabReminderMessages.Clear();
            foreach (var section in sections ?? Enumerable.Empty<ReminderSectionViewModel>())
            {
                var tabHeader = ParseReminderSectionTabHeader(section.Header);
                if (string.IsNullOrWhiteSpace(tabHeader))
                    continue;

                var items = (section.Items ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .Take(5)
                    .ToList();

                if (items.Count == 0)
                    continue;

                tabReminderMessages[tabHeader] = items;
            }

            UpdateTabButtons();
        }

        private void UpdateOtReminders()
        {
            if (currentObject == null)
            {
                reminderSections.Clear();
                ClearTabReminderMessages();
                SetReminderPopupVisible(false);
                return;
            }

            EnsureProjectUiSettings();
            if (currentObject.UiSettings?.ShowReminderPopup == false)
            {
                reminderSections.Clear();
                ClearTabReminderMessages();
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
                            var instructionName = x.IsPrimaryInstruction ? "первичный" : "повторный";
                            return $"Нужен {instructionName} инструктаж: {person}";
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
                ClearTabReminderMessages();
                SetReminderPopupVisible(false);
                return;
            }

            reminderSections.Clear();
            foreach (var section in sections)
                reminderSections.Add(section);
            UpdateTabReminderMessagesFromSections(ShouldHighlightReminderTabs() ? sections : null);

            var overlay = EnsureReminderOverlayWindow();
            overlay.SnoozeButtonElement.Content = $"Отложить ({GetReminderSnoozeMinutes()} мин)";

            var isSnoozed = reminderSnoozedUntil.HasValue && reminderSnoozedUntil.Value > DateTime.Now;
            var detailsHidden = currentObject.UiSettings?.HideReminderDetails == true;
            overlay.ToggleDetailsButtonElement.Content = detailsHidden ? "Показать" : "Скрыть";

            if (isSnoozed)
            {
                // По запросу: при отложении уведомления полностью скрываются до окончания срока.
                ClearTabReminderMessages();
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

            SetReminderPopupVisible(ShouldShowOverlayReminders());
        }

        private List<SummaryBalanceReminderItem> CollectSummaryBalanceReminderItems()
        {
            if (currentObject == null)
                return new List<SummaryBalanceReminderItem>();

            EnsureProjectUiSettings();
            var settings = currentObject.UiSettings ?? new ProjectUiSettings();
            if (!settings.SummaryReminderOnOverage && !settings.SummaryReminderOnDeficit)
                return new List<SummaryBalanceReminderItem>();

            return BuildSummaryBalanceItems(
                settings.SummaryReminderOnOverage,
                settings.SummaryReminderOnDeficit,
                settings.SummaryReminderOnlyMain);
        }

        private List<SummaryBalanceReminderItem> BuildSummaryBalanceItems(bool includeOverage, bool includeDeficit, bool onlyMainCategory)
        {
            var result = new List<SummaryBalanceReminderItem>();
            if (currentObject == null || (!includeOverage && !includeDeficit))
                return result;

            var keyComparer = StringComparer.CurrentCultureIgnoreCase;
            var candidates = new HashSet<(string Group, string Material)>();

            if (currentObject.Demand != null)
            {
                foreach (var key in currentObject.Demand.Keys.Where(x => !string.IsNullOrWhiteSpace(x)))
                {
                    var parts = key.Split(new[] { "::" }, 2, StringSplitOptions.None);
                    if (parts.Length == 2
                        && !string.IsNullOrWhiteSpace(parts[0])
                        && !string.IsNullOrWhiteSpace(parts[1]))
                    {
                        candidates.Add((parts[0].Trim(), parts[1].Trim()));
                    }
                }
            }

            foreach (var row in journal)
            {
                var group = row.MaterialGroup?.Trim();
                var material = row.MaterialName?.Trim();
                if (!string.IsNullOrWhiteSpace(group) && !string.IsNullOrWhiteSpace(material))
                    candidates.Add((group, material));
            }

            if (currentObject.MaterialCatalog != null)
            {
                foreach (var item in currentObject.MaterialCatalog)
                {
                    var group = item.TypeName?.Trim();
                    var material = item.MaterialName?.Trim();
                    if (!string.IsNullOrWhiteSpace(group) && !string.IsNullOrWhiteSpace(material))
                        candidates.Add((group, material));
                }
            }

            foreach (var (group, material) in candidates.OrderBy(x => x.Group, keyComparer).ThenBy(x => x.Material, keyComparer))
            {
                var category = ResolveMaterialCategory(group, material);
                if (onlyMainCategory && !string.Equals(category, "Основные", StringComparison.CurrentCultureIgnoreCase))
                    continue;

                var records = journal
                    .Where(x => string.Equals((x.MaterialGroup ?? string.Empty).Trim(), group, StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase))
                    .ToList();

                if (onlyMainCategory)
                    records = records.Where(x => string.Equals(x.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)).ToList();

                var unit = records.Select(x => x.Unit).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? GetUnitForMaterial(group, material);
                var totalArrived = NormalizeQuantityByUnit(records.Sum(x => x.Quantity), unit);
                var demand = GetOrCreateDemand(BuildDemandKey(group, material), unit);
                var totalNeed = NormalizeQuantityByUnit(
                    BuildSummaryBlocks(group)
                        .SelectMany(x => x.Levels.Select(level => GetDemandValue(demand, x.Block, level)))
                        .Sum(),
                    unit);

                var delta = totalArrived - totalNeed;
                if (!CalculationCore.HasDifference(totalArrived, totalNeed))
                    continue;

                if (CalculationCore.ShouldNotifySummaryDelta(totalArrived, totalNeed, includeOverage, includeDeficit))
                {
                    result.Add(new SummaryBalanceReminderItem
                    {
                        Category = category,
                        Group = group,
                        Material = material,
                        Unit = unit,
                        Quantity = Math.Abs(delta),
                        IsOverage = CalculationCore.IsOverage(totalArrived, totalNeed)
                    });
                }
            }

            return result;
        }

        private string ResolveMaterialCategory(string group, string material)
        {
            if (currentObject?.MaterialCatalog != null)
            {
                var catalogCategory = currentObject.MaterialCatalog
                    .Where(x => string.Equals((x.TypeName ?? string.Empty).Trim(), group, StringComparison.CurrentCultureIgnoreCase)
                             && string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.CategoryName?.Trim())
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
                if (!string.IsNullOrWhiteSpace(catalogCategory))
                    return catalogCategory;
            }

            var rowCategory = journal
                .Where(x => string.Equals((x.MaterialGroup ?? string.Empty).Trim(), group, StringComparison.CurrentCultureIgnoreCase)
                         && string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase))
                .Select(x => x.Category?.Trim())
                .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
            return string.IsNullOrWhiteSpace(rowCategory) ? "Основные" : rowCategory;
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
            RefreshOtFilterOptions();
            RequestReminderRefresh();
            SaveState();
        }
        private void MarkRepeatDone(OtJournalEntry row)
        {
            if (row == null)
                return;

            if (!row.IsActionEnabled)
            {
                MessageBox.Show("Для выбранной записи действие недоступно.");
                return;
            }

            var isPrimaryInstruction = row.IsPrimaryInstruction;
            row.InstructionDate = DateTime.Today;
            row.InstructionType = isPrimaryInstruction
                ? "Первичный на рабочем месте"
                : BuildRepeatInstructionType(GetRepeatIndexForRow(row));
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

            PushUndo();
            currentObject.OtJournal.Remove(row);
            SortOtJournal();
            RefreshBrigadierNames();
            RefreshSpecialties();
            RefreshProfessions();
            RefreshOtFilterOptions();
            RequestReminderRefresh();
            MarkTimesheetOtSyncDirty();
            RequestTimesheetRebuild();
            SaveState();
        }
        private void OtSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            otSearchText = OtSearchTextBox.Text?.Trim() ?? string.Empty;
            RequestOtSearchRefresh();
        }

        private void OtJournalGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit)
                return;

            var changedRow = e.Row.Item as OtJournalEntry;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                RefreshBrigadierNames();
                RefreshSpecialties();
                RefreshProfessions();
                RefreshOtFilterOptions();
                if (changedRow != null)
                    SyncOtEntryToTimesheet(changedRow, refreshVisibleTimesheet: true);
                else
                {
                    MarkTimesheetOtSyncDirty();
                    RequestTimesheetRebuild();
                }
                RequestReminderRefresh();
            }));
        }

        private void OtJournalGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            var changedRow = e.Row.Item as OtJournalEntry;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                RefreshBrigadierNames();
                RefreshSpecialties();
                RefreshProfessions();
                RefreshOtFilterOptions();
                NormalizeOtRows();
                SortOtJournal();
                RequestReminderRefresh();
                if (changedRow != null)
                    SyncOtEntryToTimesheet(changedRow, refreshVisibleTimesheet: true);
                else
                {
                    MarkTimesheetOtSyncDirty();
                    RequestTimesheetRebuild();
                }
                SaveState();
            }));
        }

        private void OtJournalGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (ShouldHandleGridRowDeleteShortcut(e) && OtJournalGrid.SelectedItem is OtJournalEntry)
            {
                DeleteSelectedOtRow_Click(this, new RoutedEventArgs());
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
            timesheetInitialized = true;
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
                person.ArchivedMonths ??= new List<TimesheetMonthEntry>();
                person.Months.RemoveAll(m => m == null || string.IsNullOrWhiteSpace(m.MonthKey));

                var outOfWindow = person.Months
                    .Where(m => !allowedKeys.Contains(m.MonthKey))
                    .ToList();
                foreach (var archiveMonth in outOfWindow)
                {
                    person.Months.Remove(archiveMonth);
                    var existingArchive = person.ArchivedMonths.FirstOrDefault(x => string.Equals(x.MonthKey, archiveMonth.MonthKey, StringComparison.Ordinal));
                    if (existingArchive != null)
                        person.ArchivedMonths.Remove(existingArchive);
                    person.ArchivedMonths.Add(archiveMonth);
                }

                person.ArchivedMonths.RemoveAll(m => m == null || string.IsNullOrWhiteSpace(m.MonthKey));

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

                foreach (var key in allowedKeys.Where(x => !person.Months.Any(m => string.Equals(m.MonthKey, x, StringComparison.Ordinal))))
                {
                    var archived = person.ArchivedMonths.FirstOrDefault(m => string.Equals(m.MonthKey, key, StringComparison.Ordinal));
                    if (archived != null)
                    {
                        person.Months.Add(archived);
                        person.ArchivedMonths.Remove(archived);
                    }
                    else
                    {
                        person.Months.Add(new TimesheetMonthEntry
                        {
                            MonthKey = key,
                            DayEntries = new Dictionary<int, TimesheetDayEntry>(),
                            DayValues = new Dictionary<int, string>()
                        });
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

        private void SyncOtEntryToTimesheet(OtJournalEntry sourceRow, bool refreshVisibleTimesheet)
        {
            if (sourceRow == null || currentObject?.OtJournal == null)
                return;

            EnsureTimesheetStorage();

            var sourceNameKey = NormalizePersonNameKey(sourceRow.FullName);
            if (sourceRow.PersonId == Guid.Empty && string.IsNullOrWhiteSpace(sourceNameKey))
                return;

            var relatedRows = currentObject.OtJournal
                .Where(x => x != null
                    && ((sourceRow.PersonId != Guid.Empty && x.PersonId == sourceRow.PersonId)
                        || (!string.IsNullOrWhiteSpace(sourceNameKey)
                            && string.Equals(NormalizePersonNameKey(x.FullName), sourceNameKey, StringComparison.CurrentCultureIgnoreCase))))
                .ToList();

            if (relatedRows.Count == 0)
                relatedRows.Add(sourceRow);

            var latestActive = relatedRows
                .Where(x => !x.IsDismissed && !string.IsNullOrWhiteSpace(x.FullName))
                .OrderByDescending(x => x.InstructionDate)
                .FirstOrDefault();

            if (latestActive == null)
            {
                timesheetNeedsRebuild = true;
                return;
            }

            var resolvedId = relatedRows
                .Select(x => x.PersonId)
                .FirstOrDefault(x => x != Guid.Empty);

            if (resolvedId == Guid.Empty)
            {
                var latestNameKey = NormalizePersonNameKey(latestActive.FullName);
                var byName = currentObject.TimesheetPeople
                    .FirstOrDefault(x => string.Equals(NormalizePersonNameKey(x.FullName), latestNameKey, StringComparison.CurrentCultureIgnoreCase));
                resolvedId = byName?.PersonId ?? Guid.NewGuid();
            }

            foreach (var related in relatedRows.Where(x => x.PersonId != resolvedId))
                related.PersonId = resolvedId;

            latestActive.PersonId = resolvedId;

            var person = currentObject.TimesheetPeople.FirstOrDefault(x => x.PersonId == resolvedId);
            if (person == null)
            {
                var latestNameKey = NormalizePersonNameKey(latestActive.FullName);
                person = currentObject.TimesheetPeople
                    .FirstOrDefault(x => string.Equals(NormalizePersonNameKey(x.FullName), latestNameKey, StringComparison.CurrentCultureIgnoreCase));
            }

            if (person == null)
            {
                person = new TimesheetPersonEntry
                {
                    PersonId = resolvedId
                };
                currentObject.TimesheetPeople.Add(person);
                RefreshTimesheetPersonSubscriptions();
            }

            person.FullName = latestActive.FullName?.Trim();
            person.Specialty = latestActive.Specialty;
            person.Rank = latestActive.Rank;
            person.IsBrigadier = latestActive.IsBrigadier;
            person.BrigadeName = latestActive.IsBrigadier ? latestActive.FullName?.Trim() : latestActive.BrigadierName;

            if (refreshVisibleTimesheet
                && timesheetInitialized
                && ReferenceEquals(MainTabs?.SelectedItem, TimesheetTab))
            {
                RefreshTimesheetBrigades();
                RefreshTimesheetRows();
                TimesheetGrid?.Items.Refresh();
                UpdateTimesheetMissingDocsNotification();
                timesheetNeedsRebuild = false;
            }
            else
            {
                timesheetNeedsRebuild = true;
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
            var specialtyCellFactory = new FrameworkElementFactory(typeof(TextBlock));
            specialtyCellFactory.SetBinding(TextBlock.TextProperty, new Binding(nameof(TimesheetRowViewModel.Specialty)));
            var specialtyCellTemplate = new DataTemplate { VisualTree = specialtyCellFactory };

            var specialtyEditFactory = new FrameworkElementFactory(typeof(ComboBox));
            specialtyEditFactory.SetValue(ComboBox.IsEditableProperty, true);
            specialtyEditFactory.SetValue(ComboBox.IsTextSearchEnabledProperty, true);
            specialtyEditFactory.SetValue(ComboBox.StaysOpenOnEditProperty, true);
            specialtyEditFactory.SetValue(ComboBox.ItemsSourceProperty, specialties);
            specialtyEditFactory.SetBinding(ComboBox.TextProperty, new Binding(nameof(TimesheetRowViewModel.Specialty))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });
            var specialtyEditTemplate = new DataTemplate { VisualTree = specialtyEditFactory };

            TimesheetGrid.Columns.Add(new DataGridTemplateColumn
            {
                Header = "Специальность",
                Width = 170,
                CellTemplate = specialtyCellTemplate,
                CellEditingTemplate = specialtyEditTemplate
            });
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
                    style.Setters.Add(new Setter(DataGridCell.BackgroundProperty, new SolidColorBrush(Color.FromRgb(230, 238, 252))));
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

            ApplyGridColumnPreferences(TimesheetGrid);
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
            TimesheetRowViewModel affectedRow = null;
            if (sender is FrameworkElement element
                && element.DataContext is TimesheetRowViewModel row
                && int.TryParse(element.Tag?.ToString(), out var day)
                && day > 0)
            {
                affectedRow = row;
                var shouldMarkPresent = sender is ToggleButton toggle && toggle.IsChecked == true;
                row.SetPresenceChecked(day, shouldMarkPresent);
                var autoValue = shouldMarkPresent
                    ? Math.Clamp(row.DailyWorkHours, 1, 24).ToString(CultureInfo.CurrentCulture)
                    : "Н";

                row[day] = autoValue;
                row.RecalculateTotal();
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (affectedRow?.Source != null)
                    UpdateRepeatRequirementByTimesheet(affectedRow.Source);
                SaveState();
                UpdateTimesheetMissingDocsNotification();
                RequestReminderRefresh();
            }), DispatcherPriority.Background);
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
            editor.SetBinding(TextBox.ForegroundProperty, new Binding { Converter = new TimesheetDayForegroundConverter(day) });
            editor.SetBinding(TextBox.FontWeightProperty, new Binding { Converter = new TimesheetDayFontWeightConverter(day) });
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

        private sealed class TimesheetDayForegroundConverter : IValueConverter
        {
            private readonly int day;
            public TimesheetDayForegroundConverter(int day) => this.day = day;

            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (value is not TimesheetRowViewModel row)
                    return Brushes.Black;

                return row.IsNonHourCode(day)
                    ? new SolidColorBrush(Color.FromRgb(127, 29, 29))
                    : Brushes.Black;
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
                => Binding.DoNothing;
        }

        private sealed class TimesheetDayFontWeightConverter : IValueConverter
        {
            private readonly int day;
            public TimesheetDayFontWeightConverter(int day) => this.day = day;

            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (value is not TimesheetRowViewModel row)
                    return FontWeights.Normal;

                return row.IsNonHourCode(day) ? FontWeights.SemiBold : FontWeights.Normal;
            }

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
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    SaveState();
                    UpdateTimesheetMissingDocsNotification();
                }), DispatcherPriority.Background);
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
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SaveState();
                UpdateTimesheetMissingDocsNotification();
            }), DispatcherPriority.Background);

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
                {
                    var hasPendingPrimary = currentObject.OtJournal.Any(x =>
                        x.PersonId == row.PersonId
                        && x.IsPendingRepeat
                        && x.IsPrimaryInstruction
                        && !x.IsDismissed);
                    var instructionName = hasPendingPrimary ? "первичный" : "повторный";
                    MessageBox.Show($"⚠ {row.FullName}: требуется срочный {instructionName} инструктаж по ОТ.", "Напоминание по ОТ", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
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
            if (ShouldHandleGridRowDeleteShortcut(e) && selectedTimesheetRow != null)
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

            PushUndo();
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

        private void TimesheetRangeOperation_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject?.TimesheetPeople == null || currentObject.TimesheetPeople.Count == 0)
            {
                MessageBox.Show("Табель пуст.");
                return;
            }

            var dialog = new Window
            {
                Title = "Операция по диапазону дат",
                Owner = this,
                Width = 520,
                Height = 290,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            dialog.Content = root;

            var fromPicker = new DatePicker { SelectedDate = timesheetMonth };
            var toPicker = new DatePicker { SelectedDate = timesheetMonth.AddDays(Math.Max(0, DateTime.DaysInMonth(timesheetMonth.Year, timesheetMonth.Month) - 1)) };
            var operationBox = new ComboBox
            {
                SelectedIndex = 0,
                ItemsSource = new[] { "Проставить часы по режиму", "Проставить Н", "Очистить" }
            };

            var line1 = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            line1.Children.Add(new TextBlock { Text = "С:", Width = 24, VerticalAlignment = VerticalAlignment.Center });
            line1.Children.Add(fromPicker);
            line1.Children.Add(new TextBlock { Text = "  По:", Margin = new Thickness(12, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center });
            line1.Children.Add(toPicker);
            Grid.SetRow(line1, 0);
            root.Children.Add(line1);

            var line2 = new StackPanel { Orientation = Orientation.Horizontal };
            line2.Children.Add(new TextBlock { Text = "Операция:", Width = 80, VerticalAlignment = VerticalAlignment.Center });
            line2.Children.Add(operationBox);
            Grid.SetRow(line2, 1);
            root.Children.Add(line2);

            var hint = new TextBlock
            {
                Text = "Если в табеле выбраны строки, операция применяется только к ним. Иначе ко всем строкам текущего фильтра.",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 10, 0, 0),
                Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139))
            };
            Grid.SetRow(hint, 2);
            root.Children.Add(hint);

            var footer = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var applyButton = new Button { Content = "Применить", MinWidth = 120 };
            var cancelButton = new Button { Content = "Отмена", MinWidth = 110, Margin = new Thickness(8, 0, 0, 0), IsCancel = true };
            footer.Children.Add(applyButton);
            footer.Children.Add(cancelButton);
            Grid.SetRow(footer, 4);
            root.Children.Add(footer);

            applyButton.Click += (_, _) =>
            {
                if (!fromPicker.SelectedDate.HasValue || !toPicker.SelectedDate.HasValue)
                {
                    MessageBox.Show("Укажите даты начала и конца.");
                    return;
                }

                var start = fromPicker.SelectedDate.Value.Date;
                var end = toPicker.SelectedDate.Value.Date;
                if (end < start)
                    (start, end) = (end, start);

                var targets = TimesheetGrid?.SelectedItems?.OfType<TimesheetRowViewModel>().ToList() ?? new List<TimesheetRowViewModel>();
                if (targets.Count == 0)
                    targets = timesheetRows.ToList();

                var mode = operationBox.SelectedItem?.ToString() ?? "Проставить часы по режиму";
                foreach (var date in Enumerable.Range(0, (end - start).Days + 1).Select(offset => start.AddDays(offset)))
                {
                    var monthKey = date.ToString("yyyy-MM", CultureInfo.InvariantCulture);
                    var day = date.Day;

                    foreach (var row in targets)
                    {
                        if (row?.Source == null)
                            continue;

                        if (string.Equals(mode, "Проставить Н", StringComparison.CurrentCulture))
                            row.Source.SetDayValue(monthKey, day, "Н");
                        else if (string.Equals(mode, "Очистить", StringComparison.CurrentCulture))
                            row.Source.SetDayValue(monthKey, day, string.Empty);
                        else
                            row.Source.SetDayValue(monthKey, day, Math.Clamp(row.DailyWorkHours, 1, 24).ToString(CultureInfo.CurrentCulture));
                    }
                }

                dialog.DialogResult = true;
                dialog.Close();
            };

            if (dialog.ShowDialog() != true)
                return;

            NormalizeTimesheetMonthsWindow();
            RefreshTimesheetRows();
            TimesheetGrid?.Items.Refresh();
            SaveState();
            UpdateTimesheetMissingDocsNotification();
        }

        private void OpenTodayActionsReport_Click(object sender, RoutedEventArgs e)
        {
            var items = new List<string>();
            if (currentObject?.OtJournal != null)
            {
                items.AddRange(currentObject.OtJournal
                    .Where(x => x.IsPendingRepeat && !x.IsDismissed)
                    .OrderBy(x => x.LastName, StringComparer.CurrentCultureIgnoreCase)
                    .Select(x => $"ОТ: нужен {(x.IsPrimaryInstruction ? "первичный" : "повторный")} инструктаж — {x.FullName}"));
            }

            var missingDocs = CollectTimesheetMissingDocsPreview(out var missingDocsCount);
            if (missingDocsCount > 0)
                items.AddRange(missingDocs.Select(x => $"Табель: нет документа — {x}"));

            if (items.Count == 0)
            {
                MessageBox.Show("На сегодня действий по ОТ и табелю нет.");
                return;
            }

            var dialog = new Window
            {
                Title = "Кто требует действие сегодня",
                Owner = this,
                Width = 760,
                Height = 520,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };
            var list = new ListBox
            {
                Margin = new Thickness(12),
                ItemsSource = items
            };
            dialog.Content = list;
            dialog.ShowDialog();
        }

        private void OpenOtPersonHistory_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject?.OtJournal == null || currentObject.OtJournal.Count == 0)
            {
                MessageBox.Show("Журнал ОТ пуст.");
                return;
            }

            var source = OtJournalGrid?.SelectedItem as OtJournalEntry;
            if (source == null)
            {
                MessageBox.Show("Выберите сотрудника в таблице ОТ.");
                return;
            }

            var sourceNameKey = NormalizePersonNameKey(source.FullName);
            var rows = currentObject.OtJournal
                .Where(x => x != null
                    && ((source.PersonId != Guid.Empty && x.PersonId == source.PersonId)
                        || (!string.IsNullOrWhiteSpace(sourceNameKey)
                            && string.Equals(NormalizePersonNameKey(x.FullName), sourceNameKey, StringComparison.CurrentCultureIgnoreCase))))
                .OrderByDescending(x => x.InstructionDate)
                .ToList();

            if (rows.Count == 0)
            {
                MessageBox.Show("Для выбранного сотрудника история не найдена.");
                return;
            }

            var dialog = new Window
            {
                Title = $"История инструктажей: {source.FullName}",
                Owner = this,
                Width = 1080,
                Height = 680,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            dialog.Content = root;

            root.Children.Add(new TextBlock
            {
                Text = "Полная история инструктажей по сотруднику. Цвета соответствуют статусам в журнале ОТ.",
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Border CreateLegendChip(string caption, string resourceKey)
            {
                var background = TryFindResource(resourceKey) as Brush ?? Brushes.White;
                return new Border
                {
                    Background = background,
                    BorderBrush = TryFindResource("StrokeBrush") as Brush ?? Brushes.LightGray,
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(8, 4, 8, 4),
                    Margin = new Thickness(0, 0, 8, 0),
                    Child = new TextBlock
                    {
                        Text = caption
                    }
                };
            }

            var legendPanel = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
            legendPanel.Children.Add(CreateLegendChip("Желтый: требуется инструктаж", "WarningSoftBrush"));
            legendPanel.Children.Add(CreateLegendChip("Зеленый: инструктаж пройден", "SuccessSoftBrush"));
            legendPanel.Children.Add(CreateLegendChip("Синий: запланирован", "SelectedBrush"));
            legendPanel.Children.Add(CreateLegendChip("Серый: снят с объекта", "SurfaceAltBrush"));
            Grid.SetRow(legendPanel, 1);
            root.Children.Add(legendPanel);

            var historyGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                ColumnWidth = DataGridLength.Auto,
                ItemsSource = rows,
                RowStyle = TryFindResource("OtDueRowStyle") as Style
            };

            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Дата", Binding = new Binding(nameof(OtJournalEntry.InstructionDate)) { StringFormat = "dd.MM.yyyy" }, Width = 110 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "ФИО", Binding = new Binding(nameof(OtJournalEntry.FullNameDisplay)), Width = 220 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Вид", Binding = new Binding(nameof(OtJournalEntry.InstructionType)), Width = 150 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Специальность", Binding = new Binding(nameof(OtJournalEntry.Specialty)), Width = 170 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Разряд", Binding = new Binding(nameof(OtJournalEntry.Rank)), Width = 60 });

            var instructionNumbersColumn = new DataGridTextColumn
            {
                Header = "Номера инструкций",
                Binding = new Binding(nameof(OtJournalEntry.InstructionNumbers)),
                Width = 240
            };
            instructionNumbersColumn.ElementStyle = new Style(typeof(TextBlock));
            instructionNumbersColumn.ElementStyle.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap));
            historyGrid.Columns.Add(instructionNumbersColumn);

            historyGrid.Columns.Add(new DataGridTextColumn { Header = "След. повторный", Binding = new Binding(nameof(OtJournalEntry.NextRepeatDate)) { StringFormat = "dd.MM.yyyy" }, Width = 120 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Статус", Binding = new Binding(nameof(OtJournalEntry.StatusLabel)), Width = 220 });

            DataGridSizingHelper.SetEnableSmartSizing(historyGrid, true);
            Grid.SetRow(historyGrid, 2);
            root.Children.Add(historyGrid);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            var closeButton = new Button
            {
                Content = "Закрыть",
                MinWidth = 120,
                IsCancel = true,
                Style = TryFindResource("SecondaryButton") as Style
            };
            footer.Children.Add(closeButton);
            Grid.SetRow(footer, 3);
            root.Children.Add(footer);

            dialog.ShowDialog();
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
            RefreshSpecialties();
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
                IsPendingRepeat = true,
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

            var gridBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225));
            var weekendBrush = new SolidColorBrush(Color.FromRgb(230, 238, 252));
            var nonHourBrush = new SolidColorBrush(Color.FromRgb(254, 226, 226));
            var acceptedDocBrush = new SolidColorBrush(Color.FromRgb(254, 240, 138));

            TableCell CreateCell(string text, bool center = true, Brush background = null, Brush foreground = null, bool bold = false)
            {
                var paragraph = new Paragraph(new Run(text ?? string.Empty))
                {
                    Margin = new Thickness(0),
                    TextAlignment = center ? TextAlignment.Center : TextAlignment.Left
                };

                if (foreground != null)
                    paragraph.Foreground = foreground;
                if (bold)
                    paragraph.FontWeight = FontWeights.SemiBold;

                var cell = new TableCell(paragraph)
                {
                    BorderBrush = gridBrush,
                    BorderThickness = new Thickness(0.5),
                    Padding = new Thickness(3, 2, 3, 2)
                };

                if (background != null)
                    cell.Background = background;

                return cell;
            }

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
            header.Cells.Add(CreateCell("№"));
            header.Cells.Add(CreateCell("ФИО"));
            header.Cells.Add(CreateCell("Спец."));
            header.Cells.Add(CreateCell("Раз."));
            for (var i = 1; i <= daysInMonth; i++)
            {
                var date = new DateTime(timesheetMonth.Year, timesheetMonth.Month, i);
                var isWeekend = date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday;
                header.Cells.Add(CreateCell(i.ToString(), background: isWeekend ? weekendBrush : null));
            }
            header.Cells.Add(CreateCell("Итого"));

            foreach (var row in timesheetRows)
            {
                var tr = new TableRow();
                if (row.IsBrigadier)
                    tr.FontWeight = FontWeights.Bold;
                group.Rows.Add(tr);

                tr.Cells.Add(CreateCell(row.Number.ToString()));
                tr.Cells.Add(CreateCell(row.FullName, center: false));
                tr.Cells.Add(CreateCell(row.Specialty, center: false));
                tr.Cells.Add(CreateCell(row.Rank));
                for (var i = 1; i <= daysInMonth; i++)
                {
                    var date = new DateTime(timesheetMonth.Year, timesheetMonth.Month, i);
                    var isWeekend = date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday;
                    var dayValue = blank ? string.Empty : row.GetDayValue(i);
                    Brush background = isWeekend ? weekendBrush : null;
                    Brush foreground = null;
                    var isNonHourCode = !blank && row.IsNonHourCode(i);
                    if (isNonHourCode)
                    {
                        background = row.IsDocumentAccepted(i) == true ? acceptedDocBrush : nonHourBrush;
                        foreground = new SolidColorBrush(Color.FromRgb(127, 29, 29));
                    }

                    tr.Cells.Add(CreateCell(dayValue, background: background, foreground: foreground, bold: isNonHourCode));
                }

                tr.Cells.Add(CreateCell(blank ? string.Empty : row.MonthTotalHours.ToString("0.##")));
            }

            return doc;
        }

        private void InitializeProductionJournal()
        {
            EnsureProductionJournalStorage();
            productionLookupsDirty = true;
            RefreshProductionJournalState();
            productionJournalInitialized = true;
        }

        private void EnsureProductionJournalStorage()
        {
            if (currentObject == null)
                return;

            currentObject.ProductionJournal ??= new List<ProductionJournalEntry>();
            currentObject.ProductionAutoFillSettings ??= new ProductionAutoFillSettings();
            currentObject.ProductionAutoFillProfiles ??= new List<ProductionAutoFillProfile>();
            if (string.IsNullOrWhiteSpace(currentObject.SelectedProductionAutoFillProfileName)
                && currentObject.ProductionAutoFillProfiles.Count > 0)
            {
                currentObject.SelectedProductionAutoFillProfileName = currentObject.ProductionAutoFillProfiles
                    .Select(x => x.Name?.Trim())
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;
            }
            currentObject.ProductionTemplates ??= new List<ProductionJournalTemplate>();
            currentObject.SummaryMarksByGroup ??= new Dictionary<string, List<string>>();
            EnsureReferenceMappingsStorage();
            RefreshProductionAutoFillProfileOptions();
        }

        private void RefreshProductionJournalState()
        {
            EnsureProductionJournalStorage();
            NormalizeProductionJournalRows();
            RefreshProductionBlockDisplayValues();
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
            productionStateDirty = false;
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

        private void RefreshProductionJournalLookups(bool force = false)
        {
            if (!force && !productionLookupsDirty)
                return;

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

            productionBlockOptions.Clear();
            if (currentObject != null)
            {
                for (var i = 1; i <= currentObject.BlocksCount; i++)
                    productionBlockOptions.Add(i.ToString());
            }

            RefreshProductionElementOptions();
            RefreshProductionMarkOptions();
            RefreshProductionDeviationOptions();
            RefreshProductionAutoFillProfileOptions();
            productionLookupsDirty = false;
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

        private void RefreshProductionAutoFillProfileOptions()
        {
            isRefreshingProductionProfileSelection = true;
            productionAutoFillProfileNames.Clear();
            if (currentObject?.ProductionAutoFillProfiles == null)
            {
                isRefreshingProductionProfileSelection = false;
                return;
            }

            foreach (var name in currentObject.ProductionAutoFillProfiles
                .Select(x => x.Name?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
            {
                productionAutoFillProfileNames.Add(name);
            }

            if (ProductionAutoFillProfileBox == null)
            {
                isRefreshingProductionProfileSelection = false;
                return;
            }

            var preferred = currentObject.SelectedProductionAutoFillProfileName?.Trim();
            if (!string.IsNullOrWhiteSpace(preferred)
                && productionAutoFillProfileNames.Any(x => string.Equals(x, preferred, StringComparison.CurrentCultureIgnoreCase)))
            {
                ProductionAutoFillProfileBox.SelectedItem = productionAutoFillProfileNames
                    .First(x => string.Equals(x, preferred, StringComparison.CurrentCultureIgnoreCase));
                isRefreshingProductionProfileSelection = false;
                return;
            }

            ProductionAutoFillProfileBox.SelectedItem = productionAutoFillProfileNames.FirstOrDefault();
            isRefreshingProductionProfileSelection = false;
        }

        private void ProductionAutoMasterButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button button || button.ContextMenu == null)
                return;

            button.ContextMenu.PlacementTarget = button;
            button.ContextMenu.Placement = PlacementMode.Bottom;
            button.ContextMenu.IsOpen = true;
        }

        private void ProductionAutoMasterMenu_Opened(object sender, RoutedEventArgs e)
        {
            if (sender is not ContextMenu menu)
                return;

            EnsureProductionJournalStorage();
            EnsureProductionAutoFillProfilesStorage();
            RefreshProductionAutoFillProfileOptions();

            var profilesRoot = menu.Items
                .OfType<MenuItem>()
                .FirstOrDefault(x => string.Equals(x.Tag?.ToString(), "AutoMasterProfilesRoot", StringComparison.CurrentCulture));

            if (profilesRoot == null)
                return;

            profilesRoot.Items.Clear();
            var selectedProfile = currentObject?.SelectedProductionAutoFillProfileName?.Trim() ?? string.Empty;
            var profileNames = productionAutoFillProfileNames.ToList();
            if (profileNames.Count == 0)
            {
                profilesRoot.Items.Add(new MenuItem
                {
                    Header = "Профили отсутствуют",
                    IsEnabled = false
                });
                return;
            }

            foreach (var profileName in profileNames)
            {
                var item = new MenuItem
                {
                    Header = profileName,
                    Tag = profileName,
                    IsCheckable = true,
                    IsChecked = string.Equals(profileName, selectedProfile, StringComparison.CurrentCultureIgnoreCase)
                };
                item.Click += ProductionAutoMasterProfileItem_Click;
                profilesRoot.Items.Add(item);
            }
        }

        private void ProductionAutoMasterProfileItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem item || currentObject == null)
                return;

            var profileName = item.Tag?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(profileName))
                return;

            if (string.Equals(currentObject.SelectedProductionAutoFillProfileName ?? string.Empty, profileName, StringComparison.CurrentCultureIgnoreCase))
                return;

            currentObject.SelectedProductionAutoFillProfileName = profileName;
            isRefreshingProductionProfileSelection = true;
            if (ProductionAutoFillProfileBox != null)
            {
                ProductionAutoFillProfileBox.SelectedItem = productionAutoFillProfileNames
                    .FirstOrDefault(x => string.Equals(x, profileName, StringComparison.CurrentCultureIgnoreCase));
            }
            isRefreshingProductionProfileSelection = false;
            SaveState();
        }

        private void ProductionAutoFillProfileBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (currentObject == null || isRefreshingProductionProfileSelection)
                return;

            var selected = ProductionAutoFillProfileBox?.SelectedItem?.ToString()?.Trim() ?? string.Empty;
            if (string.Equals(currentObject.SelectedProductionAutoFillProfileName ?? string.Empty, selected, StringComparison.CurrentCultureIgnoreCase))
                return;

            currentObject.SelectedProductionAutoFillProfileName = selected;
            SaveState();
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

        private void RefreshProductionMarkOptions()
        {
            var selectedWork = ProductionWorkBox?.Text?.Trim();
            var currentMarksText = ProductionMarksBox?.Text?.Trim() ?? string.Empty;
            var values = new List<string>();

            if (!string.IsNullOrWhiteSpace(selectedWork))
            {
                values.AddRange(LevelMarkHelper.GetMarksForGroup(currentObject, selectedWork));
                values.AddRange(currentObject?.ProductionJournal?
                    .Where(x => string.Equals((x.WorkName ?? string.Empty).Trim(), selectedWork, StringComparison.CurrentCultureIgnoreCase))
                    .SelectMany(x => LevelMarkHelper.ParseMarks(x.MarksText))
                    ?? Enumerable.Empty<string>());
            }
            else
            {
                values.AddRange((currentObject?.SummaryMarksByGroup?.Values
                        .Where(x => x != null)
                        .SelectMany(x => x ?? Enumerable.Empty<string>())
                        ?? Enumerable.Empty<string>())
                    .Concat(LevelMarkHelper.GetDefaultMarks(currentObject)));

                values.AddRange(currentObject?.ProductionJournal?
                    .SelectMany(x => LevelMarkHelper.ParseMarks(x.MarksText))
                    ?? Enumerable.Empty<string>());
            }

            FillLookupCollection(productionMarkOptions, values);

            if (ProductionMarksBox == null)
                return;

            if (string.IsNullOrWhiteSpace(currentMarksText))
            {
                ProductionMarksBox.Text = productionMarkOptions.FirstOrDefault() ?? string.Empty;
                return;
            }

            var normalized = NormalizeProductionMarksText(currentMarksText);
            var typedMarks = LevelMarkHelper.ParseMarks(normalized);
            if (typedMarks.Count == 0)
            {
                ProductionMarksBox.Text = productionMarkOptions.FirstOrDefault() ?? string.Empty;
                return;
            }

            var allowed = typedMarks
                .Where(mark => productionMarkOptions.Any(option =>
                    string.Equals(option, mark, StringComparison.CurrentCultureIgnoreCase)))
                .ToList();

            if (allowed.Count == typedMarks.Count)
            {
                ProductionMarksBox.Text = string.Join(", ", allowed);
                return;
            }

            ProductionMarksBox.Text = allowed.Count > 0
                ? string.Join(", ", allowed)
                : productionMarkOptions.FirstOrDefault() ?? string.Empty;
        }

        private List<string> GetDeviationOptionsForWork(string workName, bool includeCurrentJournal = true)
        {
            EnsureReferenceMappingsStorage();
            var options = new List<string>();
            var normalizedWork = workName?.Trim() ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(normalizedWork)
                && currentObject?.ProductionDeviationsByType != null
                && currentObject.ProductionDeviationsByType.TryGetValue(normalizedWork, out var mapped)
                && mapped != null)
            {
                options.AddRange(mapped);
            }
            else if (!string.IsNullOrWhiteSpace(normalizedWork) && currentObject?.ProductionDeviationsByType != null)
            {
                var mappedFallback = currentObject.ProductionDeviationsByType.FirstOrDefault(x =>
                    string.Equals((x.Key ?? string.Empty).Trim(), normalizedWork, StringComparison.CurrentCultureIgnoreCase));
                if (!mappedFallback.Equals(default(KeyValuePair<string, List<string>>)))
                    options.AddRange(mappedFallback.Value ?? new List<string>());
            }

            if (includeCurrentJournal)
            {
                if (!string.IsNullOrWhiteSpace(normalizedWork))
                {
                    options.AddRange(currentObject?.ProductionJournal?
                        .Where(x => string.Equals((x.WorkName ?? string.Empty).Trim(), normalizedWork, StringComparison.CurrentCultureIgnoreCase))
                        .Select(x => x.Deviations)
                        ?? Enumerable.Empty<string>());
                }
                else
                {
                    options.AddRange(currentObject?.ProductionJournal?.Select(x => x.Deviations)
                        ?? Enumerable.Empty<string>());
                }
            }

            return options
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private string ResolveAutoFillDeviation(string workName, string typedDeviation)
        {
            if (!string.IsNullOrWhiteSpace(typedDeviation))
                return typedDeviation.Trim();

            var options = GetDeviationOptionsForWork(workName, includeCurrentJournal: true);
            if (options.Count == 0)
                return string.Empty;

            var index = productionAutoRandom.Next(0, options.Count);
            return options[index];
        }

        private void RefreshProductionDeviationOptions()
        {
            var selectedWork = ProductionWorkBox?.Text?.Trim();
            var currentText = ProductionDeviationBox?.Text?.Trim() ?? string.Empty;
            List<string> values;

            if (!string.IsNullOrWhiteSpace(selectedWork))
            {
                values = GetDeviationOptionsForWork(selectedWork);
            }
            else
            {
                var allMapped = currentObject?.ProductionDeviationsByType?
                    .Where(x => !string.IsNullOrWhiteSpace(x.Key))
                    .SelectMany(x => x.Value ?? new List<string>())
                    ?? Enumerable.Empty<string>();
                values = allMapped
                    .Concat(currentObject?.ProductionJournal?.Select(x => x.Deviations) ?? Enumerable.Empty<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
            }

            FillLookupCollection(productionDeviationOptions, values);

            if (ProductionDeviationBox == null)
                return;

            if (!string.IsNullOrWhiteSpace(currentText))
            {
                ProductionDeviationBox.Text = currentText;
                return;
            }

            ProductionDeviationBox.Text = productionDeviationOptions.FirstOrDefault() ?? string.Empty;
        }

        private void ProductionWorkBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshProductionElementOptions();
            RefreshProductionMarkOptions();
            RefreshProductionDeviationOptions();
        }

        private void ProductionWorkBox_LostFocus(object sender, RoutedEventArgs e)
        {
            RefreshProductionElementOptions();
            RefreshProductionMarkOptions();
            RefreshProductionDeviationOptions();
        }

        private void InitializeInspectionJournal()
        {
            EnsureInspectionJournalStorage();
            inspectionLookupsDirty = true;
            RefreshInspectionJournalState();
            inspectionJournalInitialized = true;
        }

        private void EnsureInspectionJournalStorage()
        {
            if (currentObject == null)
                return;

            currentObject.InspectionJournal ??= new List<InspectionJournalEntry>();
            currentObject.InspectionTemplates ??= new List<InspectionJournalTemplate>();
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
            inspectionStateDirty = false;
        }

        private void RefreshInspectionLookups(bool force = false)
        {
            if (!force && !inspectionLookupsDirty)
                return;

            FillLookupCollection(inspectionJournalNames, currentObject?.InspectionJournal?.Select(x => x.JournalName));
            FillLookupCollection(inspectionNames, currentObject?.InspectionJournal?.Select(x => x.InspectionName));
            inspectionLookupsDirty = false;
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
            inspectionLookupsDirty = true;
            RefreshInspectionJournalState();
            if (InspectionJournalGrid != null && row != null && InspectionJournalGrid.Items.Contains(row))
            {
                InspectionJournalGrid.SelectedItem = row;
                InspectionJournalGrid.ScrollIntoView(row);
            }
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

            PushUndo();
            currentObject.InspectionJournal.Remove(row);
            selectedInspectionRow = null;
            inspectionLookupsDirty = true;
            RefreshInspectionJournalState();
            SaveState();
        }

        private void InspectionJournalGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!ShouldHandleGridRowDeleteShortcut(e))
                return;

            if (selectedInspectionRow == null && InspectionJournalGrid?.SelectedItem is not InspectionJournalEntry)
                return;

            DeleteInspectionRow_Click(this, new RoutedEventArgs());
            e.Handled = true;
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

        private void SaveInspectionTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            EnsureInspectionJournalStorage();
            var name = Microsoft.VisualBasic.Interaction.InputBox(
                "Введите название шаблона осмотра:",
                "Сохранить шаблон",
                (InspectionJournalNameBox?.Text ?? string.Empty).Trim())?.Trim();

            if (string.IsNullOrWhiteSpace(name))
                return;

            var template = new InspectionJournalTemplate
            {
                Name = name,
                JournalName = InspectionJournalNameBox?.Text?.Trim() ?? string.Empty,
                InspectionName = InspectionNameBox?.Text?.Trim() ?? string.Empty,
                ReminderPeriodDays = int.TryParse(InspectionPeriodDaysTextBox?.Text?.Trim(), out var days) && days > 0 ? days : 7,
                Notes = InspectionNotesTextBox?.Text?.Trim() ?? string.Empty
            };

            var existing = currentObject.InspectionTemplates
                .FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.CurrentCultureIgnoreCase));
            if (existing != null)
                currentObject.InspectionTemplates.Remove(existing);

            currentObject.InspectionTemplates.Add(template);
            SaveState();
        }

        private void ApplyInspectionTemplate_Click(object sender, RoutedEventArgs e)
        {
            var template = SelectInspectionTemplate();
            if (template == null)
                return;

            InspectionJournalNameBox.Text = template.JournalName ?? string.Empty;
            InspectionNameBox.Text = template.InspectionName ?? string.Empty;
            InspectionReminderStartDatePicker.SelectedDate = DateTime.Today;
            InspectionPeriodDaysTextBox.Text = Math.Max(1, template.ReminderPeriodDays).ToString(CultureInfo.InvariantCulture);
            InspectionLastCompletedDatePicker.SelectedDate = DateTime.Today;
            InspectionNotesTextBox.Text = template.Notes ?? string.Empty;
        }

        private void CreateInspectionFromTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            var template = SelectInspectionTemplate();
            if (template == null)
                return;

            var countText = Microsoft.VisualBasic.Interaction.InputBox(
                "Сколько записей добавить по шаблону?",
                "Добавить осмотры по шаблону",
                "1")?.Trim();
            if (string.IsNullOrWhiteSpace(countText))
                return;

            if (!int.TryParse(countText, out var count) || count <= 0)
                count = 1;

            count = Math.Clamp(count, 1, 52);
            var period = Math.Max(1, template.ReminderPeriodDays);
            var startDate = InspectionReminderStartDatePicker?.SelectedDate ?? DateTime.Today;
            var added = 0;

            for (var i = 0; i < count; i++)
            {
                var row = new InspectionJournalEntry
                {
                    JournalName = template.JournalName ?? string.Empty,
                    InspectionName = template.InspectionName ?? string.Empty,
                    ReminderStartDate = startDate.AddDays(i * period),
                    ReminderPeriodDays = period,
                    LastCompletedDate = null,
                    Notes = template.Notes ?? string.Empty
                };

                if (!ValidateInspectionRow(row, out _))
                    continue;

                currentObject.InspectionJournal.Add(row);
                added++;
            }

            if (added == 0)
            {
                MessageBox.Show("Не удалось добавить записи по шаблону.");
                return;
            }

            inspectionLookupsDirty = true;
            RefreshInspectionJournalState();
            SaveState();
            MessageBox.Show($"Добавлено записей осмотров: {added}.");
        }

        private InspectionJournalTemplate SelectInspectionTemplate()
        {
            if (currentObject?.InspectionTemplates == null || currentObject.InspectionTemplates.Count == 0)
            {
                MessageBox.Show("Сначала сохраните хотя бы один шаблон осмотра.");
                return null;
            }

            var selectedName = PromptSelectOption(
                "Выберите шаблон осмотра",
                "Шаблон",
                currentObject.InspectionTemplates.Select(x => x.Name));

            if (string.IsNullOrWhiteSpace(selectedName))
                return null;

            return currentObject.InspectionTemplates
                .FirstOrDefault(x => string.Equals(x.Name, selectedName, StringComparison.CurrentCultureIgnoreCase));
        }

        private void OpenInspectionReport_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject?.InspectionJournal == null || currentObject.InspectionJournal.Count == 0)
            {
                MessageBox.Show("Журнал осмотров пуст.");
                return;
            }

            var activeRows = currentObject.InspectionJournal.Where(x => !x.IsCompletionHistory).ToList();
            var completedRows = currentObject.InspectionJournal.Where(x => x.IsCompletionHistory).ToList();
            var overdueRows = activeRows.Where(x => x.IsDue).OrderBy(x => x.NextReminderDate).ToList();
            var dueTodayRows = activeRows.Where(x => !x.IsDue && x.NextReminderDate.Date == DateTime.Today.Date).ToList();

            var dialog = new Window
            {
                Title = "Отчет по осмотрам",
                Owner = this,
                Width = 980,
                Height = 680,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            dialog.Content = root;

            var summary = new TextBlock
            {
                Text = $"Активных: {activeRows.Count}  |  Просроченных: {overdueRows.Count}  |  На сегодня: {dueTodayRows.Count}  |  Выполнено (история): {completedRows.Count}",
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 10)
            };
            root.Children.Add(summary);

            var tabs = new TabControl();
            Grid.SetRow(tabs, 1);
            root.Children.Add(tabs);

            TabItem BuildTab(string header, IEnumerable<InspectionJournalEntry> rows)
            {
                var grid = new DataGrid
                {
                    AutoGenerateColumns = false,
                    CanUserAddRows = false,
                    CanUserDeleteRows = false,
                    IsReadOnly = true,
                    ItemsSource = rows.ToList(),
                    ColumnWidth = DataGridLength.Auto
                };
                grid.Columns.Add(new DataGridTextColumn { Header = "Журнал", Binding = new Binding(nameof(InspectionJournalEntry.JournalDisplay)), Width = 280 });
                grid.Columns.Add(new DataGridTextColumn { Header = "Осмотр", Binding = new Binding(nameof(InspectionJournalEntry.InspectionDisplay)), Width = 320 });
                grid.Columns.Add(new DataGridTextColumn { Header = "Следующая дата", Binding = new Binding(nameof(InspectionJournalEntry.NextReminderDate)) { StringFormat = "dd.MM.yyyy" }, Width = 120 });
                grid.Columns.Add(new DataGridTextColumn { Header = "Статус", Binding = new Binding(nameof(InspectionJournalEntry.ReminderStatusDisplay)), Width = 240 });
                grid.Columns.Add(new DataGridTextColumn { Header = "Комментарий", Binding = new Binding(nameof(InspectionJournalEntry.NotesDisplay)), Width = 260 });
                DataGridSizingHelper.SetEnableSmartSizing(grid, true);

                return new TabItem
                {
                    Header = header,
                    Content = grid
                };
            }

            tabs.Items.Add(BuildTab("Просроченные", overdueRows));
            tabs.Items.Add(BuildTab("На сегодня", dueTodayRows));
            tabs.Items.Add(BuildTab("Выполненные", completedRows.OrderByDescending(x => x.LastCompletedDate ?? DateTime.MinValue)));

            dialog.ShowDialog();
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
            inspectionLookupsDirty = true;
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
            productionLookupsDirty = true;
            RefreshProductionJournalState();
            if (ProductionJournalGrid != null && row != null && ProductionJournalGrid.Items.Contains(row))
            {
                ProductionJournalGrid.SelectedItem = row;
                ProductionJournalGrid.ScrollIntoView(row);
            }
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

            PushUndo();
            currentObject.ProductionJournal.Remove(row);
            selectedProductionRow = null;
            productionLookupsDirty = true;
            RefreshProductionJournalState();
            SaveState();
            RefreshSummaryTable();
        }

        private void ProductionJournalGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!ShouldHandleGridRowDeleteShortcut(e))
                return;

            if (selectedProductionRow == null && ProductionJournalGrid?.SelectedItem is not ProductionJournalEntry)
                return;

            DeleteProductionRow_Click(this, new RoutedEventArgs());
            e.Handled = true;
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

        private void EnsureProductionAutoFillProfilesStorage()
        {
            EnsureProductionJournalStorage();
            currentObject.ProductionAutoFillProfiles ??= new List<ProductionAutoFillProfile>();

            if (currentObject.ProductionAutoFillProfiles.Count == 0)
            {
                currentObject.ProductionAutoFillProfiles.Add(new ProductionAutoFillProfile
                {
                    Name = "Базовый",
                    Settings = CloneProductionAutoFillSettings(currentObject.ProductionAutoFillSettings)
                });
            }

            if (string.IsNullOrWhiteSpace(currentObject.SelectedProductionAutoFillProfileName))
            {
                currentObject.SelectedProductionAutoFillProfileName = currentObject.ProductionAutoFillProfiles
                    .Select(x => x.Name?.Trim())
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))
                    ?? string.Empty;
            }

            RefreshProductionAutoFillProfileOptions();
        }

        private static ProductionAutoFillSettings CloneProductionAutoFillSettings(ProductionAutoFillSettings source)
        {
            source ??= new ProductionAutoFillSettings();
            return new ProductionAutoFillSettings
            {
                MinQuantityPerRow = source.MinQuantityPerRow,
                MaxQuantityPerRow = source.MaxQuantityPerRow,
                MinRowsPerRun = source.MinRowsPerRun,
                TargetRowsPerRun = source.TargetRowsPerRun,
                MaxRowsPerRun = source.MaxRowsPerRun,
                MaxItemsPerRow = source.MaxItemsPerRow,
                PreferSelectedTypeOnly = source.PreferSelectedTypeOnly,
                UseBalancedDistribution = source.UseBalancedDistribution,
                PreferDemandDeficit = source.PreferDemandDeficit,
                RespectSelectedBlocksAndMarks = source.RespectSelectedBlocksAndMarks,
                AllowMixedMaterialsInRow = source.AllowMixedMaterialsInRow
            };
        }

        private static void CopyProductionAutoFillSettings(ProductionAutoFillSettings source, ProductionAutoFillSettings target)
        {
            if (source == null || target == null)
                return;

            target.MinQuantityPerRow = source.MinQuantityPerRow;
            target.MaxQuantityPerRow = source.MaxQuantityPerRow;
            target.MinRowsPerRun = source.MinRowsPerRun;
            target.TargetRowsPerRun = source.TargetRowsPerRun;
            target.MaxRowsPerRun = source.MaxRowsPerRun;
            target.MaxItemsPerRow = source.MaxItemsPerRow;
            target.PreferSelectedTypeOnly = source.PreferSelectedTypeOnly;
            target.UseBalancedDistribution = source.UseBalancedDistribution;
            target.PreferDemandDeficit = source.PreferDemandDeficit;
            target.RespectSelectedBlocksAndMarks = source.RespectSelectedBlocksAndMarks;
            target.AllowMixedMaterialsInRow = source.AllowMixedMaterialsInRow;
        }

        private void ConfigureProductionAutoFill_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            EnsureProductionJournalStorage();
            EnsureProductionAutoFillProfilesStorage();
            var settings = currentObject.ProductionAutoFillSettings ??= new ProductionAutoFillSettings();
            var working = CloneProductionAutoFillSettings(settings);
            var profileNames = new ObservableCollection<string>(currentObject.ProductionAutoFillProfiles
                .Where(x => !string.IsNullOrWhiteSpace(x.Name))
                .Select(x => x.Name.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase));

            var dialog = new Window
            {
                Title = "Настройки автомастера ПР",
                Owner = this,
                Width = 620,
                Height = 760,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            dialog.Content = root;

            var profilePanel = new Grid { Margin = new Thickness(0, 0, 0, 12) };
            profilePanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            profilePanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            profilePanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            profilePanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            profilePanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetRow(profilePanel, 0);
            root.Children.Add(profilePanel);

            profilePanel.Children.Add(new TextBlock
            {
                Text = "Профиль:",
                VerticalAlignment = VerticalAlignment.Center,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 8, 0)
            });

            var profileBox = new ComboBox
            {
                Margin = new Thickness(0, 0, 8, 0),
                ItemsSource = profileNames,
                IsEditable = false
            };
            var initiallySelectedProfile = currentObject.SelectedProductionAutoFillProfileName?.Trim();
            if (!string.IsNullOrWhiteSpace(initiallySelectedProfile)
                && profileNames.Any(x => string.Equals(x, initiallySelectedProfile, StringComparison.CurrentCultureIgnoreCase)))
            {
                profileBox.SelectedItem = profileNames.First(x => string.Equals(x, initiallySelectedProfile, StringComparison.CurrentCultureIgnoreCase));
            }
            else
            {
                profileBox.SelectedItem = profileNames.FirstOrDefault();
            }
            Grid.SetColumn(profileBox, 1);
            profilePanel.Children.Add(profileBox);

            var applyProfileButton = new Button { Content = "Применить", MinWidth = 110, Style = FindResource("SecondaryButton") as Style };
            Grid.SetColumn(applyProfileButton, 2);
            profilePanel.Children.Add(applyProfileButton);
            var saveProfileButton = new Button { Content = "Сохранить как", MinWidth = 130, Margin = new Thickness(8, 0, 0, 0), Style = FindResource("SecondaryButton") as Style };
            Grid.SetColumn(saveProfileButton, 3);
            profilePanel.Children.Add(saveProfileButton);
            var deleteProfileButton = new Button { Content = "Удалить", MinWidth = 100, Margin = new Thickness(8, 0, 0, 0), Style = FindResource("SecondaryButton") as Style };
            Grid.SetColumn(deleteProfileButton, 4);
            profilePanel.Children.Add(deleteProfileButton);

            var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            Grid.SetRow(scroll, 1);
            root.Children.Add(scroll);

            var contentPanel = new StackPanel();
            scroll.Content = contentPanel;

            TextBox CreateNumericBox(string text, string caption)
            {
                var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };
                panel.Children.Add(new TextBlock { Text = caption, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 6) });
                var box = new TextBox { Text = text };
                panel.Children.Add(box);
                contentPanel.Children.Add(panel);
                return box;
            }

            var minBox = CreateNumericBox(working.MinQuantityPerRow.ToString(CultureInfo.InvariantCulture), "Минимум в строке");
            var maxBox = CreateNumericBox(working.MaxQuantityPerRow.ToString(CultureInfo.InvariantCulture), "Максимум в строке");
            var minRowsBox = CreateNumericBox(working.MinRowsPerRun.ToString(CultureInfo.InvariantCulture), "Минимум строк за один запуск");
            var targetRowsBox = CreateNumericBox(working.TargetRowsPerRun.ToString(CultureInfo.InvariantCulture), "Целевое число строк");
            var maxRowsBox = CreateNumericBox(working.MaxRowsPerRun.ToString(CultureInfo.InvariantCulture), "Максимум строк за один запуск");
            var itemsPerRowBox = CreateNumericBox(working.MaxItemsPerRow.ToString(CultureInfo.InvariantCulture), "Максимум элементов в одной строке");

            var preferTypeCheck = new CheckBox
            {
                Content = "Брать материалы только из выбранного типа",
                IsChecked = working.PreferSelectedTypeOnly,
                Margin = new Thickness(0, 0, 0, 10)
            };
            contentPanel.Children.Add(preferTypeCheck);

            var balanceCheck = new CheckBox
            {
                Content = "Распределять количество более равномерно",
                IsChecked = working.UseBalancedDistribution,
                Margin = new Thickness(0, 0, 0, 10)
            };
            contentPanel.Children.Add(balanceCheck);

            var deficitCheck = new CheckBox
            {
                Content = "Сначала закрывать дефицит по сводке",
                IsChecked = working.PreferDemandDeficit,
                Margin = new Thickness(0, 0, 0, 10)
            };
            contentPanel.Children.Add(deficitCheck);

            var selectedCellsCheck = new CheckBox
            {
                Content = "Учитывать только выбранные блоки и отметки из формы",
                IsChecked = working.RespectSelectedBlocksAndMarks,
                Margin = new Thickness(0, 0, 0, 10)
            };
            contentPanel.Children.Add(selectedCellsCheck);

            var mixedRowsCheck = new CheckBox
            {
                Content = "Разрешать несколько элементов в одной строке",
                IsChecked = working.AllowMixedMaterialsInRow,
                Margin = new Thickness(0, 0, 0, 10)
            };
            contentPanel.Children.Add(mixedRowsCheck);

            void LoadUiFromSettings(ProductionAutoFillSettings source)
            {
                minBox.Text = source.MinQuantityPerRow.ToString(CultureInfo.InvariantCulture);
                maxBox.Text = source.MaxQuantityPerRow.ToString(CultureInfo.InvariantCulture);
                minRowsBox.Text = source.MinRowsPerRun.ToString(CultureInfo.InvariantCulture);
                targetRowsBox.Text = source.TargetRowsPerRun.ToString(CultureInfo.InvariantCulture);
                maxRowsBox.Text = source.MaxRowsPerRun.ToString(CultureInfo.InvariantCulture);
                itemsPerRowBox.Text = source.MaxItemsPerRow.ToString(CultureInfo.InvariantCulture);
                preferTypeCheck.IsChecked = source.PreferSelectedTypeOnly;
                balanceCheck.IsChecked = source.UseBalancedDistribution;
                deficitCheck.IsChecked = source.PreferDemandDeficit;
                selectedCellsCheck.IsChecked = source.RespectSelectedBlocksAndMarks;
                mixedRowsCheck.IsChecked = source.AllowMixedMaterialsInRow;
            }

            void ReadUiToSettings(ProductionAutoFillSettings target)
            {
                target.MinQuantityPerRow = Math.Max(1, int.TryParse(minBox.Text, out var min) ? min : 4);
                target.MaxQuantityPerRow = Math.Max(target.MinQuantityPerRow, int.TryParse(maxBox.Text, out var max) ? max : 8);
                target.MinRowsPerRun = Math.Clamp(int.TryParse(minRowsBox.Text, out var minRows) ? minRows : 4, 1, 12);
                target.TargetRowsPerRun = Math.Clamp(int.TryParse(targetRowsBox.Text, out var targetRows) ? targetRows : 5, target.MinRowsPerRun, 16);
                target.MaxRowsPerRun = Math.Clamp(int.TryParse(maxRowsBox.Text, out var maxRows) ? maxRows : 6, target.TargetRowsPerRun, 20);
                target.MaxItemsPerRow = Math.Clamp(int.TryParse(itemsPerRowBox.Text, out var maxItems) ? maxItems : 2, 1, 6);
                target.PreferSelectedTypeOnly = preferTypeCheck.IsChecked == true;
                target.UseBalancedDistribution = balanceCheck.IsChecked == true;
                target.PreferDemandDeficit = deficitCheck.IsChecked == true;
                target.RespectSelectedBlocksAndMarks = selectedCellsCheck.IsChecked == true;
                target.AllowMixedMaterialsInRow = mixedRowsCheck.IsChecked == true;
            }

            ProductionAutoFillProfile FindProfile(string profileName)
                => currentObject.ProductionAutoFillProfiles.FirstOrDefault(x =>
                    string.Equals(x.Name, profileName, StringComparison.CurrentCultureIgnoreCase));

            void ApplySelectedProfile()
            {
                var selectedProfileName = profileBox.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(selectedProfileName))
                    return;

                var profile = FindProfile(selectedProfileName);
                if (profile?.Settings == null)
                    return;

                currentObject.SelectedProductionAutoFillProfileName = selectedProfileName.Trim();
                working = CloneProductionAutoFillSettings(profile.Settings);
                LoadUiFromSettings(working);
            }

            applyProfileButton.Click += (_, _) => ApplySelectedProfile();
            profileBox.SelectionChanged += (_, _) => ApplySelectedProfile();

            saveProfileButton.Click += (_, _) =>
            {
                ReadUiToSettings(working);
                var inputName = Microsoft.VisualBasic.Interaction.InputBox(
                    "Введите название профиля автомастера ПР:",
                    "Сохранить профиль",
                    profileBox.Text?.Trim() ?? string.Empty)?.Trim();

                if (string.IsNullOrWhiteSpace(inputName))
                    return;

                var existing = FindProfile(inputName);
                if (existing == null)
                {
                    existing = new ProductionAutoFillProfile { Name = inputName };
                    currentObject.ProductionAutoFillProfiles.Add(existing);
                }

                existing.Name = inputName;
                existing.Settings = CloneProductionAutoFillSettings(working);

                if (!profileNames.Contains(inputName))
                    profileNames.Add(inputName);

                var sortedNames = profileNames.OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase).ToList();
                profileNames.Clear();
                foreach (var name in sortedNames)
                    profileNames.Add(name);

                profileBox.SelectedItem = inputName;
                currentObject.SelectedProductionAutoFillProfileName = inputName;
                RefreshProductionAutoFillProfileOptions();
                SaveState();
            };

            deleteProfileButton.Click += (_, _) =>
            {
                var selectedProfileName = profileBox.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(selectedProfileName))
                {
                    MessageBox.Show("Выберите профиль для удаления.");
                    return;
                }

                var profile = FindProfile(selectedProfileName);
                if (profile == null)
                    return;

                if (MessageBox.Show($"Удалить профиль \"{selectedProfileName}\"?", "Автомастер ПР", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                    return;

                currentObject.ProductionAutoFillProfiles.Remove(profile);
                profileNames.Remove(selectedProfileName);
                profileBox.SelectedItem = profileNames.FirstOrDefault();
                currentObject.SelectedProductionAutoFillProfileName = profileBox.SelectedItem?.ToString() ?? string.Empty;
                RefreshProductionAutoFillProfileOptions();
                SaveState();
            };

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            var saveButton = new Button { Content = "Сохранить", MinWidth = 120, IsDefault = true };
            var cancelButton = new Button { Content = "Отмена", Style = FindResource("SecondaryButton") as Style, MinWidth = 110, Margin = new Thickness(10, 0, 0, 0), IsCancel = true };
            footer.Children.Add(saveButton);
            footer.Children.Add(cancelButton);

            saveButton.Click += (_, _) =>
            {
                ReadUiToSettings(working);
                CopyProductionAutoFillSettings(working, settings);
                currentObject.SelectedProductionAutoFillProfileName = profileBox.SelectedItem?.ToString() ?? currentObject.SelectedProductionAutoFillProfileName;
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
            EnsureProductionAutoFillProfilesStorage();
            var settings = ResolveProductionAutoFillSettingsForRun();
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
                var rowWork = string.IsNullOrWhiteSpace(baseWork) ? groupForMarks : baseWork;
                var rowDeviation = ResolveAutoFillDeviation(rowWork, baseDeviation);
                var row = new ProductionJournalEntry
                {
                    Date = baseDate,
                    ActionName = baseAction,
                    WorkName = rowWork,
                    ElementsText = FormatProductionItems(plannedItems),
                    BlocksText = baseBlocks,
                    MarksText = baseMarks,
                    BrigadeName = baseBrigade,
                    Weather = baseWeather,
                    Deviations = rowDeviation,
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
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (selectedProductionRow == null || ProductionJournalGrid == null)
                    return;

                if (ProductionJournalGrid.Items.Contains(selectedProductionRow))
                {
                    ProductionJournalGrid.SelectedItem = selectedProductionRow;
                    ProductionJournalGrid.ScrollIntoView(selectedProductionRow);
                }
            }), DispatcherPriority.Background);
            SaveState();
            MessageBox.Show($"Автозаполнение добавило {addedRows.Count} строк(и) в ПР по выбранным блокам и отметкам.");
        }

        private ProductionAutoFillSettings ResolveProductionAutoFillSettingsForRun()
        {
            EnsureProductionAutoFillProfilesStorage();
            var fallback = CloneProductionAutoFillSettings(currentObject?.ProductionAutoFillSettings);
            if (currentObject?.ProductionAutoFillProfiles == null || currentObject.ProductionAutoFillProfiles.Count == 0)
                return fallback;

            var selectedProfileName = ProductionAutoFillProfileBox?.SelectedItem?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(selectedProfileName))
                selectedProfileName = currentObject.SelectedProductionAutoFillProfileName?.Trim();

            if (string.IsNullOrWhiteSpace(selectedProfileName))
                return fallback;

            var profile = currentObject.ProductionAutoFillProfiles.FirstOrDefault(x =>
                string.Equals(x.Name?.Trim(), selectedProfileName, StringComparison.CurrentCultureIgnoreCase));
            if (profile?.Settings == null)
                return fallback;

            currentObject.SelectedProductionAutoFillProfileName = profile.Name?.Trim() ?? string.Empty;
            return CloneProductionAutoFillSettings(profile.Settings);
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
            ProductionBrigadeBox.Text = timesheetAssignableBrigades.FirstOrDefault() ?? string.Empty;
            ProductionWeatherBox.Text = string.Empty;
            ProductionDeviationBox.Text = string.Empty;
            ProductionHiddenWorksCheckBox.IsChecked = false;
            RefreshProductionElementOptions();
            RefreshProductionMarkOptions();
            RefreshProductionDeviationOptions();
        }

        private void FillProductionForm(ProductionJournalEntry row)
        {
            if (row == null || ProductionDatePicker == null)
                return;

            ProductionDatePicker.SelectedDate = row.Date;
            ProductionActionBox.Text = row.ActionName ?? string.Empty;
            ProductionWorkBox.Text = row.WorkName ?? string.Empty;
            RefreshProductionElementOptions();
            RefreshProductionMarkOptions();
            RefreshProductionDeviationOptions();
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

        private void SaveProductionTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            EnsureProductionJournalStorage();
            var name = Microsoft.VisualBasic.Interaction.InputBox(
                "Введите название шаблона ПР:",
                "Сохранить шаблон",
                (ProductionWorkBox?.Text ?? string.Empty).Trim())?.Trim();

            if (string.IsNullOrWhiteSpace(name))
                return;

            var template = new ProductionJournalTemplate
            {
                Name = name,
                ActionName = ProductionActionBox?.Text?.Trim() ?? string.Empty,
                WorkName = ProductionWorkBox?.Text?.Trim() ?? string.Empty,
                ElementsText = ProductionElementsBox?.Text?.Trim() ?? string.Empty,
                BlocksText = ProductionBlocksBox?.Text?.Trim() ?? string.Empty,
                MarksText = ProductionMarksBox?.Text?.Trim() ?? string.Empty,
                BrigadeName = ProductionBrigadeBox?.Text?.Trim() ?? string.Empty,
                Weather = ProductionWeatherBox?.Text?.Trim() ?? string.Empty,
                Deviations = ProductionDeviationBox?.Text?.Trim() ?? string.Empty,
                RequiresHiddenWorkAct = ProductionHiddenWorksCheckBox?.IsChecked == true
            };

            var existing = currentObject.ProductionTemplates
                .FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.CurrentCultureIgnoreCase));
            if (existing != null)
                currentObject.ProductionTemplates.Remove(existing);

            currentObject.ProductionTemplates.Add(template);
            SaveState();
        }

        private void ApplyProductionTemplate_Click(object sender, RoutedEventArgs e)
        {
            var template = SelectProductionTemplate();
            if (template == null)
                return;

            ProductionDatePicker.SelectedDate ??= DateTime.Today;
            ProductionActionBox.Text = template.ActionName ?? string.Empty;
            ProductionWorkBox.Text = template.WorkName ?? string.Empty;
            RefreshProductionElementOptions();
            RefreshProductionMarkOptions();
            RefreshProductionDeviationOptions();
            ProductionElementsBox.Text = template.ElementsText ?? string.Empty;
            ProductionBlocksBox.Text = template.BlocksText ?? string.Empty;
            ProductionMarksBox.Text = template.MarksText ?? string.Empty;
            ProductionBrigadeBox.Text = template.BrigadeName ?? string.Empty;
            ProductionWeatherBox.Text = template.Weather ?? string.Empty;
            ProductionDeviationBox.Text = template.Deviations ?? string.Empty;
            ProductionHiddenWorksCheckBox.IsChecked = template.RequiresHiddenWorkAct;
        }

        private void CreateProductionFromTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            var template = SelectProductionTemplate();
            if (template == null)
                return;

            var countText = Microsoft.VisualBasic.Interaction.InputBox(
                "Сколько одинаковых записей добавить по шаблону?",
                "Добавить по шаблону ПР",
                "1")?.Trim();
            if (string.IsNullOrWhiteSpace(countText))
                return;

            if (!int.TryParse(countText, out var count) || count <= 0)
                count = 1;

            count = Math.Clamp(count, 1, 31);
            var baseDate = ProductionDatePicker?.SelectedDate ?? DateTime.Today;
            var added = 0;

            for (var i = 0; i < count; i++)
            {
                var row = new ProductionJournalEntry
                {
                    Date = baseDate.AddDays(i),
                    ActionName = template.ActionName ?? string.Empty,
                    WorkName = template.WorkName ?? string.Empty,
                    ElementsText = template.ElementsText ?? string.Empty,
                    BlocksText = template.BlocksText ?? string.Empty,
                    MarksText = template.MarksText ?? string.Empty,
                    BrigadeName = template.BrigadeName ?? string.Empty,
                    Weather = template.Weather ?? string.Empty,
                    Deviations = template.Deviations ?? string.Empty,
                    RequiresHiddenWorkAct = template.RequiresHiddenWorkAct
                };

                if (!ApplyProductionRowChanges(row))
                    continue;

                currentObject.ProductionJournal.Add(row);
                added++;
            }

            if (added == 0)
            {
                MessageBox.Show("Не удалось добавить записи по шаблону. Проверьте доступные остатки.");
                return;
            }

            productionLookupsDirty = true;
            RefreshProductionJournalState();
            SaveState();
            MessageBox.Show($"Добавлено строк по шаблону: {added}.");
        }

        private ProductionJournalTemplate SelectProductionTemplate()
        {
            if (currentObject?.ProductionTemplates == null || currentObject.ProductionTemplates.Count == 0)
            {
                MessageBox.Show("Сначала сохраните хотя бы один шаблон ПР.");
                return null;
            }

            var options = currentObject.ProductionTemplates
                .Where(x => !string.IsNullOrWhiteSpace(x.Name))
                .OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var selectedName = PromptSelectOption("Выберите шаблон ПР", "Шаблон", options.Select(x => x.Name));
            if (string.IsNullOrWhiteSpace(selectedName))
                return null;

            return options.FirstOrDefault(x => string.Equals(x.Name, selectedName, StringComparison.CurrentCultureIgnoreCase));
        }

        private string PromptSelectOption(string title, string caption, IEnumerable<string> options)
        {
            var items = (options ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (items.Count == 0)
                return null;

            var dialog = new Window
            {
                Title = title,
                Owner = this,
                Width = 460,
                Height = 520,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            dialog.Content = root;

            root.Children.Add(new TextBlock
            {
                Text = caption,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            });

            var list = new ListBox
            {
                ItemsSource = items
            };
            Grid.SetRow(list, 1);
            root.Children.Add(list);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            var okButton = new Button { Content = "Выбрать", MinWidth = 110, IsDefault = true };
            var cancelButton = new Button { Content = "Отмена", MinWidth = 110, IsCancel = true, Style = FindResource("SecondaryButton") as Style, Margin = new Thickness(8, 0, 0, 0) };
            footer.Children.Add(okButton);
            footer.Children.Add(cancelButton);
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            string selected = null;
            okButton.Click += (_, _) =>
            {
                selected = list.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(selected))
                    return;

                dialog.DialogResult = true;
            };

            list.MouseDoubleClick += (_, _) =>
            {
                selected = list.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(selected))
                    return;

                dialog.DialogResult = true;
            };

            return dialog.ShowDialog() == true ? selected : null;
        }

        private bool ApplyProductionRowChanges(ProductionJournalEntry row)
        {
            if (row == null || currentObject?.ProductionJournal == null)
                return false;

            ApplyProductionDefaults(row);
            row.IsAutoCorrectedQuantity = false;

            var originalItems = ParseProductionItems(row.ElementsText);
            var adjustedMessages = new List<string>();
            var adjustedItems = AdjustProductionItems(row, originalItems, adjustedMessages, out var hasCorrections);
            row.IsAutoCorrectedQuantity = hasCorrections;
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
                var mappedDeviation = GetDeviationOptionsForWork(row.WorkName, includeCurrentJournal: false)
                    .FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(mappedDeviation))
                {
                    row.Deviations = mappedDeviation;
                }

                var previousDeviation = currentObject?.ProductionJournal?
                    .Where(x => !ReferenceEquals(x, row)
                        && string.Equals(x.ActionName?.Trim(), row.ActionName?.Trim(), StringComparison.CurrentCultureIgnoreCase)
                        && string.Equals(x.WorkName?.Trim(), row.WorkName?.Trim(), StringComparison.CurrentCultureIgnoreCase)
                        && !string.IsNullOrWhiteSpace(x.Deviations))
                    .OrderByDescending(x => x.Date)
                    .Select(x => x.Deviations?.Trim())
                    .FirstOrDefault();

                if (string.IsNullOrWhiteSpace(row.Deviations) && !string.IsNullOrWhiteSpace(previousDeviation))
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
            row.IsAutoCorrectedQuantity = snapshot.IsAutoCorrectedQuantity;
            row.IsGeneratedCompanion = snapshot.IsGeneratedCompanion;
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

        private List<ProductionItemQuantity> AdjustProductionItems(ProductionJournalEntry row, List<ProductionItemQuantity> items, List<string> adjustedMessages, out bool hasCorrections)
        {
            var result = new List<ProductionItemQuantity>();
            hasCorrections = false;

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
                var finalQuantity = CalculationCore.ClampToAvailable(item.Quantity, arrived, alreadyMounted);
                if (finalQuantity <= 0)
                {
                    hasCorrections = true;
                    adjustedMessages?.Add($"\"{item.MaterialName}\" не добавлен: доступное количество закончилось.");
                    continue;
                }

                if (CalculationCore.HasDifference(finalQuantity, item.Quantity) && finalQuantity < item.Quantity)
                {
                    hasCorrections = true;
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
                RequiresHiddenWorkAct = row.RequiresHiddenWorkAct,
                IsGeneratedCompanion = true
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
                    var group = FindMaterialGroupByName(item.MaterialName) ?? row.WorkName?.Trim();
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

        private void RefreshProductionBlockDisplayValues()
        {
            if (currentObject?.ProductionJournal == null)
                return;

            foreach (var row in currentObject.ProductionJournal)
                row.BlocksDisplayText = BuildProductionBlocksDisplayText(row?.BlocksText);
        }

        private string BuildProductionBlocksDisplayText(string blocksText)
        {
            var blocks = LevelMarkHelper.ParseBlocks(blocksText);
            if (blocks.Count == 0)
                return blocksText?.Trim() ?? string.Empty;

            return string.Join(", ", blocks.Select(GetProductionBlockLabel));
        }

        private string GetProductionBlockLabel(int block)
        {
            if (block <= 0)
                return string.Empty;

            if (currentObject?.BlockAxesByNumber != null
                && currentObject.BlockAxesByNumber.TryGetValue(block, out var axes)
                && !string.IsNullOrWhiteSpace(axes))
            {
                return axes.Trim();
            }

            return block.ToString(CultureInfo.CurrentCulture);
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
                var group = FindMaterialGroupByName(item.MaterialName) ?? row.WorkName?.Trim();
                if (string.IsNullOrWhiteSpace(group))
                    continue;

                var demandKey = BuildDemandKey(group, item.MaterialName);
                var fallbackUnit = GetUnitForMaterial(group, item.MaterialName);
                var demand = GetOrCreateDemand(demandKey, fallbackUnit);
                if (demand != null && string.IsNullOrWhiteSpace(demand.Unit))
                    demand.Unit = fallbackUnit;

                var unit = !string.IsNullOrWhiteSpace(demand?.Unit) ? demand.Unit : fallbackUnit;
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

                        parts.Add($"{item.MaterialName}: {GetProductionBlockLabel(block)} {mark} — остаток {FormatNumberByUnit(remaining, unit)}");
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

        private string NormalizeAccessRole(string role)
        {
            var normalized = (role ?? string.Empty).Trim();
            if (string.Equals(normalized, ProjectAccessRoles.View, StringComparison.CurrentCultureIgnoreCase)
                || normalized.Contains("РџСЂРѕСЃ", StringComparison.OrdinalIgnoreCase))
            {
                return ProjectAccessRoles.View;
            }

            if (string.Equals(normalized, ProjectAccessRoles.Edit, StringComparison.CurrentCultureIgnoreCase)
                || normalized.Contains("Р РµРґР°РєС‚", StringComparison.OrdinalIgnoreCase))
            {
                return ProjectAccessRoles.Edit;
            }

            return ProjectAccessRoles.Critical;
        }

        private string GetCurrentAccessRole()
        {
            EnsureProjectUiSettings();
            return NormalizeAccessRole(currentObject?.UiSettings?.AccessRole);
        }

        private bool EnsureCanEditOperation(string operationName)
        {
            var role = GetCurrentAccessRole();
            if (!string.Equals(role, ProjectAccessRoles.View, StringComparison.CurrentCultureIgnoreCase))
                return true;

            MessageBox.Show(
                $"Операция \"{operationName}\" недоступна для роли \"{ProjectAccessRoles.View}\".",
                "Недостаточно прав",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        private bool EnsureCanRunCriticalOperation(string operationName, bool requireCode = true)
        {
            if (!EnsureCanEditOperation(operationName))
                return false;

            var role = GetCurrentAccessRole();
            if (!string.Equals(role, ProjectAccessRoles.Critical, StringComparison.CurrentCultureIgnoreCase))
            {
                MessageBox.Show(
                    $"Операция \"{operationName}\" доступна только для роли \"{ProjectAccessRoles.Critical}\".",
                    "Недостаточно прав",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return false;
            }

            if (requireCode && (currentObject?.UiSettings?.RequireCodeForCriticalOperations ?? true))
                return ConfirmOperationWithCode(operationName);

            return true;
        }

        private bool ConfirmOperationWithCode(string operationName)
        {
            var code = Random.Shared.Next(100000, 1000000).ToString(CultureInfo.InvariantCulture);
            var entered = Microsoft.VisualBasic.Interaction.InputBox(
                $"Подтвердите операцию \"{operationName}\".\nВведите код: {code}",
                "Подтверждение операции",
                string.Empty)?.Trim();

            if (string.Equals(entered, code, StringComparison.Ordinal))
                return true;

            MessageBox.Show("Код введен неверно. Операция отменена.", "Подтверждение операции", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }


        // ================= МЕНЮ =================

        private void IntegrityCheck_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();

            var (issues, warnings) = ValidateDataIntegrity();
            var sb = new StringBuilder();
            sb.AppendLine("Проверка целостности завершена.");
            sb.AppendLine($"Ошибок: {issues.Count}");
            sb.AppendLine($"Предупреждений: {warnings.Count}");
            sb.AppendLine();

            if (issues.Count > 0)
            {
                sb.AppendLine("Ошибки:");
                foreach (var item in issues.Take(40))
                    sb.AppendLine($"• {item}");
                if (issues.Count > 40)
                    sb.AppendLine($"• ... и ещё {issues.Count - 40}");
                sb.AppendLine();
            }

            if (warnings.Count > 0)
            {
                sb.AppendLine("Предупреждения:");
                foreach (var item in warnings.Take(40))
                    sb.AppendLine($"• {item}");
                if (warnings.Count > 40)
                    sb.AppendLine($"• ... и ещё {warnings.Count - 40}");
            }

            var title = issues.Count > 0 ? "Проверка целостности: найдены проблемы" : "Проверка целостности";
            ShowLargeTextDialog(title, sb.ToString());

            AppendChangeLog("Проверка целостности", $"Ошибок: {issues.Count}, предупреждений: {warnings.Count}");
            SaveState(SaveTrigger.System);
        }

        private void ExportDiagnostics_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            var dialog = new SaveFileDialog
            {
                Filter = "ConstructionControl diagnostics (*.ccdiag)|*.ccdiag|ZIP (*.zip)|*.zip",
                FileName = $"diagnostics_{DateTime.Now:yyyyMMdd_HHmm}.ccdiag"
            };
            if (dialog.ShowDialog() != true)
                return;

            var (issues, warnings) = ValidateDataIntegrity();
            var report = BuildDiagnosticsReport(issues, warnings);
            var changeLog = currentObject?.ChangeLog == null
                ? string.Empty
                : string.Join(Environment.NewLine, currentObject.ChangeLog
                    .OrderByDescending(x => x.TimestampUtc)
                    .Take(200)
                    .Select(FormatChangeLogEntry));
            var stateSnapshot = BuildCurrentStateSnapshotJson();

            using (var stream = new FileStream(dialog.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteTextEntry(archive, "diagnostics.txt", report);
                WriteTextEntry(archive, "changelog.txt", changeLog);
                WriteTextEntry(archive, "state_snapshot.json", stateSnapshot);
            }

            AppendChangeLog("Экспорт диагностики", $"Сформирован файл диагностики: {dialog.FileName}");
            SaveState(SaveTrigger.System);
            MessageBox.Show("Диагностика экспортирована.", "Надежность", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void CreateRecoveryPackage_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanEditOperation("создание пакета восстановления"))
                return;

            CommitOpenEdits();
            var dialog = new SaveFileDialog
            {
                Filter = "ConstructionControl recovery (*.ccrecovery)|*.ccrecovery",
                FileName = $"recovery_{DateTime.Now:yyyyMMdd_HHmm}.ccrecovery"
            };
            if (dialog.ShowDialog() != true)
                return;

            var state = new AppState
            {
                SchemaVersion = AppState.LatestSchemaVersion,
                SavedAtUtc = DateTime.UtcNow,
                CurrentObject = currentObject,
                Journal = journal
            };

            var stateJson = JsonSerializer.Serialize(state);
            var (issues, warnings) = ValidateDataIntegrity();
            var diagnostics = BuildDiagnosticsReport(issues, warnings);
            var storageFiles = 0;

            using (var stream = new FileStream(dialog.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteTextEntry(archive, "state.json", stateJson);
                WriteTextEntry(archive, "diagnostics.txt", diagnostics);
                WriteTextEntry(archive, "metadata.json", JsonSerializer.Serialize(new
                {
                    CreatedAtUtc = DateTime.UtcNow,
                    AppVersion = GetType().Assembly.GetName().Version?.ToString() ?? "unknown",
                    SchemaVersion = AppState.LatestSchemaVersion,
                    SaveFile = currentSaveFileName
                }));
                storageFiles = ExportProjectStorageToArchive(archive);
            }

            AppendChangeLog("Пакет восстановления", $"Создан пакет восстановления: {dialog.FileName} (файлов документов: {storageFiles}).");
            SaveState(SaveTrigger.System);
            MessageBox.Show("Пакет восстановления создан.", "Надежность", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void RestoreFromRecoveryPackage_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanRunCriticalOperation("восстановление из пакета восстановления"))
                return;

            CommitOpenEdits();
            var dialog = new OpenFileDialog
            {
                Filter = "ConstructionControl recovery (*.ccrecovery)|*.ccrecovery|ConstructionControl backup (*.ccbak)|*.ccbak"
            };
            if (dialog.ShowDialog() != true)
                return;

            try
            {
                _ = CreateSafetyBackupBeforeOperation("before_restore_recovery_package");
            }
            catch (Exception backupEx)
            {
                MessageBox.Show(
                    $"Не удалось создать предохранительный бэкап перед восстановлением.{Environment.NewLine}{backupEx.Message}",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }

            if (!TryLoadBackupState(dialog.FileName, out var state, out var restoredStorageFiles, out var importError) || state == null)
            {
                MessageBox.Show(
                    string.IsNullOrWhiteSpace(importError)
                        ? "Не удалось восстановить проект из пакета."
                        : importError,
                    "Ошибка восстановления",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            PushUndo();
            RestoreState(state);
            RebuildArchiveFromCurrentData();
            AppendChangeLog("Восстановление", $"Восстановлены данные из пакета: {dialog.FileName}");
            SaveState(SaveTrigger.System);
            RefreshTreePreserveState();
            RefreshArrivalTypes();
            RefreshArrivalNames();
            RefreshDocumentLibraries();
            ApplyAllFilters();
            RequestReminderRefresh(immediate: true);
            MessageBox.Show($"Восстановление завершено. Восстановлено файлов ПДФ/смет: {restoredStorageFiles}.", "Восстановление", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private string BuildDiagnosticsReport(List<string> issues, List<string> warnings)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Дата: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            sb.AppendLine($"Пользователь: {Environment.UserName}");
            sb.AppendLine($"Версия: {GetType().Assembly.GetName().Version}");
            sb.AppendLine($"Файл проекта: {currentSaveFileName}");
            sb.AppendLine($"Схема: {AppState.LatestSchemaVersion}");
            sb.AppendLine($"Роль доступа: {GetCurrentAccessRole()}");
            sb.AppendLine($"Безопасный старт: {(IsSafeStartupEnabled() ? "вкл" : "выкл")}");
            sb.AppendLine();
            sb.AppendLine($"Ошибок целостности: {issues?.Count ?? 0}");
            if (issues != null)
            {
                foreach (var issue in issues.Take(200))
                    sb.AppendLine($"[ERR] {issue}");
            }

            sb.AppendLine();
            sb.AppendLine($"Предупреждений: {warnings?.Count ?? 0}");
            if (warnings != null)
            {
                foreach (var warning in warnings.Take(200))
                    sb.AppendLine($"[WRN] {warning}");
            }

            if (lastStorageIntegrityIssues.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Проблемы хранилища документов:");
                foreach (var issue in lastStorageIntegrityIssues.Take(200))
                    sb.AppendLine($"[DOC] {issue}");
            }

            if (tabOpenDiagnostics.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Диагностика открытия вкладок:");
                foreach (var pair in tabOpenDiagnostics.OrderByDescending(x => x.Value))
                    sb.AppendLine($"{pair.Key}: {pair.Value.TotalMilliseconds:0} ms");
            }

            if (currentObject != null)
            {
                sb.AppendLine();
                sb.AppendLine("Счетчики данных:");
                sb.AppendLine($"Приход: {journal?.Count ?? 0}");
                sb.AppendLine($"Потребность: {currentObject.Demand?.Count ?? 0}");
                sb.AppendLine($"ОТ: {currentObject.OtJournal?.Count ?? 0}");
                sb.AppendLine($"Табель: {currentObject.TimesheetPeople?.Count ?? 0}");
                sb.AppendLine($"ПР: {currentObject.ProductionJournal?.Count ?? 0}");
                sb.AppendLine($"Осмотры: {currentObject.InspectionJournal?.Count ?? 0}");
                sb.AppendLine($"PDF узлы: {CountDocumentNodes(currentObject.PdfDocuments)}");
                sb.AppendLine($"Сметы узлы: {CountDocumentNodes(currentObject.EstimateDocuments)}");
            }

            return sb.ToString();
        }

        private static int CountDocumentNodes(IEnumerable<DocumentTreeNode> nodes)
        {
            if (nodes == null)
                return 0;

            var total = 0;
            foreach (var node in nodes)
            {
                if (node == null)
                    continue;

                total++;
                if (node.Children != null && node.Children.Count > 0)
                    total += CountDocumentNodes(node.Children);
            }

            return total;
        }

        private static void WriteTextEntry(ZipArchive archive, string entryName, string text)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, Encoding.UTF8);
            writer.Write(text ?? string.Empty);
        }

        private void ShowChangeLog_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            currentObject.ChangeLog ??= new List<ProjectChangeLogEntry>();
            if (currentObject.ChangeLog.Count == 0)
            {
                MessageBox.Show("Журнал изменений пока пуст.", "Журнал изменений", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var text = string.Join(
                Environment.NewLine,
                currentObject.ChangeLog
                    .OrderByDescending(x => x.TimestampUtc)
                    .Take(300)
                    .Select(FormatChangeLogEntry));

            ShowLargeTextDialog("Журнал изменений", text);
        }

        private void RestoreFromAutoBackup_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanRunCriticalOperation("восстановление из автосохранения"))
                return;

            var latestBackup = GetLatestAutoBackupFile(currentSaveFileName);
            if (string.IsNullOrWhiteSpace(latestBackup) || !File.Exists(latestBackup))
            {
                MessageBox.Show("Автосохранения не найдены.", "Восстановление", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var backupTimeLocal = File.GetLastWriteTime(latestBackup);
            var result = MessageBox.Show(
                $"Найдено последнее автосохранение от {backupTimeLocal:dd.MM.yyyy HH:mm:ss}.{Environment.NewLine}Восстановить его?",
                "Восстановление из автосохранения",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                _ = CreateSafetyBackupBeforeOperation("before_restore_from_autosave");
                var backupJson = File.ReadAllText(latestBackup);
                if (!TryDeserializeAppState(backupJson, out _, out var parseError))
                    throw new InvalidDataException($"Автосохранение повреждено: {parseError}");

                SaveStateJsonTransactional(backupJson, currentSaveFileName);
                LoadState();
                ArrivalPanel.SetObject(currentObject, journal);
                RefreshTreePreserveState();
                RefreshSummaryTable();
                RefreshArrivalTypes();
                RefreshArrivalNames();
                RefreshProductionJournalState();
                RefreshInspectionJournalState();
                ApplyAllFilters();
                RequestReminderRefresh(immediate: true);

                AppendChangeLog("Восстановление", $"Выполнено восстановление из автосохранения {System.IO.Path.GetFileName(latestBackup)}");
                SaveState(SaveTrigger.System);
                MessageBox.Show("Восстановление завершено.", "Восстановление", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось восстановить данные из автосохранения.{Environment.NewLine}{ex.Message}",
                    "Ошибка восстановления",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void ShowLargeTextDialog(string title, string text)
        {
            var viewer = new Window
            {
                Title = title,
                Owner = this,
                Width = 860,
                Height = 620,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                MinWidth = 640,
                MinHeight = 420,
                Background = (Brush)FindResource("AppBgBrush")
            };

            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var textBox = new TextBox
            {
                Text = text ?? string.Empty,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(12)
            };
            Grid.SetRow(textBox, 0);
            root.Children.Add(textBox);

            var closeButton = new Button
            {
                Content = "Закрыть",
                Width = 120,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            closeButton.Click += (_, _) => viewer.Close();
            Grid.SetRow(closeButton, 1);
            root.Children.Add(closeButton);

            viewer.Content = root;
            viewer.ShowDialog();
        }

        private void ObjectWizard_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanEditOperation("мастер объекта"))
                return;

            CommitOpenEdits();
            var hasCurrentObject = currentObject != null;
            var working = CloneProjectObjectForWizard(currentObject);
            var window = new ObjectSettingsWindow(working, allowCreateAsNew: hasCurrentObject)
            {
                Owner = this
            };

            if (window.ShowDialog() != true)
                return;

            var result = window.ResultObject ?? working;
            if (!hasCurrentObject || window.CreateAsNewObject)
            {
                currentObject = new ProjectObject
                {
                    BlocksCount = 1,
                    FloorsPerBlock = 1,
                    SameFloorsInBlocks = true
                };
                ApplyObjectMetadata(currentObject, result);

                journal.Clear();
                EnsureOtJournalStorage();
                BindOtJournal();
                ArrivalPanel.SetObject(currentObject, journal);

                timesheetInitialized = false;
                productionJournalInitialized = false;
                inspectionJournalInitialized = false;
                timesheetRows.Clear();
                productionJournalRows.Clear();
                inspectionJournalRows.Clear();
                RefreshDocumentLibraries();
                AppendChangeLog("Создание объекта", $"Создан новый объект: {currentObject.Name}");
            }
            else
            {
                ApplyObjectMetadata(currentObject, result);
                AppendChangeLog("Настройки объекта", "Изменены параметры объекта.");
            }

            SaveState(SaveTrigger.System);
            RefreshTreePreserveState();
            ApplyProjectUiSettings();
            EnsureProductionJournalStorage();
            RefreshProductionJournalState();
            RefreshSummaryTable();
        }

        private static ProjectObject CloneProjectObjectForWizard(ProjectObject source)
        {
            if (source == null)
            {
                return new ProjectObject
                {
                    BlocksCount = 1,
                    FloorsPerBlock = 1,
                    SameFloorsInBlocks = true
                };
            }

            var json = JsonSerializer.Serialize(source);
            var clone = JsonSerializer.Deserialize<ProjectObject>(json);
            if (clone == null)
            {
                clone = new ProjectObject
                {
                    BlocksCount = 1,
                    FloorsPerBlock = 1,
                    SameFloorsInBlocks = true
                };
            }

            return clone;
        }

        private static void ApplyObjectMetadata(ProjectObject target, ProjectObject source)
        {
            if (target == null || source == null)
                return;

            target.Name = source.Name ?? string.Empty;
            target.FullObjectName = source.FullObjectName ?? string.Empty;
            target.BlocksCount = Math.Max(1, source.BlocksCount);
            target.HasBasement = source.HasBasement;
            target.SameFloorsInBlocks = source.SameFloorsInBlocks;
            target.FloorsPerBlock = Math.Max(1, source.FloorsPerBlock);
            target.FloorsByBlock = source.FloorsByBlock != null
                ? new Dictionary<int, int>(source.FloorsByBlock)
                : new Dictionary<int, int>();
            target.BlockAxesByNumber = source.BlockAxesByNumber != null
                ? new Dictionary<int, string>(source.BlockAxesByNumber)
                : new Dictionary<int, string>();
            target.GeneralContractorRepresentative = source.GeneralContractorRepresentative ?? string.Empty;
            target.TechnicalSupervisorRepresentative = source.TechnicalSupervisorRepresentative ?? string.Empty;
            target.ProjectOrganizationRepresentative = source.ProjectOrganizationRepresentative ?? string.Empty;
            target.ProjectDocumentationName = source.ProjectDocumentationName ?? string.Empty;
            target.MasterNames = source.MasterNames?.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()).ToList() ?? new List<string>();
            target.ForemanNames = source.ForemanNames?.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()).ToList() ?? new List<string>();
            target.ResponsibleForeman = source.ResponsibleForeman ?? string.Empty;
            target.SiteManagerName = source.SiteManagerName ?? string.Empty;
        }

        private void CreateObject_Click(object sender, RoutedEventArgs e) => ObjectWizard_Click(sender, e);

        private void ObjectSettings_Click(object sender, RoutedEventArgs e) => ObjectWizard_Click(sender, e);

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
            if (!EnsureCanEditOperation("база материалов"))
                return;

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
            RefreshJournalAnomalies();
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
            RequestArrivalFilterRefresh(immediate: true);
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
            RequestArrivalFilterRefresh();
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
                    RequestSummaryRefresh();
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
            RefreshJournalAnomalies();


            RenderJvk();
            if (ArrivalLegacyGrid != null)
                ArrivalLegacyGrid.ItemsSource = filteredJournal;

            if (initialUiPrepared && IsArrivalTabActive() && arrivalMatrixMode)
                RenderArrivalMatrix();
            else
                RenderArrivalMatrixPlaceholder();
            RequestSummaryRefresh();

        }

        private void RefreshJournalAnomalies()
        {
            if (journal == null || journal.Count == 0)
                return;

            var duplicateMap = journal
                .Where(x => x != null)
                .GroupBy(x => $"{x.Date:yyyy-MM-dd}|{(x.Ttn ?? string.Empty).Trim()}|{(x.MaterialGroup ?? string.Empty).Trim()}|{(x.MaterialName ?? string.Empty).Trim()}|{x.Quantity:0.###}")
                .Where(x => x.Count() > 1)
                .ToDictionary(x => x.Key, x => x.Count(), StringComparer.CurrentCultureIgnoreCase);

            foreach (var row in journal.Where(x => x != null))
            {
                var notes = new List<string>();
                if (string.IsNullOrWhiteSpace(row.MaterialGroup))
                    notes.Add("Пустой тип");
                if (string.IsNullOrWhiteSpace(row.MaterialName))
                    notes.Add("Пустое наименование");
                if (row.Quantity <= 0)
                    notes.Add("Количество <= 0");

                if (string.Equals(row.Category, "Основные", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(row.Ttn))
                        notes.Add("Пустой ТТН");
                    if (string.IsNullOrWhiteSpace(row.Supplier))
                        notes.Add("Пустой поставщик");
                }

                var duplicateKey = $"{row.Date:yyyy-MM-dd}|{(row.Ttn ?? string.Empty).Trim()}|{(row.MaterialGroup ?? string.Empty).Trim()}|{(row.MaterialName ?? string.Empty).Trim()}|{row.Quantity:0.###}";
                if (duplicateMap.ContainsKey(duplicateKey))
                    notes.Add("Вероятный дубль записи");

                row.IsAnomaly = notes.Count > 0;
                row.AnomalyText = notes.Count == 0 ? string.Empty : string.Join("; ", notes);
            }
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
            RequestArrivalFilterRefresh(immediate: true);

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

        private void UpdateAutoSaveTimerInterval()
        {
            int minutes = 5;
            if (currentObject?.UiSettings != null)
                minutes = currentObject.UiSettings.AutoSaveIntervalMinutes;

            if (minutes < 1)
                minutes = 1;
            if (minutes > 240)
                minutes = 240;

            autoSaveTimer.Stop();
            autoSaveTimer.Interval = TimeSpan.FromMinutes(minutes);
            autoSaveTimer.Start();
            UpdateStatusBar();
        }

        private void TryAutoSaveState()
        {
            if (currentObject == null || string.IsNullOrWhiteSpace(currentSaveFileName))
                return;

            try
            {
                var snapshot = BuildCurrentStateSnapshotJson();
                if (string.Equals(snapshot, lastSavedStateSnapshot, StringComparison.Ordinal))
                    return;

                SaveState(SaveTrigger.Auto);
            }
            catch
            {
                // Автосохранение не должно прерывать работу пользователя всплывающими ошибками.
            }
        }

        private void SaveState(SaveTrigger trigger = SaveTrigger.Manual)
        {
            if (string.IsNullOrWhiteSpace(currentSaveFileName))
                throw new InvalidOperationException("Не задан путь файла состояния.");

            if (!TryAcquireProjectLock(currentSaveFileName))
                throw new InvalidOperationException("Файл проекта сейчас заблокирован другим экземпляром.");

            MarkSummaryDataDirty();
            SaveGridColumnPreferencesForAll();
            var snapshotJson = BuildCurrentStateSnapshotJson();
            var json = BuildCurrentStateJson();
            SaveStateJsonTransactional(json, currentSaveFileName);
            lastSavedStateSnapshot = snapshotJson;
            lastSuccessfulSaveLocalTime = DateTime.Now;

            var saveReason = trigger switch
            {
                SaveTrigger.Auto => "Автосохранение выполнено",
                SaveTrigger.System => "Системные изменения сохранены",
                _ => "Данные сохранены"
            };
            SetLastOperationStatus(saveReason);
            AddOperationLogEntry(
                "Сохранение",
                trigger == SaveTrigger.Auto ? "Авто" : trigger == SaveTrigger.System ? "Система" : "Ручное",
                saveReason);

            if (trigger == SaveTrigger.Auto)
            {
                try
                {
                    WriteAutoBackupSnapshot(json);
                }
                catch
                {
                    // Ошибка автобэкапа не должна ломать основное сохранение.
                }
            }

            UpdateStatusBar();
        }

        private void SaveStateJsonTransactional(string json, string targetPath)
        {
            var targetDirectory = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(targetPath));
            if (!string.IsNullOrWhiteSpace(targetDirectory))
                Directory.CreateDirectory(targetDirectory);

            var tempFileName = $"{targetPath}.tmp";
            try
            {
                File.WriteAllText(tempFileName, json);

                var writtenTempJson = File.ReadAllText(tempFileName);
                if (!TryDeserializeAppState(writtenTempJson, out _, out var tempValidationError))
                    throw new InvalidDataException($"Временный файл состояния не прошёл проверку: {tempValidationError}");

                File.Copy(tempFileName, targetPath, overwrite: true);

                var writtenFinalJson = File.ReadAllText(targetPath);
                if (!TryDeserializeAppState(writtenFinalJson, out _, out var finalValidationError))
                    throw new InvalidDataException($"Файл состояния после записи повреждён: {finalValidationError}");
            }
            finally
            {
                try
                {
                    if (File.Exists(tempFileName))
                        File.Delete(tempFileName);
                }
                catch
                {
                    // ignore
                }
            }
        }

        private string BuildCurrentStateJson()
        {
            return JsonSerializer.Serialize(BuildCurrentState(savedAtUtc: DateTime.UtcNow));
        }

        private string BuildCurrentStateSnapshotJson()
        {
            return JsonSerializer.Serialize(BuildCurrentState(savedAtUtc: DateTime.UnixEpoch));
        }

        private AppState BuildCurrentState(DateTime savedAtUtc)
        {
            return new AppState
            {
                SchemaVersion = AppState.LatestSchemaVersion,
                SavedAtUtc = savedAtUtc,
                CurrentObject = currentObject,
                Journal = journal
            };
        }

        private bool TryDeserializeAppState(string json, out AppState state, out string errorMessage)
        {
            state = null;
            errorMessage = string.Empty;

            if (string.IsNullOrWhiteSpace(json))
            {
                errorMessage = "Пустой JSON.";
                return false;
            }

            try
            {
                state = JsonSerializer.Deserialize<AppState>(json);
                if (state == null)
                {
                    errorMessage = "JSON не удалось преобразовать в состояние приложения.";
                    return false;
                }

                ApplyStateMigrations(state);
                return true;
            }
            catch (Exception primaryEx)
            {
                if (!TryDeserializeAppStateWithLegacyFixes(json, out state, out errorMessage))
                {
                    errorMessage = string.IsNullOrWhiteSpace(errorMessage) ? primaryEx.Message : errorMessage;
                    return false;
                }

                return true;
            }
        }

        private bool TryDeserializeAppStateWithLegacyFixes(string json, out AppState state, out string errorMessage)
        {
            state = null;
            errorMessage = string.Empty;

            try
            {
                var rootNode = JsonNode.Parse(json) as JsonObject;
                if (rootNode == null)
                {
                    errorMessage = "Не удалось разобрать JSON как объект.";
                    return false;
                }

                if (!rootNode.ContainsKey(nameof(AppState.SchemaVersion)))
                    rootNode[nameof(AppState.SchemaVersion)] = 1;

                if (!rootNode.ContainsKey(nameof(AppState.SavedAtUtc)))
                    rootNode[nameof(AppState.SavedAtUtc)] = DateTime.UtcNow;

                NormalizeLegacyDateNodes(rootNode);

                var currentObjectNode = rootNode[nameof(AppState.CurrentObject)] as JsonObject;
                var catalogNode = currentObjectNode?["MaterialCatalog"] as JsonArray;
                if (catalogNode != null)
                {
                    foreach (var itemNode in catalogNode.OfType<JsonObject>())
                    {
                        NormalizeCatalogStringField(itemNode, "CategoryName");
                        NormalizeCatalogStringField(itemNode, "TypeName");
                        NormalizeCatalogStringField(itemNode, "SubTypeName");
                        NormalizeCatalogStringField(itemNode, "MaterialName");
                    }
                }

                var normalizedJson = rootNode.ToJsonString();
                state = JsonSerializer.Deserialize<AppState>(normalizedJson);
                if (state == null)
                {
                    errorMessage = "Нормализованный JSON не удалось десериализовать.";
                    return false;
                }

                ApplyStateMigrations(state);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"Не удалось применить миграции старого формата: {ex.Message}";
                return false;
            }
        }

        private static readonly Regex LegacyMicrosoftJsonDateRegex = new(
            @"^/Date\(([-+]?\d+)([+-]\d{4})?\)/$",
            RegexOptions.Compiled);

        private static void NormalizeLegacyDateNodes(JsonNode node, string propertyName = null)
        {
            switch (node)
            {
                case JsonObject obj:
                    foreach (var pair in obj.ToList())
                    {
                        if (pair.Value == null)
                            continue;

                        if (pair.Value is JsonValue valueNode
                            && valueNode.TryGetValue<string>(out var rawString)
                            && TryNormalizeLegacyDateString(
                                rawString,
                                string.Equals(pair.Key, nameof(AppState.SavedAtUtc), StringComparison.Ordinal),
                                out var normalizedDate))
                        {
                            obj[pair.Key] = normalizedDate;
                            continue;
                        }

                        NormalizeLegacyDateNodes(pair.Value, pair.Key);
                    }

                    break;

                case JsonArray array:
                    for (var index = 0; index < array.Count; index++)
                    {
                        if (array[index] != null)
                            NormalizeLegacyDateNodes(array[index], propertyName);
                    }

                    break;
            }
        }

        private static bool TryNormalizeLegacyDateString(string rawValue, bool asUtc, out string normalizedDate)
        {
            normalizedDate = string.Empty;
            if (string.IsNullOrWhiteSpace(rawValue))
                return false;

            var match = LegacyMicrosoftJsonDateRegex.Match(rawValue.Trim());
            if (!match.Success || !long.TryParse(match.Groups[1].Value, out var unixMilliseconds))
                return false;

            var dto = DateTimeOffset.FromUnixTimeMilliseconds(unixMilliseconds);
            normalizedDate = asUtc
                ? dto.UtcDateTime.ToString("O", CultureInfo.InvariantCulture)
                : dto.ToLocalTime().ToString("yyyy-MM-ddTHH:mm:sszzz", CultureInfo.InvariantCulture);
            return true;
        }

        private static void NormalizeCatalogStringField(JsonObject itemNode, string fieldName)
        {
            if (itemNode == null || !itemNode.TryGetPropertyValue(fieldName, out var node) || node == null)
                return;

            if (node is JsonValue)
                return;

            if (node is JsonObject objNode)
            {
                if (objNode.TryGetPropertyValue("Name", out var nameNode) && nameNode is JsonValue)
                {
                    itemNode[fieldName] = nameNode.GetValue<string>();
                    return;
                }

                var firstStringValue = objNode
                    .Where(x => x.Value is JsonValue)
                    .Select(x => x.Value)
                    .FirstOrDefault();
                if (firstStringValue is JsonValue valueNode)
                {
                    itemNode[fieldName] = valueNode.GetValue<string>();
                    return;
                }

                itemNode[fieldName] = objNode.ToJsonString();
                return;
            }

            itemNode[fieldName] = node.ToJsonString();
        }

        private static void ApplyStateMigrations(AppState state)
        {
            AppStateMigration.Apply(state);
        }

        private void WriteAutoBackupSnapshot(string json)
        {
            var backupDir = GetAutoBackupDirectory(currentSaveFileName);
            var backupFile = System.IO.Path.Combine(backupDir, $"autosave_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.json");
            File.WriteAllText(backupFile, json);
            TrimAutoBackups(backupDir, MaxAutoBackupFiles);
        }

        private static void TrimAutoBackups(string backupDir, int maxFiles)
        {
            if (string.IsNullOrWhiteSpace(backupDir) || !Directory.Exists(backupDir) || maxFiles <= 0)
                return;

            var files = new DirectoryInfo(backupDir)
                .GetFiles("autosave_*.json", SearchOption.TopDirectoryOnly)
                .OrderByDescending(x => x.LastWriteTimeUtc)
                .ToList();

            foreach (var file in files.Skip(maxFiles))
            {
                try
                {
                    file.Delete();
                }
                catch
                {
                    // ignore
                }
            }
        }

        private string GetLatestAutoBackupFile(string saveFileName)
        {
            var backupDir = GetAutoBackupDirectory(saveFileName);
            if (!Directory.Exists(backupDir))
                return string.Empty;

            return new DirectoryInfo(backupDir)
                .GetFiles("autosave_*.json", SearchOption.TopDirectoryOnly)
                .OrderByDescending(x => x.LastWriteTimeUtc)
                .Select(x => x.FullName)
                .FirstOrDefault() ?? string.Empty;
        }

        private string CreateSafetyBackupBeforeOperation(string reason)
        {
            if (string.IsNullOrWhiteSpace(currentSaveFileName))
                return string.Empty;

            var state = new AppState
            {
                SchemaVersion = AppState.LatestSchemaVersion,
                SavedAtUtc = DateTime.UtcNow,
                CurrentObject = currentObject,
                Journal = journal
            };

            var stateJson = JsonSerializer.Serialize(state);
            var backupRoot = System.IO.Path.Combine(GetProjectRuntimeDirectory(currentSaveFileName), "safety");
            Directory.CreateDirectory(backupRoot);
            var backupPath = System.IO.Path.Combine(
                backupRoot,
                $"safety_{DateTime.UtcNow:yyyyMMdd_HHmmss}_{SanitizeFileName(reason)}.ccbak");

            using var stream = new FileStream(backupPath, FileMode.Create, FileAccess.Write, FileShare.None);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
            var stateEntry = archive.CreateEntry("state.json", CompressionLevel.Optimal);
            using (var entryStream = stateEntry.Open())
            using (var writer = new StreamWriter(entryStream))
            {
                writer.Write(stateJson);
            }

            _ = ExportProjectStorageToArchive(archive);
            return backupPath;
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "backup";

            var invalid = System.IO.Path.GetInvalidFileNameChars();
            var sanitized = new string(value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
            return string.IsNullOrWhiteSpace(sanitized) ? "backup" : sanitized;
        }

        private void AppendChangeLog(string action, string details = "")
        {
            if (currentObject == null)
                return;

            currentObject.ChangeLog ??= new List<ProjectChangeLogEntry>();
            currentObject.ChangeLog.Add(new ProjectChangeLogEntry
            {
                TimestampUtc = DateTime.UtcNow,
                UserName = Environment.UserName ?? string.Empty,
                Action = action ?? string.Empty,
                Details = details ?? string.Empty
            });

            if (currentObject.ChangeLog.Count > MaxChangeLogEntries)
            {
                var removeCount = currentObject.ChangeLog.Count - MaxChangeLogEntries;
                currentObject.ChangeLog.RemoveRange(0, removeCount);
            }

            if (!string.IsNullOrWhiteSpace(action))
                lastOperationStatusText = string.IsNullOrWhiteSpace(details) ? action.Trim() : $"{action.Trim()}: {details.Trim()}";

            AddOperationLogEntry("Журнал изменений", action, details);
            UpdateStatusBar();
        }

        private static string FormatChangeLogEntry(ProjectChangeLogEntry entry)
        {
            if (entry == null)
                return string.Empty;

            var localTime = entry.TimestampUtc.ToLocalTime();
            var details = string.IsNullOrWhiteSpace(entry.Details) ? string.Empty : $" — {entry.Details}";
            return $"{localTime:dd.MM.yyyy HH:mm:ss} | {entry.UserName} | {entry.Action}{details}";
        }

        private (List<string> issues, List<string> warnings) ValidateDataIntegrity()
        {
            var issues = new List<string>();
            var warnings = new List<string>();

            if (currentObject == null)
            {
                issues.Add("Объект не создан.");
                return (issues, warnings);
            }

            if (journal == null)
                issues.Add("Журнал прихода не инициализирован.");

            currentObject.Demand ??= new Dictionary<string, MaterialDemand>(StringComparer.CurrentCultureIgnoreCase);
            currentObject.ChangeLog ??= new List<ProjectChangeLogEntry>();

            var arrivalKeys = (journal ?? new List<JournalRecord>())
                .Where(x => !string.IsNullOrWhiteSpace(x.MaterialGroup) && !string.IsNullOrWhiteSpace(x.MaterialName))
                .Select(x => BuildDemandKey(x.MaterialGroup, x.MaterialName))
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase);

            foreach (var pair in currentObject.Demand)
            {
                var demandKey = pair.Key ?? string.Empty;
                if (string.IsNullOrWhiteSpace(demandKey))
                {
                    issues.Add("Обнаружен пустой ключ в таблице потребности.");
                    continue;
                }

                if (!arrivalKeys.Contains(demandKey))
                    warnings.Add($"Для потребности \"{demandKey}\" нет записей прихода.");

                var demand = pair.Value;
                if (demand?.Levels != null)
                {
                    foreach (var block in demand.Levels)
                    {
                        foreach (var mark in block.Value ?? new Dictionary<string, double>())
                        {
                            if (mark.Value < 0)
                                issues.Add($"Отрицательная потребность: {demandKey}, блок {block.Key}, отметка {mark.Key}.");
                        }
                    }
                }
            }

            foreach (var row in currentObject.ProductionJournal ?? new List<ProductionJournalEntry>())
            {
                if (string.IsNullOrWhiteSpace(row?.WorkName) || string.IsNullOrWhiteSpace(row?.ElementsText))
                    continue;

                var keys = ParseProductionItems(row.ElementsText)
                    .Select(x => BuildDemandKey(row.WorkName, x.MaterialName))
                    .Distinct(StringComparer.CurrentCultureIgnoreCase);

                foreach (var key in keys)
                {
                    if (!currentObject.Demand.ContainsKey(key))
                        warnings.Add($"В ПР используется элемент без потребности в сводке: {key}.");
                }
            }

            var timesheetNames = (currentObject.TimesheetPeople ?? new List<TimesheetPersonEntry>())
                .Select(x => x?.FullName?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase);

            foreach (var row in currentObject.OtJournal ?? new List<OtJournalEntry>())
            {
                var fullName = row?.FullName?.Trim();
                if (string.IsNullOrWhiteSpace(fullName))
                    continue;

                if (!timesheetNames.Contains(fullName))
                    warnings.Add($"В ОТ есть сотрудник \"{fullName}\", которого нет в табеле.");
            }

            void ValidateDocumentNodes(IEnumerable<DocumentTreeNode> nodes, string caption)
            {
                if (nodes == null)
                    return;

                foreach (var node in nodes)
                {
                    if (node == null)
                        continue;

                    if (!node.IsFolder)
                    {
                        var path = ResolveDocumentPath(node);
                        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                            warnings.Add($"{caption}: не найден файл \"{node.Name}\".");
                        else if (!string.IsNullOrWhiteSpace(node.ContentHash))
                        {
                            if (TryComputeFileHash(path, out var hash, out var size))
                            {
                                if (!string.Equals(hash, node.ContentHash, StringComparison.OrdinalIgnoreCase))
                                    warnings.Add($"{caption}: изменилось содержимое файла \"{node.Name}\" (хэш не совпадает).");

                                if (node.FileSizeBytes.HasValue && node.FileSizeBytes.Value != size)
                                    warnings.Add($"{caption}: изменился размер файла \"{node.Name}\".");
                            }
                        }
                    }

                    ValidateDocumentNodes(node.Children, caption);
                }
            }

            ValidateDocumentNodes(currentObject.PdfDocuments, "ПДФ");
            ValidateDocumentNodes(currentObject.EstimateDocuments, "Сметы");

            return (issues, warnings);
        }

        private List<DiagnosticsMetricRow> BuildDiagnosticsMetricRows()
        {
            EnsureProjectUiSettings();

            return new List<DiagnosticsMetricRow>
            {
                new() { Metric = "Объект", Value = string.IsNullOrWhiteSpace(currentObject?.Name) ? "Не создан" : currentObject.Name },
                new() { Metric = "Файл проекта", Value = string.IsNullOrWhiteSpace(currentSaveFileName) ? "—" : currentSaveFileName },
                new() { Metric = "Каталог данных", Value = string.IsNullOrWhiteSpace(currentObject?.UiSettings?.DataRootDirectory) ? "По умолчанию" : currentObject.UiSettings.DataRootDirectory },
                new() { Metric = "Автосохранение", Value = $"{Math.Max(1, currentObject?.UiSettings?.AutoSaveIntervalMinutes ?? 5)} мин" },
                new() { Metric = "Последнее сохранение", Value = lastSuccessfulSaveLocalTime?.ToString("dd.MM.yyyy HH:mm:ss", CultureInfo.CurrentCulture) ?? "Нет данных" },
                new() { Metric = "Роль доступа", Value = GetAccessRoleDisplayName(GetCurrentAccessRole()) },
                new() { Metric = "Блокировка", Value = isLocked ? "Включена" : "Выключена" },
                new() { Metric = "Безопасный старт", Value = IsSafeStartupEnabled() ? "Да" : "Нет" },
                new() { Metric = "Последняя операция", Value = string.IsNullOrWhiteSpace(lastOperationStatusText) ? "Готово" : lastOperationStatusText },
                new() { Metric = "Записей прихода", Value = (journal?.Count ?? 0).ToString(CultureInfo.CurrentCulture) },
                new() { Metric = "Записей ОТ", Value = (currentObject?.OtJournal?.Count ?? 0).ToString(CultureInfo.CurrentCulture) },
                new() { Metric = "Сотрудников табеля", Value = (currentObject?.TimesheetPeople?.Count ?? 0).ToString(CultureInfo.CurrentCulture) },
                new() { Metric = "Записей ПР", Value = (currentObject?.ProductionJournal?.Count ?? 0).ToString(CultureInfo.CurrentCulture) },
                new() { Metric = "Записей осмотров", Value = (currentObject?.InspectionJournal?.Count ?? 0).ToString(CultureInfo.CurrentCulture) },
                new() { Metric = "Проблем хранения", Value = lastStorageIntegrityIssues.Count.ToString(CultureInfo.CurrentCulture) }
            };
        }

        private void OpenDiagnosticsDashboard_Click(object sender, RoutedEventArgs e)
        {
            var (issues, warnings) = ValidateDataIntegrity();
            var metrics = BuildDiagnosticsMetricRows();
            var integrityRows = issues.Select(x => new DiagnosticsMetricRow { Metric = "Ошибка", Value = x })
                .Concat(warnings.Select(x => new DiagnosticsMetricRow { Metric = "Предупреждение", Value = x }))
                .Concat(lastStorageIntegrityIssues.Select(x => new DiagnosticsMetricRow { Metric = "Хранилище", Value = x }))
                .ToList();

            if (integrityRows.Count == 0)
                integrityRows.Add(new DiagnosticsMetricRow { Metric = "Состояние", Value = "Проблем не найдено" });

            var timings = tabOpenDiagnostics.Count == 0
                ? new List<DiagnosticsMetricRow> { new() { Metric = "Открытие вкладок", Value = "Нет данных" } }
                : tabOpenDiagnostics
                    .OrderByDescending(x => x.Value)
                    .Select(x => new DiagnosticsMetricRow { Metric = x.Key, Value = $"{x.Value.TotalMilliseconds:F0} мс" })
                    .ToList();

            var changeLogRows = (currentObject?.ChangeLog ?? new List<ProjectChangeLogEntry>())
                .OrderByDescending(x => x.TimestampUtc)
                .Take(100)
                .Select(x => new DiagnosticsMetricRow
                {
                    Metric = $"{x.TimestampUtc.ToLocalTime():dd.MM HH:mm}",
                    Value = $"{x.Action}{(string.IsNullOrWhiteSpace(x.Details) ? string.Empty : $" — {x.Details}")}"
                })
                .ToList();

            if (changeLogRows.Count == 0)
                changeLogRows.Add(new DiagnosticsMetricRow { Metric = "Журнал", Value = "Пока пуст" });

            var window = new Window
            {
                Title = "Панель диагностики",
                Owner = this,
                Width = 1180,
                Height = 760,
                MinWidth = 980,
                MinHeight = 620,
                Background = (Brush)FindResource("AppBgBrush"),
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            var header = new Border
            {
                Background = (Brush)FindResource("SurfaceBrush"),
                BorderBrush = (Brush)FindResource("StrokeBrush"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(16),
                Margin = new Thickness(0, 0, 0, 12),
                Child = new StackPanel
                {
                    Children =
                    {
                        new TextBlock { Text = "Диагностика проекта", FontSize = 22, FontWeight = FontWeights.SemiBold },
                        new TextBlock
                        {
                            Text = "Сводка состояния, проблемы целостности, время открытия вкладок, журнал изменений и внутренние операции.",
                            Margin = new Thickness(0, 8, 0, 0),
                            Foreground = (Brush)FindResource("TextSecondaryBrush"),
                            TextWrapping = TextWrapping.Wrap
                        }
                    }
                }
            };
            Grid.SetRow(header, 0);
            root.Children.Add(header);

            var tabs = new TabControl();
            Grid.SetRow(tabs, 1);

            tabs.Items.Add(CreateDiagnosticsTab("Состояние", metrics));
            tabs.Items.Add(CreateDiagnosticsTab("Проблемы", integrityRows));
            tabs.Items.Add(CreateDiagnosticsTab("Вкладки", timings));
            tabs.Items.Add(CreateDiagnosticsTab("Изменения", changeLogRows));
            tabs.Items.Add(CreateOperationLogTab());

            root.Children.Add(tabs);
            window.Content = root;
            window.ShowDialog();
        }

        private TabItem CreateDiagnosticsTab(string header, IEnumerable<DiagnosticsMetricRow> rows)
        {
            var grid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                ItemsSource = rows?.ToList()
            };
            grid.Columns.Add(new DataGridTextColumn { Header = "Параметр", Binding = new Binding(nameof(DiagnosticsMetricRow.Metric)), Width = 220 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Значение", Binding = new Binding(nameof(DiagnosticsMetricRow.Value)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });

            return new TabItem
            {
                Header = header,
                Content = new Border
                {
                    Margin = new Thickness(0, 10, 0, 0),
                    Background = (Brush)FindResource("SurfaceBrush"),
                    BorderBrush = (Brush)FindResource("StrokeBrush"),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(10),
                    Padding = new Thickness(8),
                    Child = grid
                }
            };
        }

        private TabItem CreateOperationLogTab()
        {
            var grid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                ItemsSource = operationLogEntries
            };
            grid.Columns.Add(new DataGridTextColumn { Header = "Время", Binding = new Binding(nameof(OperationLogEntry.TimestampLocal)) { StringFormat = "dd.MM.yyyy HH:mm:ss" }, Width = 150 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Тип", Binding = new Binding(nameof(OperationLogEntry.Kind)), Width = 180 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Статус", Binding = new Binding(nameof(OperationLogEntry.Status)), Width = 140 });
            grid.Columns.Add(new DataGridTextColumn { Header = "Подробности", Binding = new Binding(nameof(OperationLogEntry.Details)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });

            return new TabItem
            {
                Header = "Операции",
                Content = new Border
                {
                    Margin = new Thickness(0, 10, 0, 0),
                    Background = (Brush)FindResource("SurfaceBrush"),
                    BorderBrush = (Brush)FindResource("StrokeBrush"),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(10),
                    Padding = new Thickness(8),
                    Child = grid
                }
            };
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
            SaveState(SaveTrigger.System);
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
            currentObject.ChangeLog ??= new List<ProjectChangeLogEntry>();

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

            var json = File.ReadAllText(currentSaveFileName);
            if (!TryDeserializeAppState(json, out var state, out var error))
            {
                MessageBox.Show(
                    $"Не удалось загрузить сохранённые данные. Файл состояния повреждён или имеет неверный формат.{Environment.NewLine}{Environment.NewLine}{error}",
                    "Ошибка загрузки",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            currentObject = state.CurrentObject;
            journal = state.Journal ?? new();
            MarkSummaryDataDirty();
            timesheetInitialized = false;
            productionJournalInitialized = false;
            inspectionJournalInitialized = false;
            productionStateDirty = true;
            inspectionStateDirty = true;
            timesheetRows.Clear();
            productionJournalRows.Clear();
            inspectionJournalRows.Clear();
            tabOpenDiagnostics.Clear();
            UpdateTabOpenDiagnosticsText();
            lastSavedStateSnapshot = BuildCurrentStateSnapshotJson();
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
            RefreshArrivalFilterTemplates();
            RefreshJournalAnomalies();

        }

        private bool HasUnsavedChanges()
        {
            var currentSnapshot = BuildCurrentStateSnapshotJson();
            return !string.Equals(currentSnapshot, lastSavedStateSnapshot, StringComparison.Ordinal);
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (closeConfirmed)
                return;

            CommitOpenEdits();
            SaveGridColumnPreferencesForAll();
            if (pendingGridPreferenceSave)
            {
                pendingGridPreferenceSave = false;
                gridPreferenceSaveDebounceTimer?.Stop();
            }
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
            if (!EnsureCanEditOperation("сохранение"))
                return;

            CommitOpenEdits();
            AppendChangeLog("Сохранение", "Пользователь сохранил проект.");
            SaveState(SaveTrigger.Manual);
            MessageBox.Show("Данные сохранены");
        }

        private void SaveAs_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanEditOperation("сохранение как"))
                return;

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
            if (!TryAcquireProjectLock(currentSaveFileName))
            {
                currentSaveFileName = previousSaveFileName;
                _ = TryAcquireProjectLock(currentSaveFileName, showMessageOnFailure: false);
                return;
            }

            RemoveSessionMarker();
            previousSessionCrashed = false;
            _ = CheckIfPreviousSessionCrashed(currentSaveFileName, out var markerPath);
            WriteSessionMarker(markerPath);

            CopyProjectStorage(previousSaveFileName, currentSaveFileName);
            EnsureDocumentLibraries();
            AppendChangeLog("Сохранить как", $"Проект сохранён как {currentSaveFileName}");
            SaveState(SaveTrigger.Manual);
            MessageBox.Show("Проект сохранён в новый файл.");
        }

        private void CopyProjectStorage(string sourceSaveFileName, string targetSaveFileName)
        {
            var sourceRoot = BuildStorageRootPathForCurrentSettings(sourceSaveFileName, createIfMissing: false);
            if (string.IsNullOrWhiteSpace(sourceRoot) || !Directory.Exists(sourceRoot))
                sourceRoot = BuildLegacyStorageRootPath(sourceSaveFileName);

            var targetRoot = BuildStorageRootPathForCurrentSettings(targetSaveFileName, createIfMissing: true);
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
            if (!EnsureCanEditOperation("настройки интерфейса"))
                return;

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

            var previousSettings = currentObject.UiSettings ?? new ProjectUiSettings();
            var previousDataRoot = NormalizeDataRootPath(previousSettings.DataRootDirectory);
            var updatedSettings = window.ResultSettings ?? new ProjectUiSettings();
            updatedSettings.OtStatusFilter = previousSettings.OtStatusFilter;
            updatedSettings.OtSpecialtyFilter = previousSettings.OtSpecialtyFilter;
            updatedSettings.OtBrigadeFilter = previousSettings.OtBrigadeFilter;
            updatedSettings.TabDisplayModes = previousSettings.TabDisplayModes ?? new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            updatedSettings.GridColumnPreferences = previousSettings.GridColumnPreferences ?? new Dictionary<string, List<GridColumnPreference>>(StringComparer.CurrentCultureIgnoreCase);
            updatedSettings.DataRootDirectory = NormalizeDataRootPath(updatedSettings.DataRootDirectory);
            currentObject.UiSettings = updatedSettings;
            MigrateProjectDataFolders(previousDataRoot, updatedSettings.DataRootDirectory);
            ApplyProjectUiSettings();
            RefreshDocumentLibraries();
            AppendChangeLog("Настройки интерфейса", "Изменены настройки интерфейса проекта.");
            SaveState(SaveTrigger.System);
        }

        private void MigrateProjectDataFolders(string oldDataRootPath, string newDataRootPath)
        {
            if (string.IsNullOrWhiteSpace(currentSaveFileName))
                return;

            var oldRoot = NormalizeDataRootPath(oldDataRootPath);
            var newRoot = NormalizeDataRootPath(newDataRootPath);
            if (string.Equals(oldRoot, newRoot, StringComparison.OrdinalIgnoreCase))
                return;

            var key = BuildProjectInstanceKey(currentSaveFileName);
            var oldStorageRoot = System.IO.Path.Combine(oldRoot, "storage", key);
            var newStorageRoot = System.IO.Path.Combine(newRoot, "storage", key);
            if (!string.Equals(oldStorageRoot, newStorageRoot, StringComparison.OrdinalIgnoreCase))
                CopyDirectorySafe(oldStorageRoot, newStorageRoot);

            var oldRuntimeRoot = System.IO.Path.Combine(oldRoot, "runtime", key);
            var newRuntimeRoot = System.IO.Path.Combine(newRoot, "runtime", key);
            if (!string.Equals(oldRuntimeRoot, newRuntimeRoot, StringComparison.OrdinalIgnoreCase))
                CopyDirectorySafe(oldRuntimeRoot, newRuntimeRoot);
        }

        private ObservableCollection<string> BuildProfessionReferenceOptions()
        {
            EnsureReferenceMappingsStorage();
            var values = new List<string>();

            if (currentObject?.OtInstructionNumbersByProfession != null)
                values.AddRange(currentObject.OtInstructionNumbersByProfession.Keys);

            if (currentObject?.OtJournal != null)
            {
                values.AddRange(currentObject.OtJournal
                    .Select(x => x.Profession)
                    .Where(x => !string.IsNullOrWhiteSpace(x)));
                values.AddRange(currentObject.OtJournal
                    .Select(x => x.Specialty)
                    .Where(x => !string.IsNullOrWhiteSpace(x)));
            }

            if (currentObject?.TimesheetPeople != null)
            {
                values.AddRange(currentObject.TimesheetPeople
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.Specialty))
                    .Select(x => x.Specialty));
            }

            values.AddRange(professions);
            values.AddRange(specialties);

            return new ObservableCollection<string>(values
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase));
        }

        private ObservableCollection<string> BuildProductionTypeReferenceOptions()
        {
            EnsureReferenceMappingsStorage();
            var values = new List<string>();

            if (currentObject?.ProductionDeviationsByType != null)
                values.AddRange(currentObject.ProductionDeviationsByType.Keys);

            values.AddRange(productionTargets);

            if (currentObject?.MaterialCatalog != null)
            {
                values.AddRange(currentObject.MaterialCatalog
                    .Where(x => string.Equals((x.CategoryName ?? string.Empty).Trim(), "Основные", StringComparison.CurrentCultureIgnoreCase))
                    .Select(x => x.TypeName)
                    .Where(x => !string.IsNullOrWhiteSpace(x)));
            }

            if (currentObject?.ProductionJournal != null)
            {
                values.AddRange(currentObject.ProductionJournal
                    .Select(x => x.WorkName)
                    .Where(x => !string.IsNullOrWhiteSpace(x)));
            }

            return new ObservableCollection<string>(values
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase));
        }

        private ObservableCollection<OtInstructionReferenceRow> BuildOtInstructionReferenceRows()
        {
            EnsureReferenceMappingsStorage();
            var rows = new List<OtInstructionReferenceRow>();

            if (currentObject?.OtInstructionNumbersByProfession != null)
            {
                foreach (var pair in currentObject.OtInstructionNumbersByProfession
                    .Where(x => !string.IsNullOrWhiteSpace(x.Key) && !string.IsNullOrWhiteSpace(x.Value))
                    .OrderBy(x => x.Key, StringComparer.CurrentCultureIgnoreCase))
                {
                    rows.Add(new OtInstructionReferenceRow
                    {
                        Profession = pair.Key.Trim(),
                        InstructionNumbers = pair.Value.Trim()
                    });
                }
            }

            if (rows.Count == 0 && currentObject?.OtJournal != null)
            {
                var seen = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
                foreach (var row in currentObject.OtJournal
                    .Where(x => !string.IsNullOrWhiteSpace(x.InstructionNumbers))
                    .OrderByDescending(x => x.InstructionDate))
                {
                    var profession = string.IsNullOrWhiteSpace(row.Profession) ? row.Specialty?.Trim() : row.Profession.Trim();
                    if (string.IsNullOrWhiteSpace(profession) || !seen.Add(profession))
                        continue;

                    rows.Add(new OtInstructionReferenceRow
                    {
                        Profession = profession,
                        InstructionNumbers = row.InstructionNumbers.Trim()
                    });
                }
            }

            if (rows.Count == 0)
                rows.Add(new OtInstructionReferenceRow());

            return new ObservableCollection<OtInstructionReferenceRow>(rows);
        }

        private ObservableCollection<ProductionDeviationReferenceRow> BuildProductionDeviationReferenceRows()
        {
            EnsureReferenceMappingsStorage();
            var rows = new List<ProductionDeviationReferenceRow>();
            var pairSet = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

            if (currentObject?.ProductionDeviationsByType != null)
            {
                foreach (var pair in currentObject.ProductionDeviationsByType)
                {
                    var type = pair.Key?.Trim();
                    if (string.IsNullOrWhiteSpace(type))
                        continue;

                    foreach (var deviation in pair.Value ?? new List<string>())
                    {
                        var normalizedDeviation = deviation?.Trim();
                        if (string.IsNullOrWhiteSpace(normalizedDeviation))
                            continue;

                        var key = $"{type}|||{normalizedDeviation}";
                        if (!pairSet.Add(key))
                            continue;

                        rows.Add(new ProductionDeviationReferenceRow
                        {
                            MaterialType = type,
                            Deviation = normalizedDeviation
                        });
                    }
                }
            }

            if (rows.Count == 0 && currentObject?.ProductionJournal != null)
            {
                foreach (var row in currentObject.ProductionJournal
                    .Where(x => !string.IsNullOrWhiteSpace(x.WorkName) && !string.IsNullOrWhiteSpace(x.Deviations))
                    .OrderByDescending(x => x.Date))
                {
                    var type = row.WorkName.Trim();
                    var deviation = row.Deviations.Trim();
                    var key = $"{type}|||{deviation}";
                    if (!pairSet.Add(key))
                        continue;

                    rows.Add(new ProductionDeviationReferenceRow
                    {
                        MaterialType = type,
                        Deviation = deviation
                    });
                }
            }

            if (rows.Count == 0)
                rows.Add(new ProductionDeviationReferenceRow());

            return new ObservableCollection<ProductionDeviationReferenceRow>(rows);
        }

        private void ReferenceCatalogs_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanEditOperation("редактирование справочников"))
                return;

            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            CommitOpenEdits();
            EnsureOtJournalStorage();
            EnsureProductionJournalStorage();
            EnsureReferenceMappingsStorage();

            var professionOptions = BuildProfessionReferenceOptions();
            var materialTypeOptions = BuildProductionTypeReferenceOptions();
            var otRows = BuildOtInstructionReferenceRows();
            var prRows = BuildProductionDeviationReferenceRows();

            var window = new Window
            {
                Title = "Справочники",
                Owner = this,
                Width = 980,
                Height = 700,
                MinWidth = 860,
                MinHeight = 620,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = TryFindResource("AppBgBrush") as Brush ?? Brushes.WhiteSmoke
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            window.Content = root;

            root.Children.Add(new TextBlock
            {
                Text = "Справочник используется для автоподстановки в журналах ОТ и ПР.",
                Margin = new Thickness(0, 0, 0, 10),
                Foreground = TryFindResource("TextSecondaryBrush") as Brush ?? new SolidColorBrush(Color.FromRgb(107, 114, 128))
            });

            var tabs = new TabControl();
            Grid.SetRow(tabs, 1);
            root.Children.Add(tabs);

            var otTab = new TabItem { Header = "ОТ: профессии и инструкции" };
            var otHost = new Grid { Margin = new Thickness(10) };
            otHost.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            otHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            otTab.Content = otHost;

            var otButtons = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            var addOtButton = new Button
            {
                Content = "Добавить строку",
                Style = TryFindResource("PrimaryButton") as Style
            };
            var deleteOtButton = new Button
            {
                Content = "Удалить строку",
                Style = TryFindResource("SecondaryButton") as Style,
                Margin = new Thickness(8, 0, 0, 0)
            };
            otButtons.Children.Add(addOtButton);
            otButtons.Children.Add(deleteOtButton);
            otHost.Children.Add(otButtons);

            var otGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                ItemsSource = otRows,
                ColumnWidth = DataGridLength.SizeToHeader
            };
            Grid.SetRow(otGrid, 1);
            DataGridSizingHelper.SetEnableSmartSizing(otGrid, true);

            var professionCellFactory = new FrameworkElementFactory(typeof(TextBlock));
            professionCellFactory.SetBinding(TextBlock.TextProperty, new Binding(nameof(OtInstructionReferenceRow.Profession)));
            var professionEditFactory = new FrameworkElementFactory(typeof(ComboBox));
            professionEditFactory.SetValue(ComboBox.IsEditableProperty, true);
            professionEditFactory.SetValue(ComboBox.IsTextSearchEnabledProperty, true);
            professionEditFactory.SetValue(ComboBox.StaysOpenOnEditProperty, true);
            professionEditFactory.SetValue(ComboBox.ItemsSourceProperty, professionOptions);
            professionEditFactory.SetBinding(ComboBox.TextProperty, new Binding(nameof(OtInstructionReferenceRow.Profession))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });

            otGrid.Columns.Add(new DataGridTemplateColumn
            {
                Header = "Профессия",
                Width = 280,
                CellTemplate = new DataTemplate { VisualTree = professionCellFactory },
                CellEditingTemplate = new DataTemplate { VisualTree = professionEditFactory }
            });
            otGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Номера инструкций",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(OtInstructionReferenceRow.InstructionNumbers))
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                }
            });
            otHost.Children.Add(otGrid);

            addOtButton.Click += (_, _) => otRows.Add(new OtInstructionReferenceRow());
            deleteOtButton.Click += (_, _) =>
            {
                if (otGrid.SelectedItem is OtInstructionReferenceRow selected)
                    otRows.Remove(selected);
            };

            var prTab = new TabItem { Header = "ПР: тип и отклонения" };
            var prHost = new Grid { Margin = new Thickness(10) };
            prHost.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            prHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            prTab.Content = prHost;

            var prButtons = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            var addPrButton = new Button
            {
                Content = "Добавить строку",
                Style = TryFindResource("PrimaryButton") as Style
            };
            var deletePrButton = new Button
            {
                Content = "Удалить строку",
                Style = TryFindResource("SecondaryButton") as Style,
                Margin = new Thickness(8, 0, 0, 0)
            };
            prButtons.Children.Add(addPrButton);
            prButtons.Children.Add(deletePrButton);
            prHost.Children.Add(prButtons);

            var prGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                ItemsSource = prRows,
                ColumnWidth = DataGridLength.SizeToHeader
            };
            Grid.SetRow(prGrid, 1);
            DataGridSizingHelper.SetEnableSmartSizing(prGrid, true);

            var typeCellFactory = new FrameworkElementFactory(typeof(TextBlock));
            typeCellFactory.SetBinding(TextBlock.TextProperty, new Binding(nameof(ProductionDeviationReferenceRow.MaterialType)));
            var typeEditFactory = new FrameworkElementFactory(typeof(ComboBox));
            typeEditFactory.SetValue(ComboBox.IsEditableProperty, true);
            typeEditFactory.SetValue(ComboBox.IsTextSearchEnabledProperty, true);
            typeEditFactory.SetValue(ComboBox.StaysOpenOnEditProperty, true);
            typeEditFactory.SetValue(ComboBox.ItemsSourceProperty, materialTypeOptions);
            typeEditFactory.SetBinding(ComboBox.TextProperty, new Binding(nameof(ProductionDeviationReferenceRow.MaterialType))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });

            prGrid.Columns.Add(new DataGridTemplateColumn
            {
                Header = "Тип материала",
                Width = 280,
                CellTemplate = new DataTemplate { VisualTree = typeCellFactory },
                CellEditingTemplate = new DataTemplate { VisualTree = typeEditFactory }
            });
            prGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Отклонение",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding(nameof(ProductionDeviationReferenceRow.Deviation))
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                }
            });
            prHost.Children.Add(prGrid);

            addPrButton.Click += (_, _) => prRows.Add(new ProductionDeviationReferenceRow());
            deletePrButton.Click += (_, _) =>
            {
                if (prGrid.SelectedItem is ProductionDeviationReferenceRow selected)
                    prRows.Remove(selected);
            };

            tabs.Items.Add(otTab);
            tabs.Items.Add(prTab);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            var cancelButton = new Button
            {
                Content = "Отмена",
                MinWidth = 120,
                IsCancel = true,
                Style = TryFindResource("SecondaryButton") as Style
            };
            var saveButton = new Button
            {
                Content = "Сохранить",
                MinWidth = 140,
                Margin = new Thickness(8, 0, 0, 0),
                Style = TryFindResource("PrimaryButton") as Style
            };
            footer.Children.Add(cancelButton);
            footer.Children.Add(saveButton);
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            saveButton.Click += (_, _) =>
            {
                var otMap = new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
                foreach (var row in otRows)
                {
                    var profession = row?.Profession?.Trim();
                    var numbers = row?.InstructionNumbers?.Trim();
                    if (string.IsNullOrWhiteSpace(profession) || string.IsNullOrWhiteSpace(numbers))
                        continue;

                    otMap[profession] = numbers;
                }

                var productionMap = new Dictionary<string, List<string>>(StringComparer.CurrentCultureIgnoreCase);
                foreach (var row in prRows)
                {
                    var materialType = row?.MaterialType?.Trim();
                    var deviation = row?.Deviation?.Trim();
                    if (string.IsNullOrWhiteSpace(materialType) || string.IsNullOrWhiteSpace(deviation))
                        continue;

                    if (!productionMap.TryGetValue(materialType, out var list))
                    {
                        list = new List<string>();
                        productionMap[materialType] = list;
                    }

                    if (!list.Contains(deviation, StringComparer.CurrentCultureIgnoreCase))
                        list.Add(deviation);
                }

                currentObject.OtInstructionNumbersByProfession = otMap.ToDictionary(x => x.Key, x => x.Value);
                currentObject.ProductionDeviationsByType = productionMap.ToDictionary(
                    x => x.Key,
                    x => x.Value
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Select(v => v.Trim())
                        .Distinct(StringComparer.CurrentCultureIgnoreCase)
                        .ToList());
                EnsureReferenceMappingsStorage();

                if (currentObject.OtJournal != null)
                {
                    foreach (var row in currentObject.OtJournal.Where(x => string.IsNullOrWhiteSpace(x.InstructionNumbers)))
                        FillInstructionNumbersFromTemplate(row);
                }

                RefreshSpecialties();
                RefreshProfessions();
                RefreshOtFilterOptions();
                productionLookupsDirty = true;
                RefreshProductionJournalLookups(force: true);
                RefreshProductionDeviationOptions();
                SaveState(SaveTrigger.System);
                AppendChangeLog("Справочники", "Обновлены справочники ОТ и ПР.");
                window.DialogResult = true;
            };

            window.ShowDialog();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            SaveState();
            LoadState();
            InitializeOtJournal();
            if (currentObject != null)
                ArrivalPanel.SetObject(currentObject, journal);
            RefreshTreePreserveState();
            RefreshSummaryTable();
            RefreshArrivalTypes();
            RefreshArrivalNames();
            RefreshArrivalFilterTemplates();
            EnsureSelectedTabInitialized();
            RequestArrivalFilterRefresh(immediate: true);
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
            var stateJson = JsonSerializer.Serialize(state);
            var exportedStorageFiles = 0;

            using (var stream = new FileStream(dlg.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                var stateEntry = archive.CreateEntry("state.json", CompressionLevel.Optimal);
                using (var entryStream = stateEntry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write(stateJson);
                }

                exportedStorageFiles = ExportProjectStorageToArchive(archive);
            }

            AppendChangeLog("Экспорт", $"Экспортированы данные проекта в {dlg.FileName}");
            SaveState(SaveTrigger.System);
            MessageBox.Show($"Резервная копия сохранена. В архив добавлено файлов ПДФ/смет: {exportedStorageFiles}.");
        }

        private void ImportAllData_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanRunCriticalOperation("импорт резервной копии"))
                return;

            CommitOpenEdits();

            var dlg = new OpenFileDialog
            {
                Filter = "ConstructionControl backup (*.ccbak)|*.ccbak|JSON (*.json)|*.json"
            };
            if (dlg.ShowDialog() != true)
                return;

            try
            {
                _ = CreateSafetyBackupBeforeOperation("before_import_all");
            }
            catch (Exception backupEx)
            {
                MessageBox.Show(
                    $"Не удалось создать предохранительный бэкап перед импортом.{Environment.NewLine}{backupEx.Message}",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }

            if (!TryLoadBackupState(dlg.FileName, out var state, out var restoredStorageFiles, out var importError) || state == null)
            {
                MessageBox.Show(
                    string.IsNullOrWhiteSpace(importError)
                        ? "Не удалось импортировать резервную копию."
                        : importError,
                    "Ошибка импорта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            PushUndo();
            RestoreState(state);
            RebuildArchiveFromCurrentData();
            AppendChangeLog("Импорт", $"Импортированы данные из {dlg.FileName}");
            SaveState(SaveTrigger.System);
              RefreshTreePreserveState();
              RefreshArrivalTypes();
              RefreshArrivalNames();
              RefreshDocumentLibraries();
              ApplyAllFilters();

              var integrityWarning = lastStorageIntegrityIssues.Count > 0
                  ? $"{Environment.NewLine}{Environment.NewLine}Проверка целостности: найдены проблемы ({lastStorageIntegrityIssues.Count})."
                  : string.Empty;
              MessageBox.Show($"Импорт завершен. Восстановлено файлов ПДФ/смет: {restoredStorageFiles}.{integrityWarning}", "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
          }

        private int ExportProjectStorageToArchive(ZipArchive archive)
        {
            var storageRoot = GetProjectStorageRoot(createIfMissing: false);
            if (string.IsNullOrWhiteSpace(storageRoot) || !Directory.Exists(storageRoot))
                return 0;

            var exported = 0;
            var manifest = new List<DocumentStorageManifestEntry>();
            foreach (var sourceFile in Directory.EnumerateFiles(storageRoot, "*", SearchOption.AllDirectories))
            {
                var relativePath = System.IO.Path.GetRelativePath(storageRoot, sourceFile).Replace('\\', '/');
                var entry = archive.CreateEntry($"storage/{relativePath}", CompressionLevel.Optimal);
                using var input = new FileStream(sourceFile, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var output = entry.Open();
                input.CopyTo(output);
                exported++;

                if (TryComputeFileHash(sourceFile, out var hash, out var size))
                {
                    manifest.Add(new DocumentStorageManifestEntry
                    {
                        RelativePath = relativePath,
                        Hash = hash,
                        Size = size
                    });
                }
            }

            if (manifest.Count > 0)
            {
                var manifestEntry = archive.CreateEntry("storage_manifest.json", CompressionLevel.Optimal);
                using var manifestStream = manifestEntry.Open();
                using var writer = new StreamWriter(manifestStream);
                writer.Write(JsonSerializer.Serialize(manifest));
            }

            return exported;
        }

        private bool TryLoadBackupState(string backupPath, out AppState? state, out int restoredStorageFiles, out string errorMessage)
        {
            state = null;
            restoredStorageFiles = 0;
            errorMessage = string.Empty;

            if (string.IsNullOrWhiteSpace(backupPath) || !File.Exists(backupPath))
            {
                errorMessage = "Файл резервной копии не найден.";
                return false;
            }

            var extension = System.IO.Path.GetExtension(backupPath)?.ToLowerInvariant() ?? string.Empty;
            if (extension == ".ccbak")
            {
                try
                {
                    using var stream = new FileStream(backupPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

                    var stateEntry = archive.GetEntry("state.json")
                        ?? archive.Entries.FirstOrDefault(entry => entry.FullName.EndsWith(".json", StringComparison.OrdinalIgnoreCase));

                    if (stateEntry == null)
                    {
                        errorMessage = "В резервной копии не найден файл состояния state.json.";
                        return false;
                    }

                    string json;
                    using (var stateStream = stateEntry.Open())
                    using (var reader = new StreamReader(stateStream))
                    {
                        json = reader.ReadToEnd();
                    }

                    if (!TryDeserializeAppState(json, out state, out var parseError))
                    {
                        errorMessage = string.IsNullOrWhiteSpace(parseError)
                            ? "Файл состояния в резервной копии имеет неверный формат."
                            : parseError;
                        return false;
                    }

                    restoredStorageFiles = ImportProjectStorageFromArchive(archive);
                    return true;
                }
                catch (InvalidDataException)
                {
                    // Старый формат .ccbak был обычным JSON-файлом.
                }
                catch (Exception ex)
                {
                    errorMessage = $"Ошибка чтения резервной копии: {ex.Message}";
                    return false;
                }
            }

            try
            {
                var rawJson = File.ReadAllText(backupPath);
                if (!TryDeserializeAppState(rawJson, out state, out var parseError))
                {
                    errorMessage = string.IsNullOrWhiteSpace(parseError)
                        ? "Файл резервной копии имеет неверный формат."
                        : parseError;
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"Ошибка чтения резервной копии: {ex.Message}";
                return false;
            }
        }

        private int ImportProjectStorageFromArchive(ZipArchive archive)
        {
            lastStorageIntegrityIssues.Clear();
            var storageEntries = archive.Entries
                .Where(entry => !string.IsNullOrWhiteSpace(entry.Name)
                    && entry.FullName.StartsWith("storage/", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (storageEntries.Count == 0)
                return 0;

            var storageRoot = GetProjectStorageRoot(createIfMissing: true);
            if (string.IsNullOrWhiteSpace(storageRoot))
                return 0;

            var fullRoot = System.IO.Path.GetFullPath(storageRoot);
            if (Directory.Exists(fullRoot))
            {
                foreach (var directory in Directory.EnumerateDirectories(fullRoot))
                    Directory.Delete(directory, recursive: true);
                foreach (var file in Directory.EnumerateFiles(fullRoot))
                    File.Delete(file);
            }
            else
            {
                Directory.CreateDirectory(fullRoot);
            }

            var restored = 0;
            foreach (var entry in storageEntries)
            {
                var relative = entry.FullName.Substring("storage/".Length).TrimStart('/', '\\');
                if (string.IsNullOrWhiteSpace(relative))
                    continue;

                var normalizedRelative = relative.Replace('/', System.IO.Path.DirectorySeparatorChar);
                var targetPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(fullRoot, normalizedRelative));
                var fullRootPrefix = fullRoot.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                    ? fullRoot
                    : fullRoot + System.IO.Path.DirectorySeparatorChar;

                if (!targetPath.StartsWith(fullRootPrefix, StringComparison.OrdinalIgnoreCase))
                    continue;

                var targetDirectory = System.IO.Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrWhiteSpace(targetDirectory))
                    Directory.CreateDirectory(targetDirectory);

                using var input = entry.Open();
                using var output = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None);
                input.CopyTo(output);
                restored++;
            }

            ValidateImportedStorageManifest(archive, fullRoot);

            return restored;
        }

        private void ValidateImportedStorageManifest(ZipArchive archive, string storageRoot)
        {
            if (archive == null || string.IsNullOrWhiteSpace(storageRoot) || !Directory.Exists(storageRoot))
                return;

            var manifestEntry = archive.GetEntry("storage_manifest.json");
            if (manifestEntry == null)
                return;

            try
            {
                using var stream = manifestEntry.Open();
                using var reader = new StreamReader(stream);
                var json = reader.ReadToEnd();
                var entries = JsonSerializer.Deserialize<List<DocumentStorageManifestEntry>>(json) ?? new List<DocumentStorageManifestEntry>();
                foreach (var item in entries)
                {
                    if (item == null || string.IsNullOrWhiteSpace(item.RelativePath))
                        continue;

                    var candidate = item.RelativePath.Replace('/', System.IO.Path.DirectorySeparatorChar).TrimStart(System.IO.Path.DirectorySeparatorChar);
                    var absolutePath = System.IO.Path.GetFullPath(System.IO.Path.Combine(storageRoot, candidate));
                    if (!File.Exists(absolutePath))
                    {
                        lastStorageIntegrityIssues.Add($"Отсутствует файл из манифеста: {item.RelativePath}");
                        continue;
                    }

                    if (!TryComputeFileHash(absolutePath, out var hash, out var size))
                    {
                        lastStorageIntegrityIssues.Add($"Не удалось проверить файл: {item.RelativePath}");
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(item.Hash)
                        && !string.Equals(item.Hash, hash, StringComparison.OrdinalIgnoreCase))
                    {
                        lastStorageIntegrityIssues.Add($"Хэш не совпадает: {item.RelativePath}");
                    }

                    if (item.Size > 0 && item.Size != size)
                        lastStorageIntegrityIssues.Add($"Размер не совпадает: {item.RelativePath}");
                }
            }
            catch (Exception ex)
            {
                lastStorageIntegrityIssues.Add($"Ошибка чтения манифеста целостности: {ex.Message}");
            }
        }

        private void LockToggle_Checked(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            LockButton_Checked(sender, e);
            UpdateStatusBar();
        }

        private void LockToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            LockButton_Unchecked(sender, e);
            UpdateStatusBar();
        }



        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            CommitOpenEdits();
            SaveState(SaveTrigger.Manual);
            Close();
        }

        private void ClearObject_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureCanRunCriticalOperation("очистка объекта", requireCode: false))
                return;

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

            try
            {
                _ = CreateSafetyBackupBeforeOperation("before_clear_object");
            }
            catch (Exception backupEx)
            {
                MessageBox.Show(
                    $"Не удалось создать предохранительный бэкап перед очисткой.{Environment.NewLine}{backupEx.Message}",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }

            PushUndo();
            ClearCurrentObjectData();
            AppendChangeLog("Очистка объекта", "Выполнена полная очистка объекта.");
            SaveState(SaveTrigger.System);
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
            currentObject.ChangeLog = new List<ProjectChangeLogEntry>();

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
            if (!EnsureCanEditOperation("импорт Excel"))
                return;

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
            AppendChangeLog("Импорт Excel", $"Импортировано строк: {importWindow.ImportedRecords?.Count ?? 0}");
            SaveState(SaveTrigger.System);
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
                AppendChangeLog("Архив", "Внесены изменения через окно архива.");
                SaveState(SaveTrigger.System);
                RefreshTreePreserveState();
                ApplyAllFilters();
                RefreshSummaryTable();
                ArrivalPanel.SetObject(currentObject, journal);
            }
        }





        public void RefreshSummaryTable()
        {
            RequestSummaryRefresh(immediate: true);
        }

        private async Task RefreshSummaryTableAsync(int requestId)
        {
            if (SummaryPanel == null)
                return;

            if (currentObject == null)
            {
                SummaryPanel.Items.Clear();
                return;
            }

            var mainRecords = journal
                .Where(j => string.Equals(j.Category, "Основные", StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            var journalGroups = mainRecords
                .Select(j => j.MaterialGroup?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase);

            var groupOrder = currentObject.MaterialGroups
                .Select(g => g.Name?.Trim())
                .Where(name => !string.IsNullOrWhiteSpace(name) && journalGroups.Contains(name))
                .ToList();

            if (groupOrder.Count == 0)
                groupOrder = journalGroups.OrderBy(g => g, StringComparer.CurrentCultureIgnoreCase).ToList();

            if (currentObject.SummaryVisibleGroups == null)
                currentObject.SummaryVisibleGroups = new List<string>();

            var normalizedVisibleGroups = currentObject.SummaryVisibleGroups
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Where(x => groupOrder.Contains(x, StringComparer.CurrentCultureIgnoreCase))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .Take(1)
                .ToList();

            if (normalizedVisibleGroups.Count == 0 && groupOrder.Count > 0)
                normalizedVisibleGroups.Add(groupOrder[0]);

            currentObject.SummaryVisibleGroups = normalizedVisibleGroups;
            RenderSummaryFilters(groupOrder);
            var visibleGroups = currentObject.SummaryVisibleGroups
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Where(x => groupOrder.Contains(x, StringComparer.CurrentCultureIgnoreCase))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var cacheKey = BuildSummaryMatrixCacheKey(visibleGroups);
            if (!summaryMatrixCache.TryGetValue(cacheKey, out var cacheEntry))
            {
                var catalogSnapshot = (currentObject.MaterialCatalog ?? new List<MaterialCatalogItem>())
                    .Where(x => x != null)
                    .Select(x => new MaterialCatalogItem
                    {
                        CategoryName = x.CategoryName ?? string.Empty,
                        TypeName = x.TypeName ?? string.Empty,
                        SubTypeName = x.SubTypeName ?? string.Empty,
                        MaterialName = x.MaterialName ?? string.Empty
                    })
                    .ToList();

                var namesByGroupSnapshot = (currentObject.MaterialNamesByGroup ?? new Dictionary<string, List<string>>())
                    .ToDictionary(
                        pair => pair.Key ?? string.Empty,
                        pair => pair.Value?.Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList() ?? new List<string>(),
                        StringComparer.CurrentCultureIgnoreCase);

                using var scope = BeginProcessingScope("Идет обработка сводки...");
                cacheEntry = await Task.Run(() => BuildSummaryMatrixCacheEntry(
                    visibleGroups,
                    mainRecords,
                    catalogSnapshot,
                    namesByGroupSnapshot,
                    summarySelectedSubType));

                if (requestId != summaryRefreshRequestVersion)
                    return;

                summaryMatrixCache[cacheKey] = cacheEntry;
                while (summaryMatrixCache.Count > MaxSummaryMatrixCacheEntries)
                {
                    var oldestKey = summaryMatrixCache.Keys.FirstOrDefault();
                    if (string.IsNullOrWhiteSpace(oldestKey))
                        break;
                    summaryMatrixCache.Remove(oldestKey);
                }
            }

            if (requestId != summaryRefreshRequestVersion)
                return;

            SummaryPanel.Items.Clear();
            RenderSummaryHeader();

            foreach (var groupData in cacheEntry.Groups)
            {
                RenderMaterialGroup(groupData.GroupName);
                foreach (var row in groupData.Rows)
                {
                    RenderMaterialRow(groupData.GroupName, row.MaterialName, row.Unit, row.TotalArrival, row.Position);
                }
            }

            RenderSummaryFooter();
        }

        private SummaryMatrixCacheEntry BuildSummaryMatrixCacheEntry(
            IEnumerable<string> visibleGroups,
            List<JournalRecord> mainRecords,
            List<MaterialCatalogItem> catalogSnapshot,
            Dictionary<string, List<string>> namesByGroupSnapshot,
            string selectedSubType)
        {
            var cacheEntry = new SummaryMatrixCacheEntry();
            var recordsByDemandKey = mainRecords
                .Where(x => !string.IsNullOrWhiteSpace(x.MaterialGroup) && !string.IsNullOrWhiteSpace(x.MaterialName))
                .GroupBy(x => BuildDemandKey(x.MaterialGroup.Trim(), x.MaterialName.Trim()), StringComparer.CurrentCultureIgnoreCase)
                .ToDictionary(x => x.Key, x => x.ToList(), StringComparer.CurrentCultureIgnoreCase);

            foreach (var group in visibleGroups.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                var rows = new List<SummaryMatrixRowData>();
                var materials = GetMaterialsForGroupFromSnapshot(group, selectedSubType, catalogSnapshot, namesByGroupSnapshot, mainRecords);
                foreach (var material in materials)
                {
                    var demandKey = BuildDemandKey(group, material);
                    if (!recordsByDemandKey.TryGetValue(demandKey, out var records))
                        records = new List<JournalRecord>();

                    var unit = records
                        .Select(r => r.Unit)
                        .FirstOrDefault(u => !string.IsNullOrWhiteSpace(u)) ?? string.Empty;
                    var position = records
                        .Select(r => r.Position)
                        .FirstOrDefault(p => !string.IsNullOrWhiteSpace(p)) ?? string.Empty;
                    var totalArrival = records.Sum(r => r.Quantity);

                    rows.Add(new SummaryMatrixRowData
                    {
                        MaterialName = material,
                        Unit = unit,
                        Position = position,
                        TotalArrival = totalArrival
                    });
                }

                cacheEntry.Groups.Add(new SummaryMatrixGroupData
                {
                    GroupName = group,
                    Rows = rows
                });
            }

            return cacheEntry;
        }

        private static List<string> GetMaterialsForGroupFromSnapshot(
            string group,
            string selectedSubType,
            List<MaterialCatalogItem> catalogSnapshot,
            Dictionary<string, List<string>> namesByGroupSnapshot,
            List<JournalRecord> mainRecords)
        {
            var normalizedSubType = string.IsNullOrWhiteSpace(selectedSubType) ? "Все" : selectedSubType.Trim();
            var catalogMaterials = (catalogSnapshot ?? new List<MaterialCatalogItem>())
                .Where(x => string.Equals(x.CategoryName, "Основные", StringComparison.CurrentCultureIgnoreCase)
                         && string.Equals(x.TypeName ?? string.Empty, group ?? string.Empty, StringComparison.CurrentCultureIgnoreCase)
                         && (string.Equals(normalizedSubType, "Все", StringComparison.CurrentCultureIgnoreCase)
                             || string.Equals(x.SubTypeName ?? string.Empty, normalizedSubType, StringComparison.CurrentCultureIgnoreCase)))
                .Select(x => x.MaterialName?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (catalogMaterials.Count > 0)
                return catalogMaterials;

            if (namesByGroupSnapshot != null
                && namesByGroupSnapshot.TryGetValue(group ?? string.Empty, out var names)
                && names?.Count > 0)
            {
                return names
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
            }

            return (mainRecords ?? new List<JournalRecord>())
                .Where(j => string.Equals(j.Category, "Основные", StringComparison.CurrentCultureIgnoreCase)
                         && string.Equals(j.MaterialGroup ?? string.Empty, group ?? string.Empty, StringComparison.CurrentCultureIgnoreCase))
                .Select(j => j.MaterialName?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
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
            AddCell(summaryGrid, summaryRowIndex, 1, mat, noWrap: false, minWidth: 280);

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
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(0.52, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(0.48, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var title = new TextBlock
            {
                Text = "Выберите типы, блоки, отметки и материалы для дозаказа.",
                FontWeight = FontWeights.SemiBold,
                FontSize = 15,
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(title, 0);
            Grid.SetColumnSpan(title, 4);
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
            Grid.SetRow(previewGrid, 2);
            Grid.SetColumn(previewGrid, 0);
            Grid.SetColumnSpan(previewGrid, 4);
            previewGrid.Margin = new Thickness(0, 12, 0, 0);
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
            Grid.SetRow(footer, 3);
            Grid.SetColumnSpan(footer, 4);
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

            var orderedRows = (rows ?? new List<SummaryReorderPreviewRow>())
                .Where(x => x != null && x.Quantity > 0)
                .OrderBy(x => x.Group, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.Material, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.Block)
                .ThenBy(x => x.Mark, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("Подробно по блокам и отметкам"))));

            var detailedRows = orderedRows
                .Select(x => (IReadOnlyList<string>)new[]
                {
                    x.Group ?? string.Empty,
                    x.Material ?? string.Empty,
                    $"Блок {x.Block}",
                    x.Mark ?? string.Empty,
                    FormatNumberByUnit(x.Quantity, x.Unit),
                    x.Unit ?? string.Empty
                })
                .ToList();

            body.Append(BuildWordTable(
                new[] { "Тип", "Наименование", "Блок", "Отметка", "Количество", "Ед." },
                detailedRows));

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text(string.Empty))));

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("Итого к дозаказу по одинаковым наименованиям"))));

            var aggregatedRows = orderedRows
                .GroupBy(
                    x => $"{(x.Group ?? string.Empty).Trim().ToUpperInvariant()}||{(x.Material ?? string.Empty).Trim().ToUpperInvariant()}||{(x.Unit ?? string.Empty).Trim().ToUpperInvariant()}",
                    StringComparer.Ordinal)
                .Select(g =>
                {
                    var first = g.First();
                    return new
                    {
                        Group = first.Group ?? string.Empty,
                        Material = first.Material ?? string.Empty,
                        Unit = first.Unit ?? string.Empty,
                        Quantity = g.Sum(x => x.Quantity)
                    };
                })
                .OrderBy(x => x.Group, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.Material, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var summaryRows = aggregatedRows
                .Select(x => (IReadOnlyList<string>)new[]
                {
                    x.Group,
                    x.Material,
                    FormatNumberByUnit(x.Quantity, x.Unit),
                    x.Unit
                })
                .ToList();

            body.Append(BuildWordTable(
                new[] { "Тип", "Наименование", "Итого количество", "Ед." },
                summaryRows));

            mainPart.Document.Save();
        }

        private static DocumentFormat.OpenXml.Wordprocessing.Table BuildWordTable(
            IReadOnlyList<string> headers,
            IEnumerable<IReadOnlyList<string>> rows)
        {
            var table = new DocumentFormat.OpenXml.Wordprocessing.Table();
            var borders = new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
                new DocumentFormat.OpenXml.Wordprocessing.TopBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.BottomBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.LeftBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.RightBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 });
            table.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.TableProperties(borders));

            AppendWordTableRow(table, headers, isHeader: true);

            foreach (var row in rows ?? Enumerable.Empty<IReadOnlyList<string>>())
            {
                AppendWordTableRow(table, row ?? Array.Empty<string>(), isHeader: false);
            }

            return table;
        }

        private static void AppendWordTableRow(
            DocumentFormat.OpenXml.Wordprocessing.Table table,
            IReadOnlyList<string> values,
            bool isHeader)
        {
            var tableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
            foreach (var value in values ?? Array.Empty<string>())
            {
                var run = new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text(value ?? string.Empty)
                    {
                        Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
                    });
                if (isHeader)
                    run.RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(new DocumentFormat.OpenXml.Wordprocessing.Bold());

                var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run);
                var tableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(paragraph);
                tableRow.Append(tableCell);
            }

            table.Append(tableRow);
        }

        private void OpenSummaryBalanceWindow_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            EnsureProjectUiSettings();
            currentObject.SummaryBalanceHistory ??= new List<SummaryBalanceHistoryEntry>();

            var settings = currentObject.UiSettings ?? new ProjectUiSettings();
            var sourceRows = BuildSummaryBalanceItems(includeOverage: true, includeDeficit: true, onlyMainCategory: settings.SummaryReminderOnlyMain);
            var editableRows = new ObservableCollection<SummaryBalanceEditorRow>(
                sourceRows.Select(x => new SummaryBalanceEditorRow
                {
                    Category = x.Category ?? string.Empty,
                    Group = x.Group ?? string.Empty,
                    Material = x.Material ?? string.Empty,
                    Unit = x.Unit ?? string.Empty,
                    Quantity = x.Quantity,
                    IsOverage = x.IsOverage
                }));
            var historyRows = new ObservableCollection<SummaryBalanceHistoryEntry>(
                currentObject.SummaryBalanceHistory
                    .OrderByDescending(x => x.CreatedAt)
                    .Take(500));

            var dialog = new Window
            {
                Title = "Дефициты / излишки",
                Owner = this,
                Width = 1220,
                Height = 760,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            dialog.Content = root;

            var caption = new TextBlock
            {
                Text = "Текущие позиции по сводке и история с причинами.",
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            };
            Grid.SetRow(caption, 0);
            root.Children.Add(caption);

            var contentGrid = new Grid();
            contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(0.56, GridUnitType.Star) });
            contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(0.44, GridUnitType.Star) });
            Grid.SetRow(contentGrid, 1);
            root.Children.Add(contentGrid);

            var currentGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = false,
                ItemsSource = editableRows
            };
            currentGrid.Columns.Add(new DataGridTextColumn { Header = "Тип", Binding = new Binding(nameof(SummaryBalanceEditorRow.Group)), Width = 170, IsReadOnly = true });
            currentGrid.Columns.Add(new DataGridTextColumn { Header = "Наименование", Binding = new Binding(nameof(SummaryBalanceEditorRow.Material)), Width = 240, IsReadOnly = true });
            currentGrid.Columns.Add(new DataGridTextColumn { Header = "Сценарий", Binding = new Binding(nameof(SummaryBalanceEditorRow.Scenario)), Width = 120, IsReadOnly = true });
            currentGrid.Columns.Add(new DataGridTextColumn { Header = "Количество", Binding = new Binding(nameof(SummaryBalanceEditorRow.Quantity)) { StringFormat = "0.###" }, Width = 110, IsReadOnly = true });
            currentGrid.Columns.Add(new DataGridTextColumn { Header = "Ед.", Binding = new Binding(nameof(SummaryBalanceEditorRow.Unit)), Width = 70, IsReadOnly = true });
            currentGrid.Columns.Add(new DataGridTextColumn { Header = "Причина", Binding = new Binding(nameof(SummaryBalanceEditorRow.Reason)) { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            DataGridSizingHelper.SetEnableSmartSizing(currentGrid, true);
            Grid.SetRow(currentGrid, 0);
            contentGrid.Children.Add(currentGrid);

            var historyGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                ItemsSource = historyRows
            };
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Когда", Binding = new Binding(nameof(SummaryBalanceHistoryEntry.CreatedAt)) { StringFormat = "dd.MM.yyyy HH:mm" }, Width = 140 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Тип", Binding = new Binding(nameof(SummaryBalanceHistoryEntry.Group)), Width = 160 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Наименование", Binding = new Binding(nameof(SummaryBalanceHistoryEntry.Material)), Width = 220 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Сценарий", Binding = new Binding(nameof(SummaryBalanceHistoryEntry.IsOverage)) { Converter = new BoolToScenarioConverter() }, Width = 110 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Количество", Binding = new Binding(nameof(SummaryBalanceHistoryEntry.Quantity)) { StringFormat = "0.###" }, Width = 110 });
            historyGrid.Columns.Add(new DataGridTextColumn { Header = "Причина", Binding = new Binding(nameof(SummaryBalanceHistoryEntry.Reason)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            DataGridSizingHelper.SetEnableSmartSizing(historyGrid, true);
            Grid.SetRow(historyGrid, 1);
            contentGrid.Children.Add(historyGrid);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            var saveButton = new Button { Content = "Сохранить причины в историю", MinWidth = 230 };
            var closeButton = new Button { Content = "Закрыть", MinWidth = 110, Margin = new Thickness(8, 0, 0, 0), IsCancel = true };
            footer.Children.Add(saveButton);
            footer.Children.Add(closeButton);

            saveButton.Click += (_, _) =>
            {
                foreach (var row in editableRows.Where(x => !string.IsNullOrWhiteSpace(x.Reason)))
                {
                    currentObject.SummaryBalanceHistory.Add(new SummaryBalanceHistoryEntry
                    {
                        CreatedAt = DateTime.Now,
                        Group = row.Group ?? string.Empty,
                        Material = row.Material ?? string.Empty,
                        Unit = row.Unit ?? string.Empty,
                        Quantity = row.Quantity,
                        IsOverage = row.IsOverage,
                        Reason = row.Reason?.Trim() ?? string.Empty
                    });
                }

                currentObject.SummaryBalanceHistory = currentObject.SummaryBalanceHistory
                    .OrderByDescending(x => x.CreatedAt)
                    .Take(5000)
                    .ToList();

                historyRows.Clear();
                foreach (var historyItem in currentObject.SummaryBalanceHistory.OrderByDescending(x => x.CreatedAt).Take(500))
                    historyRows.Add(historyItem);

                SaveState();
                MessageBox.Show("История обновлена.");
            };

            dialog.ShowDialog();
        }

        private void OpenSummaryComparisonWindow_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            var rows = BuildSummaryComparisonRows();
            if (rows.Count == 0)
            {
                MessageBox.Show("Нет данных для сравнения.");
                return;
            }

            var dialog = new Window
            {
                Title = "Сравнение план / пришло / смонтировано / остаток",
                Owner = this,
                Width = 1080,
                Height = 680,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var grid = new Grid { Margin = new Thickness(14) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            dialog.Content = grid;

            var caption = new TextBlock
            {
                Text = "Позиции по всем типам. Остаток = Пришло - Смонтировано.",
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            };
            Grid.SetRow(caption, 0);
            grid.Children.Add(caption);

            var dg = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                ItemsSource = rows
            };
            dg.Columns.Add(new DataGridTextColumn { Header = "Тип", Binding = new Binding(nameof(SummaryComparisonRow.Тип)), Width = 170 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Наименование", Binding = new Binding(nameof(SummaryComparisonRow.Наименование)), Width = 230 });
            dg.Columns.Add(new DataGridTextColumn { Header = "План", Binding = new Binding(nameof(SummaryComparisonRow.План)) { StringFormat = "0.###" }, Width = 110 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Пришло", Binding = new Binding(nameof(SummaryComparisonRow.Пришло)) { StringFormat = "0.###" }, Width = 110 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Смонтировано", Binding = new Binding(nameof(SummaryComparisonRow.Смонтировано)) { StringFormat = "0.###" }, Width = 130 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Остаток", Binding = new Binding(nameof(SummaryComparisonRow.Остаток)) { StringFormat = "0.###" }, Width = 110 });
            dg.Columns.Add(new DataGridTextColumn { Header = "Ед.", Binding = new Binding(nameof(SummaryComparisonRow.Ед)), Width = 70 });
            DataGridSizingHelper.SetEnableSmartSizing(dg, true);
            Grid.SetRow(dg, 1);
            grid.Children.Add(dg);

            dialog.ShowDialog();
        }

        private List<SummaryComparisonRow> BuildSummaryComparisonRows()
        {
            var keys = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

            if (currentObject?.Demand != null)
            {
                foreach (var key in currentObject.Demand.Keys.Where(x => !string.IsNullOrWhiteSpace(x)))
                    keys.Add(key.Trim());
            }

            foreach (var row in journal.Where(x => !string.IsNullOrWhiteSpace(x.MaterialGroup) && !string.IsNullOrWhiteSpace(x.MaterialName)))
                keys.Add(BuildDemandKey(row.MaterialGroup.Trim(), row.MaterialName.Trim()));

            return keys
                .Select(key =>
                {
                    var parts = key.Split(new[] { "::" }, 2, StringSplitOptions.None);
                    var group = parts[0].Trim();
                    var material = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                    if (string.IsNullOrWhiteSpace(group) || string.IsNullOrWhiteSpace(material))
                        return null;

                    var unit = GetUnitForMaterial(group, material);
                    var demand = GetOrCreateDemand(key, unit);
                    var plan = NormalizeQuantityByUnit(
                        BuildSummaryBlocks(group).SelectMany(x => x.Levels.Select(level => GetDemandValue(demand, x.Block, level))).Sum(),
                        unit);
                    var arrived = NormalizeQuantityByUnit(
                        journal.Where(x =>
                                string.Equals((x.MaterialGroup ?? string.Empty).Trim(), group, StringComparison.CurrentCultureIgnoreCase)
                                && string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase))
                            .Sum(x => x.Quantity),
                        unit);
                    var mounted = NormalizeQuantityByUnit(
                        GetMountedQuantityFromProductionJournal(material),
                        unit);

                    return new SummaryComparisonRow
                    {
                        Тип = group,
                        Наименование = material,
                        Ед = unit,
                        План = plan,
                        Пришло = arrived,
                        Смонтировано = mounted,
                        Остаток = arrived - mounted
                    };
                })
                .Where(x => x != null)
                .OrderBy(x => x.Тип, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.Наименование, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
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

        private static string NormalizeJournalToolsPanelKey(object tag)
            => (tag as string ?? string.Empty).Trim();

        private bool GetJournalToolsPanelPinned(string key) => key switch
        {
            "Ot" => isOtToolsPinned,
            "Timesheet" => isTimesheetToolsPinned,
            "Production" => isProductionToolsPinned,
            "Inspection" => isInspectionToolsPinned,
            _ => false
        };

        private void SetJournalToolsPanelPinned(string key, bool value)
        {
            switch (key)
            {
                case "Ot":
                    isOtToolsPinned = value;
                    break;
                case "Timesheet":
                    isTimesheetToolsPinned = value;
                    break;
                case "Production":
                    isProductionToolsPinned = value;
                    break;
                case "Inspection":
                    isInspectionToolsPinned = value;
                    break;
            }
        }

        private bool TryGetJournalToolsPanelElements(
            string key,
            out Border hoverStrip,
            out Border panelBorder,
            out ToggleButton pinToggle)
        {
            switch (key)
            {
                case "Ot":
                    hoverStrip = OtToolsHoverStrip;
                    panelBorder = OtToolsPanelBorder;
                    pinToggle = OtToolsPinToggle;
                    return true;
                case "Timesheet":
                    hoverStrip = TimesheetToolsHoverStrip;
                    panelBorder = TimesheetToolsPanelBorder;
                    pinToggle = TimesheetToolsPinToggle;
                    return true;
                case "Production":
                    hoverStrip = ProductionToolsHoverStrip;
                    panelBorder = ProductionToolsPanelBorder;
                    pinToggle = ProductionToolsPinToggle;
                    return true;
                case "Inspection":
                    hoverStrip = InspectionToolsHoverStrip;
                    panelBorder = InspectionToolsPanelBorder;
                    pinToggle = InspectionToolsPinToggle;
                    return true;
                default:
                    hoverStrip = null;
                    panelBorder = null;
                    pinToggle = null;
                    return false;
            }
        }

        private void UpdateJournalToolsPanelState(string key, bool forceVisible)
        {
            if (!TryGetJournalToolsPanelElements(key, out var hoverStrip, out var panelBorder, out var pinToggle))
                return;

            var isPinned = GetJournalToolsPanelPinned(key);
            var show = isPinned || forceVisible;

            if (pinToggle != null && pinToggle.IsChecked != isPinned)
                pinToggle.IsChecked = isPinned;

            if (hoverStrip != null)
                hoverStrip.Visibility = show ? Visibility.Collapsed : Visibility.Visible;

            if (panelBorder != null)
                panelBorder.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
        }

        private void UpdateAllJournalToolsPanelStates()
        {
            UpdateJournalToolsPanelState("Ot", forceVisible: false);
            UpdateJournalToolsPanelState("Timesheet", forceVisible: false);
            UpdateJournalToolsPanelState("Production", forceVisible: false);
            UpdateJournalToolsPanelState("Inspection", forceVisible: false);
        }

        private void JournalToolsPinToggle_Checked(object sender, RoutedEventArgs e)
        {
            var key = NormalizeJournalToolsPanelKey((sender as FrameworkElement)?.Tag);
            if (string.IsNullOrWhiteSpace(key))
                return;

            SetJournalToolsPanelPinned(key, value: true);
            UpdateJournalToolsPanelState(key, forceVisible: true);
        }

        private void JournalToolsPinToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            var key = NormalizeJournalToolsPanelKey((sender as FrameworkElement)?.Tag);
            if (string.IsNullOrWhiteSpace(key))
                return;

            SetJournalToolsPanelPinned(key, value: false);
            UpdateJournalToolsPanelState(key, forceVisible: false);
        }

        private void JournalToolsHoverStrip_MouseEnter(object sender, MouseEventArgs e)
        {
            var key = NormalizeJournalToolsPanelKey((sender as FrameworkElement)?.Tag);
            if (string.IsNullOrWhiteSpace(key) || GetJournalToolsPanelPinned(key))
                return;

            UpdateJournalToolsPanelState(key, forceVisible: true);
        }

        private void JournalToolsPanel_MouseEnter(object sender, MouseEventArgs e)
        {
            var key = NormalizeJournalToolsPanelKey((sender as FrameworkElement)?.Tag);
            if (string.IsNullOrWhiteSpace(key) || GetJournalToolsPanelPinned(key))
                return;

            UpdateJournalToolsPanelState(key, forceVisible: true);
        }

        private void JournalToolsPanel_MouseLeave(object sender, MouseEventArgs e)
        {
            var key = NormalizeJournalToolsPanelKey((sender as FrameworkElement)?.Tag);
            if (string.IsNullOrWhiteSpace(key) || GetJournalToolsPanelPinned(key))
                return;

            if (TryGetJournalToolsPanelElements(key, out _, out var panelBorder, out _) && !panelBorder.IsMouseOver)
                UpdateJournalToolsPanelState(key, forceVisible: false);
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
            {
                currentObject.UiSettings ??= new ProjectUiSettings();
                currentObject.ChangeLog ??= new List<ProjectChangeLogEntry>();
                currentObject.UiSettings.CommandPaletteShortcuts ??= new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
                currentObject.UiSettings.TabDisplayModes ??= new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
                currentObject.UiSettings.GridColumnPreferences ??= new Dictionary<string, List<GridColumnPreference>>(StringComparer.CurrentCultureIgnoreCase);
                currentObject.UiSettings.GridColumnPresets ??= new Dictionary<string, Dictionary<string, List<GridColumnPreference>>>(StringComparer.CurrentCultureIgnoreCase);
                EnsureReferenceMappingsStorage();
            }

            if (currentObject?.UiSettings != null && currentObject.UiSettings.ReminderSnoozeMinutes <= 0)
                currentObject.UiSettings.ReminderSnoozeMinutes = 15;

            if (currentObject?.UiSettings != null && currentObject.UiSettings.AutoSaveIntervalMinutes <= 0)
                currentObject.UiSettings.AutoSaveIntervalMinutes = 5;

            if (currentObject?.UiSettings != null)
                currentObject.UiSettings.DataRootDirectory = NormalizeDataRootPath(currentObject.UiSettings.DataRootDirectory);

            if (currentObject?.UiSettings != null && string.IsNullOrWhiteSpace(currentObject.UiSettings.UiDensityMode))
                currentObject.UiSettings.UiDensityMode = "Стандартный";

            if (currentObject?.UiSettings != null)
                currentObject.UiSettings.ReminderPresentationMode = NormalizeReminderPresentationMode(currentObject.UiSettings.ReminderPresentationMode);

            if (currentObject?.UiSettings != null && string.IsNullOrWhiteSpace(currentObject.UiSettings.AccessRole))
            {
                currentObject.UiSettings.AccessRole = ProjectAccessRoles.Critical;
                if (!currentObject.UiSettings.RequireCodeForCriticalOperations)
                    currentObject.UiSettings.RequireCodeForCriticalOperations = true;
            }

            if (currentObject?.UiSettings != null)
                currentObject.UiSettings.AccessRole = NormalizeAccessRole(currentObject.UiSettings.AccessRole);

            if (currentObject?.UiSettings != null && string.IsNullOrWhiteSpace(currentObject.UiSettings.OtStatusFilter))
                currentObject.UiSettings.OtStatusFilter = "Все";
            if (currentObject?.UiSettings != null && string.IsNullOrWhiteSpace(currentObject.UiSettings.OtSpecialtyFilter))
                currentObject.UiSettings.OtSpecialtyFilter = "Все";
            if (currentObject?.UiSettings != null && string.IsNullOrWhiteSpace(currentObject.UiSettings.OtBrigadeFilter))
                currentObject.UiSettings.OtBrigadeFilter = "Все";

            if (currentObject?.UiSettings != null
                && !currentObject.UiSettings.SummaryReminderOnOverage
                && !currentObject.UiSettings.SummaryReminderOnDeficit
                && !currentObject.UiSettings.SummaryReminderOnlyMain)
            {
                currentObject.UiSettings.SummaryReminderOnOverage = true;
                currentObject.UiSettings.SummaryReminderOnlyMain = true;
            }
        }

        private void ApplyProjectUiSettings()
        {
            EnsureProjectUiSettings();
            InitializeSpreadsheetEditorPreference();
            InitializePdfEditorPreference();
            ApplyUiDensityMode();
            UpdateAutoSaveTimerInterval();

            if (currentObject?.UiSettings == null)
            {
                isOtToolsPinned = false;
                isTimesheetToolsPinned = false;
                isProductionToolsPinned = false;
                isInspectionToolsPinned = false;
                UpdateAllJournalToolsPanelStates();
                UpdateTreePanelState(forceVisible: false);
                return;
            }

            isTreePinned = currentObject.UiSettings.PinTreeByDefault && !currentObject.UiSettings.DisableTree;
            if (TreePinToggle != null)
                TreePinToggle.IsChecked = isTreePinned;

            var pinJournalPanels = currentObject.UiSettings.PinJournalPanelsByDefault;
            isOtToolsPinned = pinJournalPanels;
            isTimesheetToolsPinned = pinJournalPanels;
            isProductionToolsPinned = pinJournalPanels;
            isInspectionToolsPinned = pinJournalPanels;

            UpdateTreePanelState(forceVisible: isTreePinned);
            UpdateAllJournalToolsPanelStates();
            ApplySavedTabDisplayModes();
            ApplyGridColumnPreferencesForAll();
            RequestReminderRefresh(immediate: true);
            UpdateStatusBar();
        }

        private void UpdateStatusBar()
        {
            if (StatusBarObjectText == null)
                return;

            var objectName = string.IsNullOrWhiteSpace(currentObject?.Name) ? "Не создан" : currentObject.Name.Trim();
            var fileText = string.IsNullOrWhiteSpace(currentSaveFileName)
                ? "—"
                : System.IO.Path.GetFileName(currentSaveFileName);
            var autoSaveMinutes = Math.Max(1, currentObject?.UiSettings?.AutoSaveIntervalMinutes ?? 5);
            var lastSaveText = lastSuccessfulSaveLocalTime.HasValue
                ? lastSuccessfulSaveLocalTime.Value.ToString("dd.MM.yyyy HH:mm:ss", CultureInfo.CurrentCulture)
                : "Нет данных";
            var processingText = processingOverlayDepth > 0
                ? processingOverlayPendingText
                : "Ожидание";

            StatusBarObjectText.Text = objectName;
            StatusBarFileText.Text = fileText;
            StatusBarAutoSaveText.Text = $"{autoSaveMinutes} мин";
            StatusBarRoleText.Text = GetAccessRoleDisplayName(GetCurrentAccessRole());
            StatusBarLockText.Text = isLocked ? "Вкл" : "Выкл";
            StatusBarLastSaveText.Text = lastSaveText;
            StatusBarOperationText.Text = string.IsNullOrWhiteSpace(lastOperationStatusText) ? "Готово" : lastOperationStatusText;
            StatusBarProcessingText.Text = processingText;
        }

        private static string GetAccessRoleDisplayName(string role)
        {
            if (string.Equals(role, ProjectAccessRoles.View, StringComparison.CurrentCultureIgnoreCase))
                return "Просмотр";
            if (string.Equals(role, ProjectAccessRoles.Edit, StringComparison.CurrentCultureIgnoreCase))
                return "Редактирование";

            return "Полный";
        }

        private void SetLastOperationStatus(string text)
        {
            lastOperationStatusText = string.IsNullOrWhiteSpace(text) ? "Готово" : text.Trim();
            UpdateStatusBar();
        }

        private void AddOperationLogEntry(string kind, string status, string details)
        {
            var entry = new OperationLogEntry
            {
                TimestampLocal = DateTime.Now,
                Kind = string.IsNullOrWhiteSpace(kind) ? "Система" : kind.Trim(),
                Status = string.IsNullOrWhiteSpace(status) ? "Информация" : status.Trim(),
                Details = string.IsNullOrWhiteSpace(details) ? "—" : details.Trim()
            };

            operationLogEntries.Insert(0, entry);
            while (operationLogEntries.Count > 300)
                operationLogEntries.RemoveAt(operationLogEntries.Count - 1);
        }

        private string NormalizeUiDensityMode(string densityMode)
        {
            var normalized = (densityMode ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
                return "Стандартный";

            if (string.Equals(normalized, "Компактный", StringComparison.CurrentCultureIgnoreCase)
                || normalized.Contains("РљРѕРјРї", StringComparison.Ordinal)
                || string.Equals(normalized, "compact", StringComparison.CurrentCultureIgnoreCase))
            {
                return "Компактный";
            }

            return "Стандартный";
        }

        private void ApplyUiDensityMode()
        {
            var mode = NormalizeUiDensityMode(currentObject?.UiSettings?.UiDensityMode);
            if (currentObject?.UiSettings != null)
                currentObject.UiSettings.UiDensityMode = mode;

            var resources = Application.Current?.Resources;
            if (resources == null)
                return;

            var compact = string.Equals(mode, "Компактный", StringComparison.CurrentCultureIgnoreCase);

            void Assign(string targetKey, string standardKey, string compactKey)
            {
                var sourceKey = compact ? compactKey : standardKey;
                if (resources.Contains(sourceKey))
                    resources[targetKey] = resources[sourceKey];
            }

            Assign("AppFontSize", "AppFontSizeStandard", "AppFontSizeCompact");
            Assign("ControlHeight", "ControlHeightStandard", "ControlHeightCompact");
            Assign("TextBoxPadding", "TextBoxPaddingStandard", "TextBoxPaddingCompact");
            Assign("ComboPadding", "ComboPaddingStandard", "ComboPaddingCompact");
            Assign("ButtonPadding", "ButtonPaddingStandard", "ButtonPaddingCompact");
            Assign("TabPadding", "TabPaddingStandard", "TabPaddingCompact");
            Assign("DataGridRowMinHeight", "DataGridRowMinHeightStandard", "DataGridRowMinHeightCompact");
            Assign("DataGridCellPadding", "DataGridCellPaddingStandard", "DataGridCellPaddingCompact");
        }

        private void ApplySavedTabDisplayModes()
        {
            if (currentObject?.UiSettings?.TabDisplayModes == null)
                return;

            if (currentObject.UiSettings.TabDisplayModes.TryGetValue("Приход", out var arrivalMode))
                arrivalMatrixMode = string.Equals(arrivalMode, "Матрица", StringComparison.CurrentCultureIgnoreCase);

            UpdateArrivalViewMode();

            if (currentObject.UiSettings.TabDisplayModes.TryGetValue(ActiveTabModeKey, out var tabHeader))
            {
                var tab = GetTabByHeader(tabHeader);
                if (tab != null && MainTabs != null && !ReferenceEquals(MainTabs.SelectedItem, tab))
                    MainTabs.SelectedItem = tab;
            }
        }

        private void SetTabDisplayMode(string tabHeader, string mode)
        {
            if (string.IsNullOrWhiteSpace(tabHeader))
                return;

            EnsureProjectUiSettings();
            currentObject?.UiSettings?.TabDisplayModes?.TryAdd(tabHeader, mode ?? string.Empty);
            if (currentObject?.UiSettings?.TabDisplayModes != null)
                currentObject.UiSettings.TabDisplayModes[tabHeader] = mode ?? string.Empty;
        }

        private TabItem GetTabByHeader(string header)
        {
            if (string.IsNullOrWhiteSpace(header))
                return null;

            if (string.Equals(header, "Сводка", StringComparison.CurrentCultureIgnoreCase))
                return SummaryTab;
            if (string.Equals(header, "ЖВК", StringComparison.CurrentCultureIgnoreCase))
                return JvkTab;
            if (string.Equals(header, "Приход", StringComparison.CurrentCultureIgnoreCase))
                return ArrivalTab;
            if (string.Equals(header, "ОТ", StringComparison.CurrentCultureIgnoreCase))
                return OtTab;
            if (string.Equals(header, "Табель", StringComparison.CurrentCultureIgnoreCase))
                return TimesheetTab;
            if (string.Equals(header, "ПР", StringComparison.CurrentCultureIgnoreCase))
                return ProductionTab;
            if (string.Equals(header, "Осмотры", StringComparison.CurrentCultureIgnoreCase))
                return InspectionTab;
            if (string.Equals(header, "ПДФ", StringComparison.CurrentCultureIgnoreCase))
                return PdfTab;
            if (string.Equals(header, "Сметы", StringComparison.CurrentCultureIgnoreCase))
                return EstimateTab;

            return null;
        }

        private Dictionary<string, DataGrid> GetManagedGridMap()
        {
            var map = new Dictionary<string, DataGrid>(StringComparer.CurrentCultureIgnoreCase);
            if (ArrivalLegacyGrid != null)
                map[GridPrefArrival] = ArrivalLegacyGrid;
            if (OtJournalGrid != null)
                map[GridPrefOt] = OtJournalGrid;
            if (TimesheetGrid != null)
                map[GridPrefTimesheet] = TimesheetGrid;
            if (ProductionJournalGrid != null)
                map[GridPrefProduction] = ProductionJournalGrid;
            if (InspectionJournalGrid != null)
                map[GridPrefInspection] = InspectionJournalGrid;
            return map;
        }

        private string GetGridPreferenceKey(DataGrid grid)
        {
            if (ReferenceEquals(grid, ArrivalLegacyGrid))
                return GridPrefArrival;
            if (ReferenceEquals(grid, OtJournalGrid))
                return GridPrefOt;
            if (ReferenceEquals(grid, TimesheetGrid))
                return GridPrefTimesheet;
            if (ReferenceEquals(grid, ProductionJournalGrid))
                return GridPrefProduction;
            if (ReferenceEquals(grid, InspectionJournalGrid))
                return GridPrefInspection;
            return string.Empty;
        }

        private static string GetColumnHeaderKey(DataGridColumn column)
        {
            if (column?.Header == null)
                return string.Empty;

            return column.Header.ToString()?.Trim() ?? string.Empty;
        }

        private static double GetColumnWidthValue(DataGridColumn column)
        {
            if (column == null)
                return double.NaN;

            if (column.Width.IsAbsolute && column.Width.Value > 0)
                return Math.Round(column.Width.Value, 2);

            if (column.ActualWidth > 0)
                return Math.Round(column.ActualWidth, 2);

            return double.NaN;
        }

        private void SaveGridColumnPreferencesForAll()
        {
            foreach (var grid in GetManagedGridMap().Values)
                SaveGridColumnPreferences(grid);
        }

        private void SaveGridColumnPreferences(DataGrid grid)
        {
            if (isApplyingColumnPreferences || grid == null || currentObject?.UiSettings == null)
                return;

            var key = GetGridPreferenceKey(grid);
            if (string.IsNullOrWhiteSpace(key))
                return;

            EnsureProjectUiSettings();

            var list = grid.Columns
                .OrderBy(x => x.DisplayIndex)
                .Select(x => new GridColumnPreference
                {
                    Header = GetColumnHeaderKey(x),
                    IsVisible = x.Visibility == Visibility.Visible,
                    DisplayIndex = x.DisplayIndex,
                    Width = GetColumnWidthValue(x)
                })
                .ToList();

            currentObject.UiSettings.GridColumnPreferences[key] = list;
        }

        private Dictionary<string, List<GridColumnPreference>> GetGridPresetStore(DataGrid grid)
        {
            if (grid == null)
                return null;

            var key = GetGridPreferenceKey(grid);
            if (string.IsNullOrWhiteSpace(key))
                return null;

            EnsureProjectUiSettings();
            currentObject.UiSettings.GridColumnPresets ??= new Dictionary<string, Dictionary<string, List<GridColumnPreference>>>(StringComparer.CurrentCultureIgnoreCase);
            if (!currentObject.UiSettings.GridColumnPresets.TryGetValue(key, out var store) || store == null)
            {
                store = new Dictionary<string, List<GridColumnPreference>>(StringComparer.CurrentCultureIgnoreCase);
                currentObject.UiSettings.GridColumnPresets[key] = store;
            }

            return store;
        }

        private List<string> GetGridPresetNames(DataGrid grid)
        {
            var store = GetGridPresetStore(grid);
            if (store == null)
                return new List<string>();

            return store.Keys.OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        private void SaveGridColumnPreset(DataGrid grid, string presetName)
        {
            if (grid == null || string.IsNullOrWhiteSpace(presetName))
                return;

            var store = GetGridPresetStore(grid);
            if (store == null)
                return;

            var list = grid.Columns
                .OrderBy(x => x.DisplayIndex)
                .Select(x => new GridColumnPreference
                {
                    Header = GetColumnHeaderKey(x),
                    IsVisible = x.Visibility == Visibility.Visible,
                    DisplayIndex = x.DisplayIndex,
                    Width = GetColumnWidthValue(x)
                })
                .ToList();

            store[presetName.Trim()] = list;
        }

        private void ApplyGridColumnPreset(DataGrid grid, string presetName)
        {
            if (grid == null || string.IsNullOrWhiteSpace(presetName))
                return;

            var store = GetGridPresetStore(grid);
            if (store == null || !store.TryGetValue(presetName.Trim(), out var prefs) || prefs == null || prefs.Count == 0)
                return;

            ApplyGridColumnPreferences(grid, prefs);
        }

        private void RemoveGridColumnPreset(DataGrid grid, string presetName)
        {
            if (grid == null || string.IsNullOrWhiteSpace(presetName))
                return;

            var store = GetGridPresetStore(grid);
            if (store == null)
                return;

            store.Remove(presetName.Trim());
        }

        private void ApplyGridColumnPreferencesForAll()
        {
            foreach (var grid in GetManagedGridMap().Values)
                ApplyGridColumnPreferences(grid);
        }

        private void ApplyGridColumnPreferences(DataGrid grid)
        {
            if (grid == null || currentObject?.UiSettings?.GridColumnPreferences == null)
                return;

            var key = GetGridPreferenceKey(grid);
            if (string.IsNullOrWhiteSpace(key))
                return;

            if (!currentObject.UiSettings.GridColumnPreferences.TryGetValue(key, out var prefs) || prefs == null || prefs.Count == 0)
                return;

            ApplyGridColumnPreferences(grid, prefs);
        }

        private void ApplyGridColumnPreferences(DataGrid grid, List<GridColumnPreference> prefs)
        {
            if (grid == null || prefs == null || prefs.Count == 0)
                return;

            var prefByHeader = prefs
                .Where(x => !string.IsNullOrWhiteSpace(x.Header))
                .GroupBy(x => x.Header.Trim(), StringComparer.CurrentCultureIgnoreCase)
                .ToDictionary(x => x.Key, x => x.First(), StringComparer.CurrentCultureIgnoreCase);

            isApplyingColumnPreferences = true;
            try
            {
                var ordered = grid.Columns
                    .Select(x =>
                    {
                        var header = GetColumnHeaderKey(x);
                        prefByHeader.TryGetValue(header, out var pref);
                        return new { Column = x, Preference = pref };
                    })
                    .OrderBy(x => x.Preference?.DisplayIndex ?? int.MaxValue)
                    .ThenBy(x => x.Column.DisplayIndex)
                    .ToList();

                for (var i = 0; i < ordered.Count; i++)
                {
                    ordered[i].Column.DisplayIndex = i;

                    var pref = ordered[i].Preference;
                    if (pref == null)
                        continue;

                    ordered[i].Column.Visibility = pref.IsVisible ? Visibility.Visible : Visibility.Collapsed;
                    if (!double.IsNaN(pref.Width) && pref.Width > 20)
                        ordered[i].Column.Width = new DataGridLength(pref.Width);
                }
            }
            catch
            {
                // Настройки колонок не должны ломать вкладку
            }
            finally
            {
                isApplyingColumnPreferences = false;
            }
        }

        private DataGrid GetCurrentTabGrid()
        {
            if (MainTabs?.SelectedItem is not TabItem tab)
                return null;

            return GetGridByTabHeader(tab.Header?.ToString());
        }

        private void UiDensityStandard_Click(object sender, RoutedEventArgs e)
            => SetUiDensityMode("Стандартный");

        private void UiDensityCompact_Click(object sender, RoutedEventArgs e)
            => SetUiDensityMode("Компактный");

        private void AutoSizeCurrentGridColumns_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetCurrentTabGrid();
            if (grid == null || grid.Columns.Count == 0)
            {
                MessageBox.Show("Для текущей вкладки нет настраиваемой таблицы.");
                return;
            }

            DataGridSizingHelper.SetEnableSmartSizing(grid, false);
            DataGridSizingHelper.SetEnableSmartSizing(grid, true);
            SaveGridColumnPreferences(grid);
            SaveState(SaveTrigger.System);
            SetLastOperationStatus("Колонки подогнаны");
        }

        private void ResetCurrentGridColumns_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetCurrentTabGrid();
            if (grid == null || grid.Columns.Count == 0)
            {
                MessageBox.Show("Для текущей вкладки нет настраиваемой таблицы.");
                return;
            }

            var key = GetGridPreferenceKey(grid);
            if (!string.IsNullOrWhiteSpace(key))
                currentObject?.UiSettings?.GridColumnPreferences?.Remove(key);

            isApplyingColumnPreferences = true;
            try
            {
                foreach (var column in grid.Columns)
                {
                    column.MinWidth = 0;
                    column.Width = DataGridLength.Auto;
                }
            }
            finally
            {
                isApplyingColumnPreferences = false;
            }

            DataGridSizingHelper.SetEnableSmartSizing(grid, false);
            DataGridSizingHelper.SetEnableSmartSizing(grid, true);
            SaveState(SaveTrigger.System);
            SetLastOperationStatus("Ширины колонок сброшены");
        }

        private void SaveCurrentGridPreset_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetCurrentTabGrid();
            if (grid == null || grid.Columns.Count == 0)
            {
                MessageBox.Show("Для текущей вкладки нет настраиваемой таблицы.");
                return;
            }

            var name = Microsoft.VisualBasic.Interaction.InputBox(
                "Введите название пресета:",
                "Пресет колонок",
                "Новый пресет");

            if (string.IsNullOrWhiteSpace(name))
                return;

            SaveGridColumnPreset(grid, name.Trim());
            SaveState(SaveTrigger.System);
            SetLastOperationStatus($"Пресет колонок сохранен: {name.Trim()}");
        }

        private void ApplyCurrentGridPreset_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetCurrentTabGrid();
            if (grid == null || grid.Columns.Count == 0)
            {
                MessageBox.Show("Для текущей вкладки нет настраиваемой таблицы.");
                return;
            }

            var names = GetGridPresetNames(grid);
            if (names.Count == 0)
            {
                MessageBox.Show("Для текущей вкладки нет сохраненных пресетов.");
                return;
            }

            var selected = PromptSelectOption("Выберите пресет колонок", "Пресет", names);
            if (string.IsNullOrWhiteSpace(selected))
                return;

            ApplyGridColumnPreset(grid, selected);
            SaveState(SaveTrigger.System);
            SetLastOperationStatus($"Пресет колонок применен: {selected}");
        }

        private void DeleteCurrentGridPreset_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetCurrentTabGrid();
            if (grid == null || grid.Columns.Count == 0)
            {
                MessageBox.Show("Для текущей вкладки нет настраиваемой таблицы.");
                return;
            }

            var names = GetGridPresetNames(grid);
            if (names.Count == 0)
            {
                MessageBox.Show("Для текущей вкладки нет сохраненных пресетов.");
                return;
            }

            var selected = PromptSelectOption("Удалить пресет колонок", "Пресет", names);
            if (string.IsNullOrWhiteSpace(selected))
                return;

            if (MessageBox.Show($"Удалить пресет \"{selected}\"?", "Пресеты колонок", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            RemoveGridColumnPreset(grid, selected);
            SaveState(SaveTrigger.System);
            SetLastOperationStatus($"Пресет колонок удален: {selected}");
        }

        private DataGrid GetGridByTabHeader(string header)
        {
            header ??= string.Empty;
            if (string.Equals(header, "Приход", StringComparison.CurrentCultureIgnoreCase))
                return ArrivalLegacyGrid;
            if (string.Equals(header, "ОТ", StringComparison.CurrentCultureIgnoreCase))
                return OtJournalGrid;
            if (string.Equals(header, "Табель", StringComparison.CurrentCultureIgnoreCase))
                return TimesheetGrid;
            if (string.Equals(header, "ПР", StringComparison.CurrentCultureIgnoreCase))
                return ProductionJournalGrid;
            if (string.Equals(header, "Осмотры", StringComparison.CurrentCultureIgnoreCase))
                return InspectionJournalGrid;

            return null;
        }

        private static bool TryParsePositiveWidth(string text, out double width)
        {
            width = 0;
            var normalized = (text ?? string.Empty).Trim().Replace(',', '.');
            if (!double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                return false;

            if (parsed <= 20)
                return false;

            width = parsed;
            return true;
        }

        private void OpenColumnManager_Click(object sender, RoutedEventArgs e)
        {
            if (isOpeningCommandDialog)
                return;

            var grid = GetCurrentTabGrid();
            if (grid == null || grid.Columns.Count == 0)
            {
                MessageBox.Show("Для текущей вкладки нет настраиваемой таблицы.");
                return;
            }

            isOpeningCommandDialog = true;
            try
            {
                ShowColumnManagerDialog(grid);
            }
            finally
            {
                isOpeningCommandDialog = false;
            }
        }

        private void ShowColumnManagerDialog(DataGrid grid)
        {
            if (grid == null)
                return;

            EnsureProjectUiSettings();

            var rows = new ObservableCollection<ColumnManagerRow>(
                grid.Columns
                    .OrderBy(x => x.DisplayIndex)
                    .Select((x, idx) => new ColumnManagerRow
                    {
                        Header = string.IsNullOrWhiteSpace(GetColumnHeaderKey(x)) ? $"Колонка {idx + 1}" : GetColumnHeaderKey(x),
                        IsVisible = x.Visibility == Visibility.Visible,
                        WidthText = GetColumnWidthValue(x).ToString("0.##", CultureInfo.InvariantCulture),
                        Order = idx + 1
                    }));

            var dialog = new Window
            {
                Title = "Менеджер колонок",
                Owner = this,
                Width = 760,
                Height = 620,
                MinWidth = 680,
                MinHeight = 520,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.CanResize
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var hint = new TextBlock
            {
                Text = "Управление колонками текущей вкладки: видимость, порядок, ширина.",
                Margin = new Thickness(0, 0, 0, 10),
                Foreground = new SolidColorBrush(Color.FromRgb(71, 85, 105))
            };
            Grid.SetRow(hint, 0);
            root.Children.Add(hint);

            var presetPanel = new Grid { Margin = new Thickness(0, 0, 0, 10) };
            presetPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            presetPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
            presetPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            presetPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            presetPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var presetNameBox = new TextBox
            {
                Margin = new Thickness(0, 0, 8, 0),
                VerticalContentAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(presetNameBox, 0);

            var presetCombo = new ComboBox
            {
                Margin = new Thickness(0, 0, 8, 0),
                MinWidth = 200,
                ItemsSource = new ObservableCollection<string>(GetGridPresetNames(grid))
            };
            if (presetCombo.Items.Count > 0)
                presetCombo.SelectedIndex = 0;
            Grid.SetColumn(presetCombo, 1);

            var savePresetButton = new Button
            {
                Content = "Сохранить пресет",
                MinWidth = 140,
                Margin = new Thickness(0, 0, 8, 0)
            };
            Grid.SetColumn(savePresetButton, 2);

            var applyPresetButton = new Button
            {
                Content = "Применить",
                MinWidth = 110,
                Margin = new Thickness(0, 0, 8, 0)
            };
            Grid.SetColumn(applyPresetButton, 3);

            var deletePresetButton = new Button
            {
                Content = "Удалить",
                MinWidth = 90
            };
            Grid.SetColumn(deletePresetButton, 4);

            presetPanel.Children.Add(presetNameBox);
            presetPanel.Children.Add(presetCombo);
            presetPanel.Children.Add(savePresetButton);
            presetPanel.Children.Add(applyPresetButton);
            presetPanel.Children.Add(deletePresetButton);
            Grid.SetRow(presetPanel, 1);
            root.Children.Add(presetPanel);

            var body = new Grid();
            body.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            body.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetRow(body, 2);
            root.Children.Add(body);

            var rowsGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = false,
                SelectionMode = DataGridSelectionMode.Single,
                SelectionUnit = DataGridSelectionUnit.FullRow,
                ItemsSource = rows
            };
            rowsGrid.Columns.Add(new DataGridTextColumn { Header = "№", Binding = new Binding(nameof(ColumnManagerRow.Order)), IsReadOnly = true, Width = 48 });
            rowsGrid.Columns.Add(new DataGridCheckBoxColumn { Header = "Показать", Binding = new Binding(nameof(ColumnManagerRow.IsVisible)), Width = 88 });
            rowsGrid.Columns.Add(new DataGridTextColumn { Header = "Колонка", Binding = new Binding(nameof(ColumnManagerRow.Header)), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            rowsGrid.Columns.Add(new DataGridTextColumn { Header = "Ширина", Binding = new Binding(nameof(ColumnManagerRow.WidthText)) { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 110 });
            Grid.SetColumn(rowsGrid, 0);
            body.Children.Add(rowsGrid);

            var actions = new StackPanel
            {
                Margin = new Thickness(10, 0, 0, 0),
                Width = 132
            };
            Grid.SetColumn(actions, 1);
            body.Children.Add(actions);

            void RefreshOrder()
            {
                for (var i = 0; i < rows.Count; i++)
                    rows[i].Order = i + 1;
            }

            var moveUpButton = new Button { Content = "Вверх", MinWidth = 120, Style = FindResource("SecondaryButton") as Style, Margin = new Thickness(0, 0, 0, 8) };
            moveUpButton.Click += (_, _) =>
            {
                if (rowsGrid.SelectedItem is not ColumnManagerRow selected)
                    return;

                var index = rows.IndexOf(selected);
                if (index <= 0)
                    return;

                rows.Move(index, index - 1);
                RefreshOrder();
                rowsGrid.SelectedItem = selected;
            };
            actions.Children.Add(moveUpButton);

            var moveDownButton = new Button { Content = "Вниз", MinWidth = 120, Style = FindResource("SecondaryButton") as Style };
            moveDownButton.Click += (_, _) =>
            {
                if (rowsGrid.SelectedItem is not ColumnManagerRow selected)
                    return;

                var index = rows.IndexOf(selected);
                if (index < 0 || index >= rows.Count - 1)
                    return;

                rows.Move(index, index + 1);
                RefreshOrder();
                rowsGrid.SelectedItem = selected;
            };
            actions.Children.Add(moveDownButton);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            Grid.SetRow(footer, 3);
            root.Children.Add(footer);

            var cancelButton = new Button { Content = "Отмена", MinWidth = 120, Style = FindResource("SecondaryButton") as Style, IsCancel = true, Margin = new Thickness(0, 0, 8, 0) };
            var applyButton = new Button { Content = "Применить", MinWidth = 130, IsDefault = true };
            footer.Children.Add(cancelButton);
            footer.Children.Add(applyButton);

            applyButton.Click += (_, _) =>
            {
                var columnQueues = grid.Columns
                    .GroupBy(x => GetColumnHeaderKey(x), StringComparer.CurrentCultureIgnoreCase)
                    .ToDictionary(
                        x => x.Key,
                        x => new Queue<DataGridColumn>(x.OrderBy(c => c.DisplayIndex)),
                        StringComparer.CurrentCultureIgnoreCase);

                isApplyingColumnPreferences = true;
                try
                {
                    var nextDisplayIndex = 0;
                    foreach (var row in rows)
                    {
                        if (!columnQueues.TryGetValue(row.Header, out var queue) || queue.Count == 0)
                            continue;

                        var column = queue.Dequeue();
                        column.Visibility = row.IsVisible ? Visibility.Visible : Visibility.Collapsed;
                        if (TryParsePositiveWidth(row.WidthText, out var width))
                            column.Width = new DataGridLength(width);
                        column.DisplayIndex = nextDisplayIndex++;
                    }
                }
                finally
                {
                    isApplyingColumnPreferences = false;
                }

                SaveGridColumnPreferences(grid);
                SaveState(SaveTrigger.System);
                dialog.DialogResult = true;
            };

            void RefreshPresetList(string selectName = null)
            {
                var names = new ObservableCollection<string>(GetGridPresetNames(grid));
                presetCombo.ItemsSource = names;
                if (!string.IsNullOrWhiteSpace(selectName) && names.Contains(selectName))
                    presetCombo.SelectedItem = selectName;
                else if (names.Count > 0)
                    presetCombo.SelectedIndex = 0;
            }

            savePresetButton.Click += (_, _) =>
            {
                var name = presetNameBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    MessageBox.Show("Введите название пресета.");
                    return;
                }

                SaveGridColumnPreset(grid, name);
                RefreshPresetList(name);
                SaveState(SaveTrigger.System);
                SetLastOperationStatus($"Пресет колонок сохранен: {name}");
            };

            applyPresetButton.Click += (_, _) =>
            {
                if (presetCombo.SelectedItem is not string presetName || string.IsNullOrWhiteSpace(presetName))
                {
                    MessageBox.Show("Выберите пресет.");
                    return;
                }

                ApplyGridColumnPreset(grid, presetName);
                SaveState(SaveTrigger.System);
                SetLastOperationStatus($"Пресет колонок применен: {presetName}");
            };

            deletePresetButton.Click += (_, _) =>
            {
                if (presetCombo.SelectedItem is not string presetName || string.IsNullOrWhiteSpace(presetName))
                {
                    MessageBox.Show("Выберите пресет.");
                    return;
                }

                if (MessageBox.Show($"Удалить пресет \"{presetName}\"?", "Пресеты колонок", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                    return;

                RemoveGridColumnPreset(grid, presetName);
                RefreshPresetList();
                SaveState(SaveTrigger.System);
                SetLastOperationStatus($"Пресет колонок удален: {presetName}");
            };

            dialog.Content = root;
            dialog.ShowDialog();
        }

        private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e == null)
                return;

            if (ReferenceEquals(MainTabs?.SelectedItem, EstimateTab)
                && EstimateExcelHost?.Visibility == Visibility.Visible
                && estimateExcelWindowHandle != IntPtr.Zero)
            {
                return;
            }

            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (TryHandleCommandPaletteShortcut(key, Keyboard.Modifiers, IsTextEditingElement(Keyboard.FocusedElement)))
            {
                e.Handled = true;
            }
        }

        private void OpenGlobalSearch_Click(object sender, RoutedEventArgs e)
        {
            var initialQuery = GlobalSearchTextBox?.Text?.Trim() ?? string.Empty;
            ShowGlobalSearchDialog(initialQuery);
        }

        private void GlobalSearchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            OpenGlobalSearch_Click(sender, new RoutedEventArgs());
            e.Handled = true;
        }

        private void DigitsOnlyTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !DigitsInputRegex.IsMatch(e.Text ?? string.Empty);
        }

        private void ProductionBlocksBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (ProductionBlocksBox == null)
                return;

            var normalized = NormalizeProductionBlocksText(ProductionBlocksBox.Text);
            if (!string.Equals(ProductionBlocksBox.Text, normalized, StringComparison.CurrentCulture))
                ProductionBlocksBox.Text = normalized;
        }

        private void ProductionMarksBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (ProductionMarksBox == null)
                return;

            var normalized = NormalizeProductionMarksText(ProductionMarksBox.Text);
            if (!string.Equals(ProductionMarksBox.Text, normalized, StringComparison.CurrentCulture))
                ProductionMarksBox.Text = normalized;
        }

        private static string NormalizeProductionBlocksText(string text)
        {
            var blocks = LevelMarkHelper.ParseBlocks(text)
                .Where(x => x > 0)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            return blocks.Count == 0 ? string.Empty : string.Join(", ", blocks);
        }

        private static string NormalizeProductionMarksText(string text)
        {
            var marks = LevelMarkHelper.ParseMarks(text)
                .Select(NormalizeProductionMarkToken)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            return marks.Count == 0 ? string.Empty : string.Join(", ", marks);
        }

        private static string NormalizeProductionMarkToken(string token)
        {
            var normalized = (token ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
                return string.Empty;

            var numericCandidate = normalized.Replace(',', '.');
            if (!Regex.IsMatch(numericCandidate, @"^[+-]?\d+(\.\d+)?$"))
                return normalized;

            if (!double.TryParse(numericCandidate, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
                return normalized;

            if (Math.Abs(value) < 0.0000001)
                return "0.000";

            var signPrefix = value > 0 ? "+" : string.Empty;
            return $"{signPrefix}{value:0.000}";
        }

        private void OpenCommandPalette_Click(object sender, RoutedEventArgs e)
        {
            if (isOpeningCommandDialog)
                return;

            isOpeningCommandDialog = true;
            try
            {
                ShowCommandPaletteDialog();
            }
            finally
            {
                isOpeningCommandDialog = false;
            }
        }

        private void ShowGlobalSearchDialog(string initialQuery)
        {
            var dialog = new Window
            {
                Title = "Глобальный поиск",
                Owner = this,
                Width = 880,
                Height = 620,
                MinWidth = 760,
                MinHeight = 520,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.CanResize
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var queryBox = new TextBox
            {
                MinHeight = 36,
                Tag = "Введите материал, ФИО, ТТН или дату",
                Text = (initialQuery ?? string.Empty).Trim()
            };
            Grid.SetRow(queryBox, 0);
            root.Children.Add(queryBox);

            var resultRows = new ObservableCollection<GlobalSearchResult>();
            var resultGrid = new DataGrid
            {
                Margin = new Thickness(0, 10, 0, 0),
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                SelectionMode = DataGridSelectionMode.Single,
                SelectionUnit = DataGridSelectionUnit.FullRow,
                ItemsSource = resultRows
            };
            resultGrid.Columns.Add(new DataGridTextColumn { Header = "Вкладка", Binding = new Binding(nameof(GlobalSearchResult.TabHeader)), Width = 130 });
            resultGrid.Columns.Add(new DataGridTextColumn { Header = "Запись", Binding = new Binding(nameof(GlobalSearchResult.Title)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            resultGrid.Columns.Add(new DataGridTextColumn { Header = "Описание", Binding = new Binding(nameof(GlobalSearchResult.Description)), Width = new DataGridLength(1.3, DataGridLengthUnitType.Star) });
            Grid.SetRow(resultGrid, 1);
            root.Children.Add(resultGrid);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            var closeButton = new Button
            {
                Content = "Закрыть",
                MinWidth = 120,
                IsCancel = true,
                Style = FindResource("SecondaryButton") as Style,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var openButton = new Button
            {
                Content = "Перейти",
                MinWidth = 120,
                IsDefault = true
            };
            footer.Children.Add(closeButton);
            footer.Children.Add(openButton);

            void ApplySearch()
            {
                var query = queryBox.Text?.Trim() ?? string.Empty;
                var found = BuildGlobalSearchResults(query).Take(500).ToList();
                resultRows.Clear();
                foreach (var row in found)
                    resultRows.Add(row);

                if (resultRows.Count > 0)
                    resultGrid.SelectedIndex = 0;
            }

            void NavigateSelected()
            {
                if (resultGrid.SelectedItem is not GlobalSearchResult selected || selected.NavigateAction == null)
                    return;

                dialog.DialogResult = true;
                dialog.Close();
                selected.NavigateAction();
            }

            queryBox.TextChanged += (_, _) => ApplySearch();
            queryBox.KeyDown += (_, args) =>
            {
                if (args.Key == Key.Enter)
                {
                    if (resultRows.Count == 0)
                    {
                        ApplySearch();
                        if (resultRows.Count == 0)
                        {
                            args.Handled = true;
                            return;
                        }
                    }

                    NavigateSelected();
                    args.Handled = true;
                }
                else if (args.Key == Key.Down)
                {
                    if (resultGrid.Items.Count > 0)
                    {
                        var nextIndex = Math.Min(resultGrid.Items.Count - 1, Math.Max(0, resultGrid.SelectedIndex + 1));
                        resultGrid.SelectedIndex = nextIndex;
                        resultGrid.ScrollIntoView(resultGrid.SelectedItem);
                        args.Handled = true;
                    }
                }
                else if (args.Key == Key.Up)
                {
                    if (resultGrid.Items.Count > 0)
                    {
                        var nextIndex = Math.Max(0, resultGrid.SelectedIndex - 1);
                        resultGrid.SelectedIndex = nextIndex;
                        resultGrid.ScrollIntoView(resultGrid.SelectedItem);
                        args.Handled = true;
                    }
                }
            };

            resultGrid.MouseDoubleClick += (_, _) => NavigateSelected();
            resultGrid.KeyDown += (_, args) =>
            {
                if (args.Key == Key.Enter)
                {
                    NavigateSelected();
                    args.Handled = true;
                }
            };

            openButton.Click += (_, _) => NavigateSelected();

            dialog.Content = root;
            ApplySearch();
            dialog.ShowDialog();
        }

        private List<GlobalSearchResult> BuildGlobalSearchResults(string query)
        {
            var normalizedQuery = (query ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedQuery))
                return new List<GlobalSearchResult>();

            var tokens = normalizedQuery
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            var results = new List<GlobalSearchResult>();
            bool Match(params string[] values)
            {
                var haystack = string.Join(" ", (values ?? Array.Empty<string>()).Where(x => !string.IsNullOrWhiteSpace(x)));
                if (string.IsNullOrWhiteSpace(haystack))
                    return false;

                return tokens.All(token => haystack.IndexOf(token, StringComparison.CurrentCultureIgnoreCase) >= 0);
            }

            foreach (var row in journal)
            {
                if (!Match(
                    row.MaterialGroup,
                    row.MaterialName,
                    row.Ttn,
                    row.Supplier,
                    row.Passport,
                    row.Stb,
                    row.Date.ToString("dd.MM.yyyy")))
                {
                    continue;
                }

                results.Add(new GlobalSearchResult
                {
                    TabHeader = "Приход",
                    Title = $"{row.MaterialGroup} / {row.MaterialName}",
                    Description = $"{row.Date:dd.MM.yyyy}; ТТН: {row.Ttn}; Кол-во: {FormatNumberByUnit(row.Quantity, row.Unit)}",
                    NavigateAction = () => NavigateToArrivalRecord(row)
                });
            }

            if (currentObject?.OtJournal != null)
            {
                foreach (var row in currentObject.OtJournal)
                {
                    if (!Match(
                        row.FullName,
                        row.Specialty,
                        row.Profession,
                        row.InstructionType,
                        row.InstructionNumbers,
                        row.StatusLabel,
                        row.InstructionDate.ToString("dd.MM.yyyy")))
                    {
                        continue;
                    }

                    results.Add(new GlobalSearchResult
                    {
                        TabHeader = "ОТ",
                        Title = row.FullName ?? "Без ФИО",
                        Description = $"{row.InstructionType}; профессия: {row.Profession}; дата: {row.InstructionDate:dd.MM.yyyy}",
                        NavigateAction = () => NavigateToOtRow(row)
                    });
                }
            }

            if (currentObject?.TimesheetPeople != null)
            {
                foreach (var row in currentObject.TimesheetPeople)
                {
                    if (!Match(row.FullName, row.Specialty, row.Rank, row.BrigadeName, row.DailyWorkHours.ToString()))
                        continue;

                    results.Add(new GlobalSearchResult
                    {
                        TabHeader = "Табель",
                        Title = row.FullName ?? "Без ФИО",
                        Description = $"{row.Specialty}; разряд: {row.Rank}; часов/день: {row.DailyWorkHours}",
                        NavigateAction = () => NavigateToTimesheetPerson(row)
                    });
                }
            }

            if (currentObject?.ProductionJournal != null)
            {
                foreach (var row in currentObject.ProductionJournal)
                {
                    if (!Match(
                        row.ActionName,
                        row.WorkName,
                        row.ElementsText,
                        row.BlocksText,
                        row.MarksText,
                        row.BrigadeName,
                        row.Weather,
                        row.Deviations,
                        row.Date.ToString("dd.MM.yyyy")))
                    {
                        continue;
                    }

                    results.Add(new GlobalSearchResult
                    {
                        TabHeader = "ПР",
                        Title = $"{row.Date:dd.MM.yyyy} | {row.ActionName} {row.WorkName}",
                        Description = $"{row.ElementsText}; блоки: {row.BlocksText}; отметки: {row.MarksText}",
                        NavigateAction = () => NavigateToProductionRow(row)
                    });
                }
            }

            if (currentObject?.InspectionJournal != null)
            {
                foreach (var row in currentObject.InspectionJournal)
                {
                    if (!Match(
                        row.JournalName,
                        row.InspectionName,
                        row.ReminderStatus,
                        row.ReminderStartDate.ToString("dd.MM.yyyy"),
                        row.NextReminderDate.ToString("dd.MM.yyyy")))
                    {
                        continue;
                    }

                    results.Add(new GlobalSearchResult
                    {
                        TabHeader = "Осмотры",
                        Title = $"{row.JournalName} — {row.InspectionName}",
                        Description = $"Напоминать с {row.ReminderStartDate:dd.MM.yyyy}; период: {row.ReminderPeriodDays} дн.",
                        NavigateAction = () => NavigateToInspectionRow(row)
                    });
                }
            }

            if (currentObject?.Demand != null)
            {
                foreach (var pair in currentObject.Demand)
                {
                    var parts = (pair.Key ?? string.Empty).Split(new[] { "::" }, StringSplitOptions.None);
                    var group = parts.Length > 0 ? parts[0] : string.Empty;
                    var material = parts.Length > 1 ? parts[1] : string.Empty;
                    if (!Match(group, material, pair.Value?.Unit))
                        continue;

                    results.Add(new GlobalSearchResult
                    {
                        TabHeader = "Сводка",
                        Title = $"{group} / {material}",
                        Description = $"Ед.: {pair.Value?.Unit}",
                        NavigateAction = () => SelectMainTab(SummaryTab)
                    });
                }
            }

            return results
                .OrderBy(x => x.TabHeader, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.Title, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private void NavigateToArrivalRecord(JournalRecord row)
        {
            if (row == null)
                return;

            SelectMainTab(ArrivalTab);
            if (arrivalMatrixMode)
            {
                arrivalMatrixMode = false;
                SetTabDisplayMode("Приход", "Таблица");
                UpdateArrivalViewMode();
            }

            filteredJournal = journal.ToList();
            if (ArrivalLegacyGrid != null)
                ArrivalLegacyGrid.ItemsSource = filteredJournal;

            Dispatcher.BeginInvoke(new Action(() => SelectGridItem(ArrivalLegacyGrid, row)), DispatcherPriority.Background);
        }

        private void NavigateToOtRow(OtJournalEntry row)
        {
            if (row == null)
                return;

            SelectMainTab(OtTab);
            Dispatcher.BeginInvoke(new Action(() => SelectGridItem(OtJournalGrid, row)), DispatcherPriority.Background);
        }

        private void NavigateToTimesheetPerson(TimesheetPersonEntry row)
        {
            if (row == null)
                return;

            SelectMainTab(TimesheetTab);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                EnsureTabInitialized(TimesheetTab);
                var target = timesheetRows.FirstOrDefault(x => x.PersonId == row.PersonId)
                    ?? timesheetRows.FirstOrDefault(x => string.Equals(x.FullName, row.FullName, StringComparison.CurrentCultureIgnoreCase));
                if (target != null)
                    SelectGridItem(TimesheetGrid, target);
            }), DispatcherPriority.Background);
        }

        private void NavigateToProductionRow(ProductionJournalEntry row)
        {
            if (row == null)
                return;

            SelectMainTab(ProductionTab);
            Dispatcher.BeginInvoke(new Action(() => SelectGridItem(ProductionJournalGrid, row)), DispatcherPriority.Background);
        }

        private void NavigateToInspectionRow(InspectionJournalEntry row)
        {
            if (row == null)
                return;

            SelectMainTab(InspectionTab);
            Dispatcher.BeginInvoke(new Action(() => SelectGridItem(InspectionJournalGrid, row)), DispatcherPriority.Background);
        }

        private static void SelectGridItem(DataGrid grid, object item)
        {
            if (grid == null || item == null)
                return;

            if (!grid.Items.Contains(item))
                return;

            grid.SelectedItem = item;
            grid.ScrollIntoView(item);
            if (grid.SelectedCells.Count == 0 && grid.Columns.Count > 0)
                grid.SelectedCells.Add(new DataGridCellInfo(item, grid.Columns[0]));
            grid.Focus();
        }

        private void ShowCommandPaletteDialog()
        {
            var actions = BuildCommandPaletteActions();
            var visibleActions = new ObservableCollection<CommandPaletteAction>(actions);

            var dialog = new Window
            {
                Title = "Командная палитра",
                Owner = this,
                Width = 760,
                Height = 560,
                MinWidth = 700,
                MinHeight = 480,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.CanResize
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var filterBox = new TextBox
            {
                MinHeight = 36,
                Tag = "Введите действие или клавиши, например: сохранить, вкладка, Ctrl+K"
            };
            Grid.SetRow(filterBox, 0);
            root.Children.Add(filterBox);

            var shortcutHintBorder = new Border
            {
                Margin = new Thickness(0, 10, 0, 0),
                Padding = new Thickness(12, 10, 12, 10),
                BorderBrush = TryFindResource("StrokeBrush") as Brush ?? Brushes.Gainsboro,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Background = TryFindResource("SurfaceAltBrush") as Brush ?? Brushes.WhiteSmoke
            };
            Grid.SetRow(shortcutHintBorder, 1);
            root.Children.Add(shortcutHintBorder);

            var shortcutHintText = new TextBlock
            {
                Text = "Выберите команду и нажмите «Назначить клавиши». Потом нажмите нужную комбинацию. Esc отменяет ввод, Delete очищает комбинацию.",
                TextWrapping = TextWrapping.Wrap
            };
            shortcutHintBorder.Child = shortcutHintText;

            var list = new DataGrid
            {
                Margin = new Thickness(0, 10, 0, 0),
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                SelectionMode = DataGridSelectionMode.Single,
                SelectionUnit = DataGridSelectionUnit.FullRow,
                ItemsSource = visibleActions
            };
            list.Columns.Add(new DataGridTextColumn { Header = "Команда", Binding = new Binding(nameof(CommandPaletteAction.Name)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            list.Columns.Add(new DataGridTextColumn { Header = "Клавиши", Binding = new Binding(nameof(CommandPaletteAction.Shortcut)) { TargetNullValue = string.Empty }, Width = 140 });
            list.Columns.Add(new DataGridTextColumn { Header = "Описание", Binding = new Binding(nameof(CommandPaletteAction.Hint)), Width = new DataGridLength(1.2, DataGridLengthUnitType.Star) });
            Grid.SetRow(list, 2);
            root.Children.Add(list);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            Grid.SetRow(footer, 3);
            root.Children.Add(footer);

            var closeButton = new Button
            {
                Content = "Закрыть",
                MinWidth = 120,
                IsCancel = true,
                Style = FindResource("SecondaryButton") as Style,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var clearShortcutButton = new Button
            {
                Content = "Очистить клавиши",
                MinWidth = 150,
                Style = FindResource("SecondaryButton") as Style,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var assignShortcutButton = new Button
            {
                Content = "Назначить клавиши",
                MinWidth = 160,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var runButton = new Button
            {
                Content = "Выполнить",
                MinWidth = 120,
                IsDefault = true
            };
            footer.Children.Add(closeButton);
            footer.Children.Add(clearShortcutButton);
            footer.Children.Add(assignShortcutButton);
            footer.Children.Add(runButton);

            var isCapturingShortcut = false;

            void SetCaptureState(bool isCapturing, string message = null)
            {
                isCapturingShortcut = isCapturing;
                shortcutHintText.Text = isCapturing
                    ? (message ?? "Нажмите новую комбинацию. Esc отменяет ввод, Delete очищает назначение.")
                    : "Выберите команду и нажмите «Назначить клавиши». Потом нажмите нужную комбинацию. Esc отменяет ввод, Delete очищает комбинацию.";
                assignShortcutButton.Content = isCapturing ? "Ожидание..." : "Назначить клавиши";
                assignShortcutButton.IsEnabled = !isCapturing;
                runButton.IsEnabled = !isCapturing;
                filterBox.IsReadOnly = isCapturing;
            }

            void ApplyFilter()
            {
                var text = filterBox.Text?.Trim() ?? string.Empty;
                var tokens = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                var filtered = actions.Where(action =>
                {
                    if (tokens.Length == 0)
                        return true;

                    var source = $"{action.Name} {action.Shortcut} {action.Hint}";
                    return tokens.All(token => source.IndexOf(token, StringComparison.CurrentCultureIgnoreCase) >= 0);
                }).ToList();

                visibleActions.Clear();
                foreach (var action in filtered)
                    visibleActions.Add(action);

                if (visibleActions.Count > 0)
                    list.SelectedIndex = 0;
            }

            void ExecuteSelectedAction()
            {
                if (isCapturingShortcut)
                    return;

                if (list.SelectedItem is not CommandPaletteAction selected || selected.ExecuteAction == null)
                    return;

                dialog.DialogResult = true;
                dialog.Close();
                selected.ExecuteAction();
            }

            void PersistShortcut(CommandPaletteAction action, string shortcut)
            {
                EnsureProjectUiSettings();
                if (currentObject?.UiSettings == null)
                {
                    MessageBox.Show("Сначала создайте или откройте объект, чтобы сохранить комбинацию.");
                    return;
                }

                var normalizedShortcut = NormalizeShortcutText(shortcut);
                var normalizedDefault = NormalizeShortcutText(action.DefaultShortcut);
                currentObject.UiSettings.CommandPaletteShortcuts ??= new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);

                if (string.Equals(normalizedShortcut, normalizedDefault, StringComparison.CurrentCultureIgnoreCase))
                    currentObject.UiSettings.CommandPaletteShortcuts.Remove(action.Id);
                else
                    currentObject.UiSettings.CommandPaletteShortcuts[action.Id] = normalizedShortcut;

                action.Shortcut = normalizedShortcut;
                SaveState(SaveTrigger.System);
            }

            void ClearShortcut(CommandPaletteAction action)
            {
                if (action == null)
                    return;

                PersistShortcut(action, string.Empty);
                SetCaptureState(false);
            }

            void AssignShortcut(CommandPaletteAction action, string shortcut)
            {
                if (action == null)
                    return;

                var normalizedShortcut = NormalizeShortcutText(shortcut);
                if (string.IsNullOrWhiteSpace(normalizedShortcut))
                    return;

                var conflict = actions.FirstOrDefault(x =>
                    !ReferenceEquals(x, action)
                    && string.Equals(NormalizeShortcutText(x.Shortcut), normalizedShortcut, StringComparison.CurrentCultureIgnoreCase));

                if (conflict != null)
                {
                    var replaceResult = MessageBox.Show(
                        $"Комбинация {normalizedShortcut} уже назначена команде \"{conflict.Name}\". Переназначить?",
                        "Командная палитра",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (replaceResult != MessageBoxResult.Yes)
                    {
                        SetCaptureState(false);
                        return;
                    }

                    PersistShortcut(conflict, string.Empty);
                }

                PersistShortcut(action, normalizedShortcut);
                SetCaptureState(false);
            }

            filterBox.TextChanged += (_, _) => ApplyFilter();
            filterBox.KeyDown += (_, args) =>
            {
                if (isCapturingShortcut)
                {
                    args.Handled = true;
                    return;
                }

                if (args.Key == Key.Enter)
                {
                    ExecuteSelectedAction();
                    args.Handled = true;
                    return;
                }

                if (args.Key == Key.Down)
                {
                    if (list.Items.Count > 0)
                    {
                        list.SelectedIndex = Math.Min(list.Items.Count - 1, Math.Max(0, list.SelectedIndex + 1));
                        list.ScrollIntoView(list.SelectedItem);
                        args.Handled = true;
                    }
                    return;
                }

                if (args.Key == Key.Up)
                {
                    if (list.Items.Count > 0)
                    {
                        list.SelectedIndex = Math.Max(0, list.SelectedIndex - 1);
                        list.ScrollIntoView(list.SelectedItem);
                        args.Handled = true;
                    }
                }
            };

            list.MouseDoubleClick += (_, _) => ExecuteSelectedAction();
            list.KeyDown += (_, args) =>
            {
                if (isCapturingShortcut)
                {
                    args.Handled = true;
                    return;
                }

                if (args.Key == Key.Enter)
                {
                    ExecuteSelectedAction();
                    args.Handled = true;
                }
            };
            clearShortcutButton.Click += (_, _) =>
            {
                if (list.SelectedItem is not CommandPaletteAction selected)
                {
                    MessageBox.Show("Выберите команду.");
                    return;
                }

                ClearShortcut(selected);
            };
            assignShortcutButton.Click += (_, _) =>
            {
                if (list.SelectedItem is not CommandPaletteAction selected)
                {
                    MessageBox.Show("Выберите команду.");
                    return;
                }

                SetCaptureState(true, $"Новая комбинация для команды \"{selected.Name}\". Нажмите клавиши, Esc отменяет ввод, Delete очищает назначение.");
            };
            runButton.Click += (_, _) => ExecuteSelectedAction();
            dialog.PreviewKeyDown += (_, args) =>
            {
                if (!isCapturingShortcut)
                    return;

                var shortcutKey = args.Key == Key.System ? args.SystemKey : args.Key;
                if (IsModifierKey(shortcutKey))
                {
                    args.Handled = true;
                    return;
                }

                if (shortcutKey == Key.Escape)
                {
                    SetCaptureState(false);
                    args.Handled = true;
                    return;
                }

                if (shortcutKey == Key.Delete && Keyboard.Modifiers == ModifierKeys.None)
                {
                    if (list.SelectedItem is CommandPaletteAction selectedAction)
                        ClearShortcut(selectedAction);

                    args.Handled = true;
                    return;
                }

                if (list.SelectedItem is not CommandPaletteAction selectedActionToAssign)
                {
                    SetCaptureState(false);
                    args.Handled = true;
                    return;
                }

                if (Keyboard.Modifiers == ModifierKeys.None && (shortcutKey < Key.F1 || shortcutKey > Key.F24))
                {
                    SetCaptureState(true, "Используйте сочетание с Ctrl, Shift, Alt или функциональную клавишу.");
                    args.Handled = true;
                    return;
                }

                AssignShortcut(selectedActionToAssign, BuildShortcutText(shortcutKey, Keyboard.Modifiers));
                args.Handled = true;
            };

            dialog.Content = root;
            ApplyFilter();
            SetCaptureState(false);
            dialog.ShowDialog();
        }

        private List<CommandPaletteAction> BuildCommandPaletteActions()
        {
            return new List<CommandPaletteAction>
            {
                new()
                {
                    Id = "command_palette",
                    DefaultShortcut = "Ctrl+K",
                    Shortcut = ResolveCommandPaletteShortcut("command_palette", "Ctrl+K"),
                    Name = "Командная палитра",
                    Hint = "Открыть список быстрых действий и назначить клавиши",
                    ExecuteAction = () => OpenCommandPalette_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "save",
                    DefaultShortcut = "Ctrl+S",
                    Shortcut = ResolveCommandPaletteShortcut("save", "Ctrl+S"),
                    Name = "Сохранить",
                    Hint = "Сохранить текущий объект",
                    ExecuteAction = () => SaveButton_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "save_as",
                    DefaultShortcut = "Ctrl+Shift+S",
                    Shortcut = ResolveCommandPaletteShortcut("save_as", "Ctrl+Shift+S"),
                    Name = "Сохранить как",
                    Hint = "Сохранить объект в новый файл",
                    ExecuteAction = () => SaveAs_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "undo",
                    DefaultShortcut = "Ctrl+Z",
                    Shortcut = ResolveCommandPaletteShortcut("undo", "Ctrl+Z"),
                    Name = "Отменить действие",
                    Hint = "Вернуться на один шаг назад",
                    ExecuteAction = () => Undo_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "redo",
                    DefaultShortcut = "Ctrl+Y",
                    Shortcut = ResolveCommandPaletteShortcut("redo", "Ctrl+Y"),
                    Name = "Повторить действие",
                    Hint = "Вернуть последний отмененный шаг",
                    ExecuteAction = () => Redo_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "refresh",
                    DefaultShortcut = "F5",
                    Shortcut = ResolveCommandPaletteShortcut("refresh", "F5"),
                    Name = "Обновить",
                    Hint = "Пересчитать и обновить текущие данные",
                    ExecuteAction = () => RefreshButton_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "global_search",
                    DefaultShortcut = "Ctrl+F",
                    Shortcut = ResolveCommandPaletteShortcut("global_search", "Ctrl+F"),
                    Name = "Глобальный поиск",
                    Hint = "Поиск по материалам, ФИО, ТТН, датам",
                    ExecuteAction = () => OpenGlobalSearch_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "column_manager",
                    DefaultShortcut = "Ctrl+Shift+M",
                    Shortcut = ResolveCommandPaletteShortcut("column_manager", "Ctrl+Shift+M"),
                    Name = "Менеджер колонок",
                    Hint = "Показать/скрыть колонки, изменить порядок и ширину",
                    ExecuteAction = () => OpenColumnManager_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "ui_settings",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("ui_settings", string.Empty),
                    Name = "Настройки интерфейса",
                    Hint = "Открыть параметры интерфейса",
                    ExecuteAction = () => AppSettings_Click(this, new RoutedEventArgs())
                },
                new()
                {
                    Id = "density_standard",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("density_standard", string.Empty),
                    Name = "Плотность: Стандартный",
                    Hint = "Обычные отступы и высота строк",
                    ExecuteAction = () => SetUiDensityMode("Стандартный")
                },
                new()
                {
                    Id = "density_compact",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("density_compact", string.Empty),
                    Name = "Плотность: Компактный",
                    Hint = "Уплотненный режим для больших таблиц",
                    ExecuteAction = () => SetUiDensityMode("Компактный")
                },
                new()
                {
                    Id = "tab_summary",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_summary", string.Empty),
                    Name = "Открыть вкладку Сводка",
                    Hint = "Перейти на вкладку Сводка",
                    ExecuteAction = () => SelectMainTab(SummaryTab)
                },
                new()
                {
                    Id = "tab_jvk",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_jvk", string.Empty),
                    Name = "Открыть вкладку ЖВК",
                    Hint = "Перейти на вкладку ЖВК",
                    ExecuteAction = () => SelectMainTab(JvkTab)
                },
                new()
                {
                    Id = "tab_arrival",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_arrival", string.Empty),
                    Name = "Открыть вкладку Приход",
                    Hint = "Перейти на вкладку Приход",
                    ExecuteAction = () => SelectMainTab(ArrivalTab)
                },
                new()
                {
                    Id = "tab_ot",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_ot", string.Empty),
                    Name = "Открыть вкладку ОТ",
                    Hint = "Перейти на вкладку ОТ",
                    ExecuteAction = () => SelectMainTab(OtTab)
                },
                new()
                {
                    Id = "tab_timesheet",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_timesheet", string.Empty),
                    Name = "Открыть вкладку Табель",
                    Hint = "Перейти на вкладку Табель",
                    ExecuteAction = () => SelectMainTab(TimesheetTab)
                },
                new()
                {
                    Id = "tab_production",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_production", string.Empty),
                    Name = "Открыть вкладку ПР",
                    Hint = "Перейти на вкладку ПР",
                    ExecuteAction = () => SelectMainTab(ProductionTab)
                },
                new()
                {
                    Id = "tab_inspection",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_inspection", string.Empty),
                    Name = "Открыть вкладку Осмотры",
                    Hint = "Перейти на вкладку Осмотры",
                    ExecuteAction = () => SelectMainTab(InspectionTab)
                },
                new()
                {
                    Id = "tab_pdf",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_pdf", string.Empty),
                    Name = "Открыть вкладку ПДФ",
                    Hint = "Перейти на вкладку ПДФ-документы",
                    ExecuteAction = () => SelectMainTab(PdfTab)
                },
                new()
                {
                    Id = "tab_estimate",
                    DefaultShortcut = string.Empty,
                    Shortcut = ResolveCommandPaletteShortcut("tab_estimate", string.Empty),
                    Name = "Открыть вкладку Сметы",
                    Hint = "Перейти на вкладку Сметы",
                    ExecuteAction = () => SelectMainTab(EstimateTab)
                }
            };
        }

        private string ResolveCommandPaletteShortcut(string actionId, string defaultShortcut)
        {
            EnsureProjectUiSettings();
            if (currentObject?.UiSettings?.CommandPaletteShortcuts != null
                && !string.IsNullOrWhiteSpace(actionId)
                && currentObject.UiSettings.CommandPaletteShortcuts.TryGetValue(actionId, out var savedShortcut))
            {
                return NormalizeShortcutText(savedShortcut);
            }

            return NormalizeShortcutText(defaultShortcut);
        }

        private bool TryHandleCommandPaletteShortcut(Key key, ModifierKeys modifiers, bool textEditingActive)
        {
            var shortcut = BuildShortcutText(key, modifiers);
            if (string.IsNullOrWhiteSpace(shortcut))
                return false;

            var action = BuildCommandPaletteActions().FirstOrDefault(x =>
                string.Equals(NormalizeShortcutText(x.Shortcut), shortcut, StringComparison.CurrentCultureIgnoreCase));
            if (action == null)
                return false;

            if (textEditingActive && ShouldSkipShortcutWhileEditingText(action.Id))
                return false;

            action.ExecuteAction?.Invoke();
            return true;
        }

        private static bool ShouldSkipShortcutWhileEditingText(string actionId)
        {
            return string.Equals(actionId, "undo", StringComparison.CurrentCultureIgnoreCase)
                || string.Equals(actionId, "redo", StringComparison.CurrentCultureIgnoreCase);
        }

        private static bool ShouldHandleGridRowDeleteShortcut(KeyEventArgs e)
        {
            if (e == null)
                return false;

            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (key != Key.Delete || Keyboard.Modifiers != ModifierKeys.None)
                return false;

            return !IsTextEditingElement(e.OriginalSource);
        }

        private static bool IsTextEditingElement(object source)
        {
            var current = source as DependencyObject;
            while (current != null)
            {
                if (current is TextBoxBase || current is PasswordBox || current is RichTextBox || current is ComboBox || current is DatePicker)
                    return true;

                current = GetVisualOrLogicalParent(current);
            }

            return false;
        }

        private static DependencyObject GetVisualOrLogicalParent(DependencyObject current)
        {
            if (current == null)
                return null;

            if (current is Visual)
                return VisualTreeHelper.GetParent(current) ?? LogicalTreeHelper.GetParent(current);

            return LogicalTreeHelper.GetParent(current);
        }

        private static bool IsModifierKey(Key key)
        {
            return key == Key.LeftCtrl
                || key == Key.RightCtrl
                || key == Key.LeftAlt
                || key == Key.RightAlt
                || key == Key.LeftShift
                || key == Key.RightShift
                || key == Key.LWin
                || key == Key.RWin;
        }

        private static string BuildShortcutText(Key key, ModifierKeys modifiers)
        {
            if (key == Key.None || IsModifierKey(key))
                return string.Empty;

            var parts = new List<string>();
            if (modifiers.HasFlag(ModifierKeys.Control))
                parts.Add("Ctrl");
            if (modifiers.HasFlag(ModifierKeys.Shift))
                parts.Add("Shift");
            if (modifiers.HasFlag(ModifierKeys.Alt))
                parts.Add("Alt");
            if (modifiers.HasFlag(ModifierKeys.Windows))
                parts.Add("Win");

            var keyText = GetShortcutKeyText(key);
            if (string.IsNullOrWhiteSpace(keyText))
                return string.Empty;

            parts.Add(keyText);
            return string.Join("+", parts);
        }

        private static string NormalizeShortcutText(string shortcut)
        {
            if (string.IsNullOrWhiteSpace(shortcut))
                return string.Empty;

            var parts = shortcut
                .Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            if (parts.Count == 0)
                return string.Empty;

            var hasCtrl = false;
            var hasShift = false;
            var hasAlt = false;
            var hasWin = false;
            var keyPart = string.Empty;
            foreach (var part in parts)
            {
                if (part.Equals("ctrl", StringComparison.CurrentCultureIgnoreCase)
                    || part.Equals("control", StringComparison.CurrentCultureIgnoreCase))
                {
                    hasCtrl = true;
                    continue;
                }

                if (part.Equals("shift", StringComparison.CurrentCultureIgnoreCase))
                {
                    hasShift = true;
                    continue;
                }

                if (part.Equals("alt", StringComparison.CurrentCultureIgnoreCase))
                {
                    hasAlt = true;
                    continue;
                }

                if (part.Equals("win", StringComparison.CurrentCultureIgnoreCase)
                    || part.Equals("windows", StringComparison.CurrentCultureIgnoreCase))
                {
                    hasWin = true;
                    continue;
                }

                keyPart = part;
            }

            var normalized = new List<string>();
            if (hasCtrl)
                normalized.Add("Ctrl");
            if (hasShift)
                normalized.Add("Shift");
            if (hasAlt)
                normalized.Add("Alt");
            if (hasWin)
                normalized.Add("Win");

            if (string.IsNullOrWhiteSpace(keyPart))
                return string.Join("+", normalized);

            normalized.Add(keyPart);
            return string.Join("+", normalized);
        }

        private static string GetShortcutKeyText(Key key)
        {
            if (key >= Key.D0 && key <= Key.D9)
                return ((int)(key - Key.D0)).ToString(CultureInfo.InvariantCulture);

            if (key >= Key.NumPad0 && key <= Key.NumPad9)
                return $"Num{(int)(key - Key.NumPad0)}";

            return key switch
            {
                Key.Return => "Enter",
                Key.Escape => "Esc",
                Key.Prior => "PageUp",
                Key.Next => "PageDown",
                Key.Back => "Backspace",
                Key.Space => "Space",
                _ => new KeyConverter().ConvertToString(key) ?? string.Empty
            };
        }

        private void SetUiDensityMode(string mode)
        {
            EnsureProjectUiSettings();
            if (currentObject?.UiSettings == null)
                return;

            currentObject.UiSettings.UiDensityMode = mode ?? "Стандартный";
            ApplyUiDensityMode();
            SaveState(SaveTrigger.System);
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
        private void EnsureArrivalFilterTemplatesStorage()
        {
            currentObject ??= new ProjectObject();
            currentObject.ArrivalFilterTemplates ??= new List<ArrivalFilterTemplate>();
        }

        private void RefreshArrivalFilterTemplates()
        {
            arrivalFilterTemplateNames.Clear();
            EnsureArrivalFilterTemplatesStorage();

            foreach (var name in currentObject.ArrivalFilterTemplates
                         .Where(x => !string.IsNullOrWhiteSpace(x.Name))
                         .Select(x => x.Name.Trim())
                         .Distinct(StringComparer.CurrentCultureIgnoreCase)
                         .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
            {
                arrivalFilterTemplateNames.Add(name);
            }

            if (ArrivalTemplateBox != null)
            {
                suppressArrivalTemplateSelectionChange = true;
                ArrivalTemplateBox.ItemsSource = arrivalFilterTemplateNames;
                if (arrivalFilterTemplateNames.Count == 0)
                    ArrivalTemplateBox.SelectedItem = null;
                suppressArrivalTemplateSelectionChange = false;
            }
        }

        private void SaveArrivalFilterTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект.");
                return;
            }

            EnsureArrivalFilterTemplatesStorage();
            var name = Microsoft.VisualBasic.Interaction.InputBox(
                "Введите название шаблона фильтра:",
                "Сохранить шаблон",
                ArrivalTemplateBox?.Text?.Trim() ?? string.Empty)?.Trim();

            if (string.IsNullOrWhiteSpace(name))
                return;

            var template = BuildCurrentArrivalFilterTemplate(name);
            var existing = currentObject.ArrivalFilterTemplates.FirstOrDefault(x =>
                string.Equals(x.Name, name, StringComparison.CurrentCultureIgnoreCase));

            if (existing != null)
                currentObject.ArrivalFilterTemplates.Remove(existing);

            currentObject.ArrivalFilterTemplates.Add(template);
            RefreshArrivalFilterTemplates();
            suppressArrivalTemplateSelectionChange = true;
            if (ArrivalTemplateBox != null)
                ArrivalTemplateBox.SelectedItem = name;
            suppressArrivalTemplateSelectionChange = false;
            SaveState();
        }

        private ArrivalFilterTemplate BuildCurrentArrivalFilterTemplate(string name)
        {
            return new ArrivalFilterTemplate
            {
                Name = name?.Trim() ?? string.Empty,
                SelectedTypes = selectedArrivalTypes.OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase).ToList(),
                SelectedNames = selectedArrivalNames.OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase).ToList(),
                ShowMain = ArrivalMainCheck?.IsChecked == true,
                ShowExtra = ArrivalExtraCheck?.IsChecked == true,
                ShowLowCost = ArrivalLowCostCheck?.IsChecked == true,
                ShowInternal = ArrivalInternalCheck?.IsChecked == true,
                DateFrom = ArrivalDateFrom?.SelectedDate,
                DateTo = ArrivalDateTo?.SelectedDate,
                SearchText = ArrivalSearchBox?.Text?.Trim() ?? string.Empty
            };
        }

        private void ArrivalTemplateBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (suppressArrivalTemplateSelectionChange || currentObject?.ArrivalFilterTemplates == null)
                return;

            var name = ArrivalTemplateBox?.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(name))
                return;

            var template = currentObject.ArrivalFilterTemplates.FirstOrDefault(x =>
                string.Equals(x.Name, name, StringComparison.CurrentCultureIgnoreCase));
            if (template == null)
                return;

            ApplyArrivalFilterTemplate(template);
            RequestArrivalFilterRefresh(immediate: true);
        }

        private void ApplyArrivalFilterTemplate(ArrivalFilterTemplate template)
        {
            if (template == null)
                return;

            selectedArrivalTypes = template.SelectedTypes?.Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase) ?? new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            selectedArrivalNames = template.SelectedNames?.Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase) ?? new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

            if (ArrivalMainCheck != null) ArrivalMainCheck.IsChecked = template.ShowMain;
            if (ArrivalExtraCheck != null) ArrivalExtraCheck.IsChecked = template.ShowExtra;
            if (ArrivalLowCostCheck != null) ArrivalLowCostCheck.IsChecked = template.ShowLowCost;
            if (ArrivalInternalCheck != null) ArrivalInternalCheck.IsChecked = template.ShowInternal;
            if (ArrivalDateFrom != null) ArrivalDateFrom.SelectedDate = template.DateFrom;
            if (ArrivalDateTo != null) ArrivalDateTo.SelectedDate = template.DateTo;
            if (ArrivalSearchBox != null) ArrivalSearchBox.Text = template.SearchText ?? string.Empty;

            RefreshArrivalTypes();
            RefreshArrivalNames();
        }

        private void DeleteArrivalFilterTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject?.ArrivalFilterTemplates == null || currentObject.ArrivalFilterTemplates.Count == 0)
                return;

            var name = ArrivalTemplateBox?.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("Выберите шаблон для удаления.");
                return;
            }

            var existing = currentObject.ArrivalFilterTemplates.FirstOrDefault(x =>
                string.Equals(x.Name, name, StringComparison.CurrentCultureIgnoreCase));
            if (existing == null)
                return;

            currentObject.ArrivalFilterTemplates.Remove(existing);
            RefreshArrivalFilterTemplates();
            SaveState();
        }

        private void BatchEditArrivalRows_Click(object sender, RoutedEventArgs e)
        {
            if (ArrivalLegacyGrid == null)
                return;

            var selectedRows = ArrivalLegacyGrid.SelectedItems?.OfType<JournalRecord>().ToList() ?? new List<JournalRecord>();
            if (selectedRows.Count == 0)
            {
                MessageBox.Show("Выберите строки в таблице прихода для пакетного изменения.");
                return;
            }

            var typeOptions = journal
                .Select(x => x.MaterialGroup?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            var supplierOptions = journal
                .Select(x => x.Supplier?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var dialog = new Window
            {
                Title = $"Пакетное редактирование ({selectedRows.Count} строк)",
                Owner = this,
                Width = 520,
                Height = 330,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(14) };
            for (var i = 0; i < 4; i++)
                root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            dialog.Content = root;

            var applyDateCheck = new CheckBox { Content = "Изменить дату", Margin = new Thickness(0, 0, 0, 6) };
            var datePicker = new DatePicker { IsEnabled = false, SelectedDate = DateTime.Today };
            applyDateCheck.Checked += (_, _) => datePicker.IsEnabled = true;
            applyDateCheck.Unchecked += (_, _) => datePicker.IsEnabled = false;
            Grid.SetRow(applyDateCheck, 0);
            Grid.SetRow(datePicker, 1);
            root.Children.Add(applyDateCheck);
            root.Children.Add(datePicker);

            var typePanel = new StackPanel { Margin = new Thickness(0, 8, 0, 0) };
            typePanel.Children.Add(new TextBlock { Text = "Новый тип (если нужно)", FontWeight = FontWeights.SemiBold });
            var typeBox = new ComboBox
            {
                IsEditable = true,
                IsTextSearchEnabled = true,
                StaysOpenOnEdit = true,
                ItemsSource = typeOptions
            };
            typePanel.Children.Add(typeBox);
            Grid.SetRow(typePanel, 2);
            root.Children.Add(typePanel);

            var supplierPanel = new StackPanel { Margin = new Thickness(0, 8, 0, 0) };
            supplierPanel.Children.Add(new TextBlock { Text = "Новый поставщик (если нужно)", FontWeight = FontWeights.SemiBold });
            var supplierBox = new ComboBox
            {
                IsEditable = true,
                IsTextSearchEnabled = true,
                StaysOpenOnEdit = true,
                ItemsSource = supplierOptions
            };
            supplierPanel.Children.Add(supplierBox);
            Grid.SetRow(supplierPanel, 3);
            root.Children.Add(supplierPanel);

            var hint = new TextBlock
            {
                Text = "Пустое поле означает: не менять это значение.",
                Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                Margin = new Thickness(0, 10, 0, 0)
            };
            Grid.SetRow(hint, 4);
            root.Children.Add(hint);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            Grid.SetRow(footer, 5);
            root.Children.Add(footer);

            var applyButton = new Button { Content = "Применить", MinWidth = 120 };
            var cancelButton = new Button { Content = "Отмена", MinWidth = 110, Margin = new Thickness(8, 0, 0, 0), IsCancel = true };
            footer.Children.Add(applyButton);
            footer.Children.Add(cancelButton);

            applyButton.Click += (_, _) =>
            {
                var newType = typeBox.Text?.Trim();
                var newSupplier = supplierBox.Text?.Trim();
                var shouldUpdateDate = applyDateCheck.IsChecked == true && datePicker.SelectedDate.HasValue;

                foreach (var row in selectedRows)
                {
                    if (shouldUpdateDate)
                        row.Date = datePicker.SelectedDate.Value.Date;
                    if (!string.IsNullOrWhiteSpace(newType))
                        row.MaterialGroup = newType;
                    if (!string.IsNullOrWhiteSpace(newSupplier))
                        row.Supplier = newSupplier;
                }

                dialog.DialogResult = true;
                dialog.Close();
            };

            if (dialog.ShowDialog() != true)
                return;

            RebuildArchiveFromCurrentData();
            RefreshJournalAnomalies();
            RefreshArrivalTypes();
            RefreshArrivalNames();
            ApplyAllFilters();
            SaveState();
        }

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
                    RequestArrivalFilterRefresh();
                };
                chip.Unchecked += (_, _) =>
                {
                    if (selectedArrivalTypes.Contains(g))
                        selectedArrivalTypes.Remove(g);
                    if (arrivalMatrixMode)
                        selectedArrivalNames.Clear();
                    RefreshArrivalNames();
                    RequestArrivalFilterRefresh();
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
                    RequestArrivalFilterRefresh();
                };
                chip.Unchecked += (_, _) =>
                {
                    selectedArrivalNames.Remove(n);
                    RequestArrivalFilterRefresh();
                };

                ArrivalNamesPanel.Children.Add(chip);
            }
        }
        private void ArrivalSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RequestArrivalFilterRefresh();
        }
        private void ExportArrival_Click(object sender, RoutedEventArgs e)
        {
            if (!filteredJournal.Any())
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            var exportAsMatrix = AskArrivalExportMatrixMode();
            if (exportAsMatrix == null)
                return;

            var dlg = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = exportAsMatrix.Value ? "Приход_матрица.xlsx" : "Приход_таблица.xlsx"
            };

            if (dlg.ShowDialog() != true)
                return;

            try
            {
                using var wb = new XLWorkbook();
                if (exportAsMatrix.Value)
                    ExportArrivalMatrix(wb);
                else
                    ExportArrival(wb);
                wb.SaveAs(dlg.FileName);
                MessageBox.Show("Экспорт завершён");
            }
            catch (IOException)
            {
                MessageBox.Show("Не удалось сохранить файл. Закройте этот файл в Excel/PlanMaker и повторите экспорт.");
            }
        }

        private void ExportSelectedRangeToExcel_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetActiveGridForRangeExport();
            if (grid == null)
            {
                MessageBox.Show("На текущей вкладке нет таблицы для экспорта выделения.");
                return;
            }

            if (!TryBuildSelectedRangeData(grid, out var headers, out var rows))
            {
                MessageBox.Show("Сначала выделите диапазон или строки в таблице.");
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = $"Выделение_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (dialog.ShowDialog() != true)
                return;

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Выделение");
            for (var c = 0; c < headers.Count; c++)
                ws.Cell(1, c + 1).Value = headers[c];
            ws.Range(1, 1, 1, headers.Count).Style.Font.Bold = true;
            ws.Range(1, 1, 1, headers.Count).Style.Fill.BackgroundColor = XLColor.FromHtml("#EEF2F7");

            for (var r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                for (var c = 0; c < row.Count; c++)
                    ws.Cell(r + 2, c + 1).Value = row[c];
            }

            ws.Columns().AdjustToContents();
            ws.Range(1, 1, Math.Max(1, rows.Count + 1), headers.Count).SetAutoFilter();
            wb.SaveAs(dialog.FileName);
            MessageBox.Show("Экспорт выделенного в Excel выполнен.");
        }

        private void ExportSelectedRangeToWord_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetActiveGridForRangeExport();
            if (grid == null)
            {
                MessageBox.Show("На текущей вкладке нет таблицы для экспорта выделения.");
                return;
            }

            if (!TryBuildSelectedRangeData(grid, out var headers, out var rows))
            {
                MessageBox.Show("Сначала выделите диапазон или строки в таблице.");
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "Word (*.docx)|*.docx",
                FileName = $"Выделение_{DateTime.Now:yyyyMMdd_HHmm}.docx"
            };
            if (dialog.ShowDialog() != true)
                return;

            using var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(
                dialog.FileName,
                DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            var body = new DocumentFormat.OpenXml.Wordprocessing.Body();
            mainPart.Document.Append(body);

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("Экспорт выделенного диапазона"))));
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text($"Дата: {DateTime.Now:dd.MM.yyyy HH:mm}"))));
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text(string.Empty))));

            foreach (var row in rows)
            {
                var pieces = headers.Zip(row, (h, v) => $"{h}: {v}");
                body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text($"• {string.Join("; ", pieces)}"))));
            }

            mainPart.Document.Save();
            MessageBox.Show("Экспорт выделенного в Word выполнен.");
        }

        private DataGrid GetActiveGridForRangeExport()
        {
            if (ReferenceEquals(MainTabs?.SelectedItem, ArrivalTab))
                return arrivalMatrixMode ? null : ArrivalLegacyGrid;
            if (ReferenceEquals(MainTabs?.SelectedItem, OtTab))
                return OtJournalGrid;
            if (ReferenceEquals(MainTabs?.SelectedItem, TimesheetTab))
                return TimesheetGrid;
            if (ReferenceEquals(MainTabs?.SelectedItem, ProductionTab))
                return ProductionJournalGrid;
            if (ReferenceEquals(MainTabs?.SelectedItem, InspectionTab))
                return InspectionJournalGrid;
            return null;
        }

        private bool TryBuildSelectedRangeData(DataGrid grid, out List<string> headers, out List<List<string>> rows)
        {
            headers = new List<string>();
            rows = new List<List<string>>();
            if (grid == null)
                return false;

            var selectedCells = grid.SelectedCells?.ToList() ?? new List<DataGridCellInfo>();
            var selectedItems = grid.SelectedItems?.Cast<object>().ToList() ?? new List<object>();

            var targetColumns = new List<DataGridColumn>();
            var targetItems = new List<object>();

            if (selectedCells.Count > 0)
            {
                targetColumns = selectedCells
                    .Select(x => x.Column)
                    .Where(x => x != null)
                    .Distinct()
                    .OrderBy(x => x.DisplayIndex)
                    .ToList();

                targetItems = selectedCells
                    .Select(x => x.Item)
                    .Where(x => x != null)
                    .Distinct()
                    .OrderBy(x => grid.Items.IndexOf(x))
                    .ToList();
            }
            else if (selectedItems.Count > 0)
            {
                targetItems = selectedItems.OrderBy(x => grid.Items.IndexOf(x)).ToList();
                targetColumns = grid.Columns
                    .Where(x => x.Visibility == Visibility.Visible)
                    .OrderBy(x => x.DisplayIndex)
                    .ToList();
            }
            else if (grid.CurrentItem != null)
            {
                targetItems.Add(grid.CurrentItem);
                targetColumns = grid.Columns
                    .Where(x => x.Visibility == Visibility.Visible)
                    .OrderBy(x => x.DisplayIndex)
                    .ToList();
            }

            if (targetItems.Count == 0 || targetColumns.Count == 0)
                return false;

            headers = targetColumns
                .Select(x => x.Header?.ToString() ?? string.Empty)
                .ToList();

            foreach (var item in targetItems)
            {
                var line = new List<string>();
                foreach (var column in targetColumns)
                    line.Add(ExtractGridCellText(column, item));
                rows.Add(line);
            }

            return rows.Count > 0;
        }

        private string ExtractGridCellText(DataGridColumn column, object item)
        {
            if (column == null || item == null)
                return string.Empty;

            if (column is DataGridBoundColumn boundColumn
                && boundColumn.Binding is Binding binding
                && !string.IsNullOrWhiteSpace(binding.Path?.Path))
            {
                var path = binding.Path.Path.Trim();
                var property = item.GetType().GetProperty(path);
                var value = property?.GetValue(item);
                return value?.ToString() ?? string.Empty;
            }

            if (item is TimesheetRowViewModel timesheetRow
                && int.TryParse(column.Header?.ToString(), out var day)
                && day > 0)
            {
                return timesheetRow.GetDayValue(day);
            }

            var cellContent = column.GetCellContent(item);
            if (cellContent is TextBlock textBlock)
                return textBlock.Text ?? string.Empty;
            if (cellContent is CheckBox checkBox)
                return checkBox.IsChecked == true ? "Да" : "Нет";

            return string.Empty;
        }

        private bool IsPdfEditorProcessAlive()
        {
            if (pdfEditorProcess == null)
                return false;

            try
            {
                return !pdfEditorProcess.HasExited;
            }
            catch
            {
                return false;
            }
        }

        private bool EnsurePdfExternalProcess()
        {
            if (!useExternalPdfEditor || string.IsNullOrWhiteSpace(preferredPdfEditorPath) || !File.Exists(preferredPdfEditorPath))
                return false;

            if (IsPdfEditorProcessAlive() && pdfEditorWindowHandle != IntPtr.Zero)
                return true;

            Process process;
            try
            {
                process = Process.Start(new ProcessStartInfo
                {
                    FileName = preferredPdfEditorPath,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Minimized
                }) ?? throw new InvalidOperationException("Не удалось запустить PDF-XChange Editor.");
            }
            catch
            {
                return false;
            }

            try
            {
                process.WaitForInputIdle(5000);
            }
            catch
            {
                // ignore input idle errors
            }

            var handle = WaitForMainWindowHandle(process, timeoutMs: 25000);
            if (handle == IntPtr.Zero)
            {
                try
                {
                    if (!process.HasExited)
                        process.Kill();
                }
                catch
                {
                    // ignore kill errors
                }

                return false;
            }

            pdfEditorProcess = process;
            pdfEditorWindowHandle = handle;
            ConfigureFloatingEstimateWindow(pdfEditorWindowHandle);
            ResetPdfEmbeddedLayoutCache();
            SchedulePdfEmbeddedLayout();
            return true;
        }

        private void OpenPdfInExternalEditor(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return;

            if (!EnsurePdfExternalProcess())
                return;

            if (string.Equals(pdfEmbeddedFilePath, filePath, StringComparison.CurrentCultureIgnoreCase))
            {
                SchedulePdfEmbeddedLayout();
                return;
            }

            pdfEmbeddedFilePath = filePath;
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = preferredPdfEditorPath,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false
                });
            }
            catch
            {
                // Ignore open errors, keep the embedded window alive.
            }

            try { ShowWindow(pdfEditorWindowHandle, SW_SHOW); } catch { }
            SchedulePdfEmbeddedLayout();
        }

        private void ResetPdfEmbeddedLayoutCache()
        {
            pdfEmbeddedWindowX = int.MinValue;
            pdfEmbeddedWindowY = int.MinValue;
            pdfEmbeddedWindowWidth = -1;
            pdfEmbeddedWindowHeight = -1;
        }

        private void SchedulePdfEmbeddedLayout()
        {
            if (pdfEditorWindowHandle == IntPtr.Zero)
                return;

            LayoutEmbeddedPdfWindow(force: true);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (pdfEditorWindowHandle != IntPtr.Zero)
                    LayoutEmbeddedPdfWindow();
            }), DispatcherPriority.Render);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (pdfEditorWindowHandle != IntPtr.Zero)
                    LayoutEmbeddedPdfWindow();
            }), DispatcherPriority.ApplicationIdle);
        }

        private void LayoutEmbeddedPdfWindow(bool force = false)
        {
            if (pdfEditorWindowHandle == IntPtr.Zero)
                return;

            if (!IsVisible
                || WindowState == WindowState.Minimized
                || !ReferenceEquals(MainTabs?.SelectedItem, PdfTab)
                || PdfPreviewContainer == null
                || PdfPreviewContainer.Visibility != Visibility.Visible)
            {
                try { ShowWindow(pdfEditorWindowHandle, SW_HIDE); } catch { }
                return;
            }

            PdfPreviewContainer.UpdateLayout();
            var screenBounds = GetScreenBounds(PdfPreviewContainer);
            var width = screenBounds.Width;
            var height = screenBounds.Height;
            if (width <= 0 || height <= 0)
            {
                try { ShowWindow(pdfEditorWindowHandle, SW_HIDE); } catch { }
                return;
            }

            var x = screenBounds.X;
            var y = screenBounds.Y;

            if (!force
                && x == pdfEmbeddedWindowX
                && y == pdfEmbeddedWindowY
                && width == pdfEmbeddedWindowWidth
                && height == pdfEmbeddedWindowHeight)
            {
                return;
            }

            pdfEmbeddedWindowX = x;
            pdfEmbeddedWindowY = y;
            pdfEmbeddedWindowWidth = width;
            pdfEmbeddedWindowHeight = height;

            SetWindowPos(
                pdfEditorWindowHandle,
                IntPtr.Zero,
                x,
                y,
                width,
                height,
                SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_SHOWWINDOW);

            SetWindowRgn(pdfEditorWindowHandle, IntPtr.Zero, true);
        }

        private void HidePdfEmbeddedPreview()
        {
            if (pdfEditorWindowHandle == IntPtr.Zero)
                return;

            try { ShowWindow(pdfEditorWindowHandle, SW_HIDE); } catch { }
        }

        private void ClosePdfExternalProcess()
        {
            if (pdfEditorProcess != null)
            {
                try
                {
                    if (!pdfEditorProcess.HasExited)
                    {
                        pdfEditorProcess.CloseMainWindow();
                        if (!pdfEditorProcess.WaitForExit(1500))
                            pdfEditorProcess.Kill();
                    }
                }
                catch
                {
                    // Ignore shutdown errors.
                }
                finally
                {
                    pdfEditorProcess.Dispose();
                    pdfEditorProcess = null;
                }
            }

            pdfEditorWindowHandle = IntPtr.Zero;
            pdfEmbeddedFilePath = string.Empty;
            ResetPdfEmbeddedLayoutCache();
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

        private bool? AskArrivalExportMatrixMode()
        {
            var result = MessageBox.Show(
                "Выберите формат экспорта:\nДа — Таблица\nНет — Матрица",
                "Экспорт прихода",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            return result switch
            {
                MessageBoxResult.Yes => false,
                MessageBoxResult.No => true,
                _ => null
            };
        }

        private void ExportArrivalMatrix(IXLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Матрица");
            var data = filteredJournal
                .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName))
                .ToList();

            if (data.Count == 0)
            {
                ws.Cell(1, 1).Value = "Нет данных по выбранным фильтрам.";
                ws.Columns().AdjustToContents();
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
                .Select(x => x.Key)
                .OrderBy(x => x.Date)
                .ThenBy(x => x.Ttn, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var materials = data
                .Select(x => (x.MaterialName ?? string.Empty).Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var unitByMaterial = data
                .GroupBy(x => (x.MaterialName ?? string.Empty).Trim(), StringComparer.CurrentCultureIgnoreCase)
                .ToDictionary(
                    x => x.Key,
                    x => x.Select(y => y.Unit ?? string.Empty).FirstOrDefault() ?? string.Empty,
                    StringComparer.CurrentCultureIgnoreCase);

            ws.Cell(1, 1).Value = "Наименование";
            ws.Cell(1, 2).Value = "Ед.";
            ws.Range(1, 1, 1, 2).Style.Font.Bold = true;
            ws.Range(1, 1, 1, 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#EEF2F7");

            for (var i = 0; i < columns.Count; i++)
            {
                var col = columns[i];
                var excelCol = i + 3;
                ws.Cell(1, excelCol).Value = $"{col.Date:dd.MM.yyyy}\n{(string.IsNullOrWhiteSpace(col.Ttn) ? "—" : col.Ttn)}";
                ws.Cell(2, excelCol).Value = string.IsNullOrWhiteSpace(col.Supplier) ? "—" : col.Supplier;
                ws.Cell(3, excelCol).Value = string.IsNullOrWhiteSpace(col.Passport) ? "—" : col.Passport;
                ws.Cell(1, excelCol).Style.Alignment.WrapText = true;
                ws.Cell(2, excelCol).Style.Alignment.WrapText = true;
                ws.Cell(3, excelCol).Style.Alignment.WrapText = true;
            }

            ws.Range(1, 3, 3, columns.Count + 2).Style.Font.Bold = true;
            ws.Range(1, 3, 3, columns.Count + 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#EEF4FF");

            for (var materialIndex = 0; materialIndex < materials.Count; materialIndex++)
            {
                var material = materials[materialIndex];
                var row = materialIndex + 4;
                ws.Cell(row, 1).Value = material;
                ws.Cell(row, 2).Value = unitByMaterial.TryGetValue(material, out var unit) ? unit : string.Empty;

                for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
                {
                    var col = columns[columnIndex];
                    var quantity = data
                        .Where(x => string.Equals((x.MaterialName ?? string.Empty).Trim(), material, StringComparison.CurrentCultureIgnoreCase)
                            && x.Date.Date == col.Date
                            && string.Equals((x.Ttn ?? string.Empty).Trim(), col.Ttn, StringComparison.CurrentCultureIgnoreCase)
                            && string.Equals((x.Supplier ?? string.Empty).Trim(), col.Supplier, StringComparison.CurrentCultureIgnoreCase)
                            && string.Equals((x.Passport ?? string.Empty).Trim(), col.Passport, StringComparison.CurrentCultureIgnoreCase))
                        .Sum(x => x.Quantity);

                    if (Math.Abs(quantity) > 0.0001)
                        ws.Cell(row, columnIndex + 3).Value = quantity;
                }
            }

            var totalRows = materials.Count + 3;
            var totalCols = columns.Count + 2;
            ws.Range(1, 1, totalRows, totalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, totalRows, totalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Columns().AdjustToContents();
            ws.Rows(1, 3).Height = 28;
            ws.SheetView.FreezeRows(3);
            ws.SheetView.FreezeColumns(2);
            ws.Range(1, 1, 3, totalCols).SetAutoFilter();
        }




    }

}
