using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
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
    "#D9E9FF", "#FFE2D9", "#E4FFD9", "#F5E9FF", "#FFFACC",
    "#D9FFF8", "#FFD9F2", "#E8D9FF", "#FFD9D9", "#D9FFEA"
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

        private const string SaveFileName = "data.json";
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
        private bool summaryFilterInitialized;
        private bool summaryFilterUpdating;
        private List<string> summaryFilterGroups = new();
        private bool summaryHasOverage;
        private TextBlock summaryOverageNote;
        private readonly ObservableCollection<string> brigadierNames = new();
        private readonly ObservableCollection<string> specialties = new();
        private readonly ObservableCollection<string> professions = new();
        private string otSearchText = string.Empty;
        private bool isTreePinned;
        private Point treeDragStart;

        private sealed class TreeNodeMeta
        {
            public string Kind { get; set; }
            public string MaterialName { get; set; }
            public string GroupName { get; set; }
            public string SubCategory { get; set; }
            public string Category { get; set; }
            public List<string> PrefixSegments { get; set; }
        }

        public ObservableCollection<string> BrigadierNames => brigadierNames;
        public ObservableCollection<string> Specialties => specialties;
        public ObservableCollection<string> Professions => professions;
        public MainWindow()
        {
            InitializeComponent();
            ArrivalLiveTable.IsReadOnly = true;
            ArrivalLiveTable.CanUserAddRows = false;
            ArrivalLiveTable.CanUserDeleteRows = false;


            // ===== БЛОКИРОВКА ВКЛЮЧЕНА ПО УМОЛЧАНИЮ =====
            isLocked = true;

            LoadState();
            InitializeOtJournal();
            ApplyAllFilters();

            ArrivalPanel.ArrivalAdded += OnArrivalAdded;

            PushUndo();
            UpdateUndoRedoButtons();

            if (currentObject != null)
                ArrivalPanel.SetObject(currentObject, journal);
            RefreshArrivalTypes();
            RefreshArrivalNames();

            RefreshTreePreserveState();
            UpdateTreePanelState(forceVisible: false);

        }
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // гарантированно после создания всех контролов

        }
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source is not TabControl tab || tab.SelectedItem is not TabItem item)
                return;

            if (item.Header?.ToString() == "Приход")
            {
                ArrivalPopup.Visibility = arrivalPanelVisible
                    ? Visibility.Visible
                    : Visibility.Collapsed;

                ShowArrivalButton.Visibility = arrivalPanelVisible
                    ? Visibility.Collapsed
                    : Visibility.Visible;

            }
            else
            {
                ArrivalPopup.Visibility = Visibility.Collapsed;
                ShowArrivalButton.Visibility = Visibility.Visible;
            }

            UpdateOtReminders();
            OtJournalGrid.Items.Refresh();
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
            UpdateOtReminders();
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

            if (e.PropertyName == nameof(OtJournalEntry.IsBrigadier) || e.PropertyName == nameof(OtJournalEntry.FullName))
                RefreshBrigadierNames();
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
                || e.PropertyName == nameof(OtJournalEntry.IsBrigadier)
                || e.PropertyName == nameof(OtJournalEntry.BrigadierName)
                
                || e.PropertyName == nameof(OtJournalEntry.IsDismissed))
            {
                NormalizeOtRows();
                SortOtJournal();
                UpdateOtReminders();
            }
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
                    .Where(x => !x.IsPendingRepeat)
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
                .ThenBy(x => GetCrewSortKey(x), StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.IsBrigadier ? 0 : 1)
                .ThenBy(x => x.FullName ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var view = CollectionViewSource.GetDefaultView(currentObject.OtJournal);
            view.Filter = OtJournalFilter;
            OtJournalGrid.ItemsSource = view;
            view.Refresh();
        }

        private static string GetCrewSortKey(OtJournalEntry row)
        {
            if (row.IsBrigadier)
                return $"00_{row.FullName}";

            if (string.IsNullOrWhiteSpace(row.BrigadierName))
                return $"99_{row.FullName}";

            return $"10_{row.BrigadierName}";
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
        private void UpdateOtReminders()
        {
            if (currentObject?.OtJournal == null)
            {
                OtReminderPopup.Visibility = Visibility.Collapsed;
                return;
            }

            var dueRows = currentObject.OtJournal.Where(x => x.IsPendingRepeat && !x.IsDismissed).ToList();
            if (dueRows.Count > 0)
            {
                OtReminderText.Text = $"Требуется повторный инструктаж: {dueRows.Count} чел.{Environment.NewLine}" +
                        string.Join(Environment.NewLine, dueRows.Take(5).Select(x => $"• {x.LastName}"));
                OtReminderPopup.Visibility = Visibility.Visible;
            }
            else
            {
                OtReminderPopup.Visibility = Visibility.Collapsed;
            }
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
            UpdateOtReminders();
            SaveState();
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
            row.IsRepeatCompleted = true;

            NormalizeOtRows();
            SortOtJournal();
            UpdateOtReminders();
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
                item.IsDismissed = true;

            UpdateOtReminders();
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
            UpdateOtReminders();
            
            SaveState();
        }
        private void OtSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            otSearchText = OtSearchTextBox.Text?.Trim() ?? string.Empty;
            CollectionViewSource.GetDefaultView(OtJournalGrid.ItemsSource)?.Refresh();
        }

        private void OtJournalGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                RefreshBrigadierNames();
                RefreshSpecialties();
                RefreshProfessions();
                NormalizeOtRows();

                SortOtJournal();
                UpdateOtReminders();
                SaveState();
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
                UpdateOtReminders();
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
                UpdateOtReminders();
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
                summaryFilterInitialized = false;
                EnsureOtJournalStorage();
                BindOtJournal();
                ArrivalPanel.SetObject(currentObject, journal);



                SaveState();
                RefreshTreePreserveState();
              
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
        private void TreeSettings_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            var materialNames = journal
                .Where(j => !string.IsNullOrWhiteSpace(j.MaterialName))
                .Select(j => new TreeSettingsWindow.MaterialSplitRuleSource
                {
                    MaterialName = j.MaterialName,
                    
                    TypeName = GetSegmentsForMaterial(j.MaterialName).FirstOrDefault() ?? j.MaterialName
                })
                .ToList();

            var w = new TreeSettingsWindow(materialNames, currentObject.MaterialTreeSplitRules ?? new())
            {
                Owner = this
            };

            if (w.ShowDialog() != true)
                return;

            currentObject.MaterialTreeSplitRules = w.ResultRules;
            SaveState();
            RefreshTreePreserveState();
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

            if (ArrivalLiveTable != null)
            {
                ArrivalLiveTable.IsReadOnly = true;
                ArrivalLiveTable.CanUserAddRows = false;
                ArrivalLiveTable.CanUserDeleteRows = false;
            }
            RefreshSummaryTable();
        }


        private void LockButton_Unchecked(object sender, RoutedEventArgs e)
        {
            isLocked = false;

            if (ArrivalLiveTable != null)
            {
                ArrivalLiveTable.IsReadOnly = false;
                ArrivalLiveTable.CanUserAddRows = true;
                ArrivalLiveTable.CanUserDeleteRows = true;
            }
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



            // ===== ТОЛЬКО ДЛЯ ОСНОВНЫХ =====
            if (arrival.Category == "Основные")
            {
                if (!currentObject.MaterialGroups.Any(g => g.Name == arrival.MaterialGroup))
                {
                    currentObject.MaterialGroups.Add(new MaterialGroup
                    {
                        Name = arrival.MaterialGroup
                    });

                    currentObject.MaterialNamesByGroup[arrival.MaterialGroup] = new List<string>();
                }
            }

            foreach (var i in arrival.Items)
            {
                if (arrival.Category == "Основные")
                {
                    // === список на дереве ===
                    if (!currentObject.MaterialNamesByGroup[arrival.MaterialGroup]
                            .Contains(i.MaterialName))
                    {
                        currentObject.MaterialNamesByGroup[arrival.MaterialGroup]
                            .Add(i.MaterialName);
                    }

                    // === список для ComboBox ===
                    var archive = currentObject.Archive;

                    if (!archive.Materials.ContainsKey(arrival.MaterialGroup))
                        archive.Materials[arrival.MaterialGroup] = new();

                    if (!archive.Materials[arrival.MaterialGroup]
                            .Contains(i.MaterialName))
                    {
                        archive.Materials[arrival.MaterialGroup].Add(i.MaterialName);
                    }
                }

                // === запись журнала ===
                journal.Add(new JournalRecord
                {
                    Date = i.Date,
                    ObjectName = currentObject.Name,
                    Category = arrival.Category,
                    SubCategory = arrival.SubCategory,
                    MaterialGroup = arrival.MaterialGroup,
                    MaterialName = i.MaterialName,
                    Unit = i.Unit,
                    Quantity = i.Quantity,
                    Passport = i.Passport,
                    Ttn = arrival.TtnNumber,
                    Stb = i.Stb,
                    Supplier = i.Supplier
                });
            }




            SaveState();
            RefreshTreePreserveState();
   

            // важно: обновляем панель прихода
            ArrivalPanel.SetObject(currentObject, journal);
            // === обновляем чипы типов и материалов ===
            RefreshArrivalTypes();
            RefreshArrivalNames();


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

                foreach (var m in g.Select(x => x.MaterialName)
                                  .Distinct()
                                  .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase))
                {
                    AddMaterialTreeNodes(groupNode, m, g.Key, "Основные", null);
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

            return Regex.Matches(materialName, @"[A-Za-zА-Яа-яЁё]+|\d+")
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
            return GetSegmentsFromText(materialName);
        }
        private void ObjectsTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            ApplyAllFilters();
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

                    
                        CleanupMaterialsAfterDelete();
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

                    CleanupMaterialsAfterDelete();
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
            }

            if (GetNodeKind(node) == "Material")
            {
                var oldMaterialName = node.Tag is TreeNodeMeta meta ? meta.MaterialName : oldName;
                foreach (var kv in currentObject.MaterialNamesByGroup)
                {
                    var idx = kv.Value.IndexOf(oldMaterialName);
                    if (idx >= 0)
                        kv.Value[idx] = input;
                }

                foreach (var j in journal.Where(x => x.MaterialName == oldMaterialName))
                    j.MaterialName = input;
                if (currentObject.MaterialTreeSplitRules.TryGetValue(oldMaterialName, out var rule))
                {
                    currentObject.MaterialTreeSplitRules[input] = rule;
                    currentObject.MaterialTreeSplitRules.Remove(oldMaterialName);
                }
            }

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
            }

            if (GetNodeKind(node) == "Material")
            {
                var materialName = node.Tag is TreeNodeMeta meta ? meta.MaterialName : name;
                foreach (var kv in currentObject.MaterialNamesByGroup)
                kv.Value.Remove(materialName);

                journal.RemoveAll(j => j.MaterialName == materialName);
                currentObject.MaterialTreeSplitRules.Remove(materialName);
            }

            SaveState();
            RefreshTreePreserveState();
          
        }

        private void ArrivalFilters_Changed(object sender, RoutedEventArgs e)
        {


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





            // === ПРИХОД: СОРТ ПО УМОЛЧАНИЮ ===
            data = data.OrderByDescending(j => j.Date);


            filteredJournal = data.ToList();


            RenderJvk();
            RefreshSummaryTable();

            if (ArrivalLiveTable != null)
                ArrivalLiveTable.ItemsSource = filteredJournal;





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


        // ================= СОХРАНЕНИЕ =================

        private void SaveState()
        {
            File.WriteAllText(
                SaveFileName,
                JsonSerializer.Serialize(new AppState
                {
                    CurrentObject = currentObject,
                    Journal = journal
                }));
        }
        private AppState CloneState()
        {
            return new AppState
            {
                CurrentObject = JsonSerializer.Deserialize<ProjectObject>(
            JsonSerializer.Serialize(currentObject)),
                Journal = JsonSerializer.Deserialize<List<JournalRecord>>(
            JsonSerializer.Serialize(journal))
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
            summaryFilterInitialized = false;
            journal = state.Journal ?? new();

            ArrivalPanel.SetObject(currentObject, journal);

            RefreshTreePreserveState();

            RefreshSummaryTable();
            EnsureOtJournalStorage();
            BindOtJournal();
            SaveState();
        }

        private void UpdateUndoRedoButtons()
        {
            UndoButton.IsEnabled = undoStack.Count > 0;
            RedoButton.IsEnabled = redoStack.Count > 0;
        }

        private void LoadState()
        {
            if (!File.Exists(SaveFileName))
                return;

            var state = JsonSerializer.Deserialize<AppState>(
                File.ReadAllText(SaveFileName));

            currentObject = state?.CurrentObject;
            summaryFilterInitialized = false;
            journal = state?.Journal ?? new();
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
            EnsureOtJournalStorage();

        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveState();
            MessageBox.Show("Данные сохранены");
        }

        private void LockToggle_Checked(object sender, RoutedEventArgs e)
        {
            LockButton_Checked(sender, e);
        }

        private void LockToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            LockButton_Unchecked(sender, e);
        }



        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            SaveState();
            Close();
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

            summaryHasOverage = false;
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
                Text = "Формат ячейки: план / пришло",
                Foreground = new SolidColorBrush(Color.FromRgb(107, 114, 128)),
                Margin = new Thickness(0, 0, 0, 8)
            };

            SummaryPanel.Items.Add(note);

            summaryOverageNote = new TextBlock
            {
                Text = "⚠️ Есть превышение прихода относительно плана.",
                Foreground = new SolidColorBrush(Color.FromRgb(180, 83, 9)),
                Margin = new Thickness(0, 0, 0, 8),
                Visibility = Visibility.Collapsed
            };

            SummaryPanel.Items.Add(summaryOverageNote);
        }

        void RenderMaterialGroup(string group)
        {
            summaryBlocks = BuildSummaryBlocks();
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
                foreach (var floor in block.Floors)
                {
                    summaryColumns.Add(new SummaryColumnInfo
                    {
                        ColumnIndex = colIndex,
                        Block = block.Block,
                        Floor = floor
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
                int blockColumns = block.Floors.Count + 1;

                AddCell(summaryGrid, 0, blockStart, $"Блок {block.Block}", bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, colspan: blockColumns, noWrap: true);

                int floorCol = blockStart;
                foreach (var floor in block.Floors)
                {
                    AddCell(summaryGrid, 1, floorCol, GetFloorLabel(floor), bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);
                    floorCol++;
                }

                AddCell(summaryGrid, 1, floorCol, "Итого", bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);
                blockStart += blockColumns;
            }

            AddCell(summaryGrid, 0, summaryTotalColumn, "Всего на здание", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);
            AddCell(summaryGrid, 0, summaryNotArrivedColumn, "Не доехало", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);
            AddCell(summaryGrid, 0, summaryArrivedColumn, "Пришло", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, noWrap: true);

            summaryRowIndex = 2;
        }

        void RenderMaterialRow(string group, string mat, string unit, double totalArrival, string position)
        {
            if (summaryGrid == null)
                return;

            var blockTotals = new Dictionary<int, double>(); // ✅ ДОБАВИТЬ

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
                foreach (var floor in block.Floors)
                {
                    blockTotal += GetDemandValue(demand, block.Block, floor);
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
            var rowHighlight = rowComplete
                ? new SolidColorBrush(Color.FromRgb(209, 250, 229))
                : null;
            var filledHighlight = new SolidColorBrush(Color.FromRgb(220, 252, 231));
            var blockHighlight = new SolidColorBrush(Color.FromRgb(219, 234, 254));
            var warningHighlight = new SolidColorBrush(Color.FromRgb(254, 243, 199));

            if (rowOverage || blockOverage.Values.Any())
            {
                summaryHasOverage = true;
                if (summaryOverageNote != null)
                    summaryOverageNote.Visibility = Visibility.Visible;
            }

            AddCell(summaryGrid, summaryRowIndex, 0, position, align: TextAlignment.Center, bg: rowHighlight, noWrap: true, minWidth: 60);
            AddCell(summaryGrid, summaryRowIndex, 1, mat, bg: rowHighlight, noWrap: true, minWidth: 220);

            foreach (var col in summaryColumns)
            {
                if (col.IsBlockTotal)
                {
                    double blockTotal = blockTotals.TryGetValue(col.Block, out var val) ? val : 0;
                   
                    bool blockComplete = blockFilled.TryGetValue(col.Block, out var complete) && complete;
                    bool blockIsOverage = blockOverage.TryGetValue(col.Block, out var over) && over;
                    Brush cellBg = rowHighlight ?? (blockIsOverage ? warningHighlight : (blockComplete ? blockHighlight : null));
                    AddCell(summaryGrid, summaryRowIndex, col.ColumnIndex, FormatNumber(blockTotal), align: TextAlignment.Right, bg: cellBg, noWrap: true, minWidth: 44);
                }
                else if (col.Floor.HasValue)
                {
                    double plan = GetDemandValue(demand, col.Block, col.Floor.Value);
                    double arrived = allocations.TryGetValue(col.Block, out var blockDict)
                        && blockDict.TryGetValue(col.Floor.Value, out var arr)
                        ? arr
                        : 0;

                    bool floorOverage = plan > 0 ? arrived > plan : arrived > 0;
                    bool floorFilled = plan > 0 && Math.Abs(arrived - plan) < 0.0001;
                    if (floorOverage)
                    {
                        summaryHasOverage = true;
                        if (summaryOverageNote != null)
                            summaryOverageNote.Visibility = Visibility.Visible;
                    }

                    Brush cellBg = rowHighlight ?? (floorOverage ? warningHighlight : (floorFilled ? filledHighlight : null));
                    AddDiagonalDemandCell(summaryGrid, summaryRowIndex, col.ColumnIndex, plan, arrived, demandKey, col.Block, col.Floor.Value, unit, cellBg, 44);
                }
            }

            double notArrived = Math.Max(0, totalPlanned - totalArrival);
            Brush arrivedBg = rowHighlight ?? (rowOverage ? warningHighlight : null);
            AddCell(summaryGrid, summaryRowIndex, summaryTotalColumn, FormatNumber(totalPlanned), align: TextAlignment.Right, bg: rowHighlight, noWrap: true, minWidth: 70);
            AddCell(summaryGrid, summaryRowIndex, summaryNotArrivedColumn, FormatNumber(notArrived), align: TextAlignment.Right, bg: rowHighlight, noWrap: true, minWidth: 70);
            AddCell(summaryGrid, summaryRowIndex, summaryArrivedColumn, FormatNumber(totalArrival), align: TextAlignment.Right, bg: arrivedBg, noWrap: true, minWidth: 70);

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

            if (SummaryTypeFilterPanel != null)
                SummaryTypeFilterPanel.Children.Clear();

            if (groups.Count == 0)
            {
                UpdateSummaryFilterSubtitle(groups, new List<string>());
                summaryFilterUpdating = false;
                return;
            }

            var selectedGroups = currentObject.SummaryVisibleGroups
                .Where(groups.Contains)
                .ToList();

            if (selectedGroups.Count == 0)
                selectedGroups = new List<string> { groups[0] };
            else if (selectedGroups.Count > 1)
                selectedGroups = selectedGroups.Take(1).ToList();

            summaryFilterInitialized = true;


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
                       UpdateSummaryFilterSubtitle(groups, selectedGroups);
                       summaryFilterUpdating = false;
        }


        private void SummaryFilterOptionChanged(object sender, RoutedEventArgs e)
        {
            if (summaryFilterUpdating || currentObject == null)
                return;

            if (sender is not RadioButton radio)
                return;

            var selectedGroup = radio.Tag?.ToString();
            var selected = string.IsNullOrWhiteSpace(selectedGroup)
                ? new List<string>()
                : new List<string> { selectedGroup };

            currentObject.SummaryVisibleGroups = selected;
            UpdateSummaryFilterSubtitle(summaryFilterGroups, selected);

            
            RefreshSummaryTable();
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



        private List<string> GetMaterialsForGroup(string group)
        {
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

        private void UpdateTreePanelState(bool forceVisible)
        {
            if (TreeColumn == null || TreePanelBorder == null)
                return;

            bool show = isTreePinned || forceVisible;
            TreePanelBorder.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            TreeColumn.Width = show ? new GridLength(260) : new GridLength(0);
        }
        private List<SummaryBlockInfo> BuildSummaryBlocks()
        {
            var blocks = new List<SummaryBlockInfo>();

            if (currentObject == null || currentObject.BlocksCount <= 0)
                return blocks;

            for (int i = 1; i <= currentObject.BlocksCount; i++)
            {
                int floors = currentObject.SameFloorsInBlocks
                    ? currentObject.FloorsPerBlock
                    : (currentObject.FloorsByBlock.TryGetValue(i, out var f) ? f : 0);

                var floorList = new List<int>();

                if (currentObject.HasBasement)
                    floorList.Add(0);

                for (int floor = 1; floor <= floors; floor++)
                    floorList.Add(floor);

                blocks.Add(new SummaryBlockInfo
                {
                    Block = i,
                    Floors = floorList
                });
            }

            return blocks;
        }

        private Dictionary<int, Dictionary<int, double>> AllocateArrival(MaterialDemand demand, double totalArrival)
        {
            var allocations = new Dictionary<int, Dictionary<int, double>>();

            if (summaryBlocks == null || summaryBlocks.Count == 0)
                return allocations;

            var levels = new List<int>();

            if (currentObject?.HasBasement == true)
                levels.Add(0);

            int maxFloor = summaryBlocks
                .SelectMany(b => b.Floors)
                .Where(f => f > 0)
                .DefaultIfEmpty(0)
                .Max();

            for (int i = 1; i <= maxFloor; i++)
                levels.Add(i);

            double remaining = totalArrival;

            foreach (var level in levels)
            {
                foreach (var block in summaryBlocks)
                {
                    if (!block.Floors.Contains(level))
                        continue;

                    double plan = GetDemandValue(demand, block.Block, level);
                    double filled = Math.Min(plan, remaining);

                    if (!allocations.ContainsKey(block.Block))
                        allocations[block.Block] = new Dictionary<int, double>();

                    allocations[block.Block][level] = filled;
                    remaining -= filled;

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
                    Floors = new Dictionary<int, Dictionary<int, double>>()
                };

                currentObject.Demand[demandKey] = demand;
            }

            if (string.IsNullOrWhiteSpace(demand.Unit))
                demand.Unit = unit;

            return demand;
        }

        private double GetDemandValue(MaterialDemand demand, int block, int floor)
        {
            if (demand.Floors.TryGetValue(block, out var floors)
                && floors.TryGetValue(floor, out var value))
                return value;

            return 0;
        }

        private string GetFloorLabel(int floor)
        {
            return floor == 0 ? "Подвал" : floor.ToString();
        }

        private class SummaryBlockInfo
        {
            public int Block { get; set; }
            public List<int> Floors { get; set; } = new();
        }

        private class SummaryColumnInfo
        {
            public int ColumnIndex { get; set; }
            public int Block { get; set; }
            public int? Floor { get; set; }
            public bool IsBlockTotal { get; set; }
        }

        private class DemandCellTag
        {
            public string DemandKey { get; set; }
            public int Block { get; set; }
            public int Floor { get; set; }
            public string Unit { get; set; }
        }

        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            if (undoStack.Count == 0)
                return;

            redoStack.Push(CloneState());
            var prev = undoStack.Pop();
            RestoreState(prev);
            UpdateUndoRedoButtons();
        }

        private void Redo_Click(object sender, RoutedEventArgs e)
        {
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
                BorderBrush = new SolidColorBrush(Color.FromRgb(220, 223, 227)),
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
        void AddDiagonalDemandCell(Grid g, int r, int c, double plan, double arrived, string demandKey, int block, int floor, string unit, Brush bg, double minWidth)
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

            var planBox = new TextBox
            {
                Text = FormatNumber(plan),
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Margin = new Thickness(2, 1, 2, 1),
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                MinWidth = 22,
                FontSize = 11,
                IsReadOnly = isLocked,
                IsEnabled = !isLocked,
                Tag = new DemandCellTag
                {
                    DemandKey = demandKey,
                    Block = block,
                    Floor = floor,
                    Unit = unit
                }
            };

            planBox.LostFocus += DemandCell_LostFocus;

            var arrivedText = new TextBlock
            {
                Text = FormatNumber(arrived),
                Margin = new Thickness(2, 1, 2, 1),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Foreground = new SolidColorBrush(Color.FromRgb(55, 65, 81)),
                FontSize = 11
            };

            container.Children.Add(planBox);
            container.Children.Add(arrivedText);

            var border = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(220, 223, 227)),
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

        private void DemandCell_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is not TextBox tb || tb.Tag is not DemandCellTag tag)
                return;

            if (isLocked)
                return;


            var text = tb.Text?.Trim() ?? string.Empty;
            double value = ParseNumber(text);

            var demand = GetOrCreateDemand(tag.DemandKey, tag.Unit);

            if (!demand.Floors.ContainsKey(tag.Block))
                demand.Floors[tag.Block] = new Dictionary<int, double>();

            demand.Floors[tag.Block][tag.Floor] = value;

            RefreshSummaryTable();
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

            int colDate = 95;
            int colTtn = Math.Max(140, maxTtn * 7);
            int colName = Math.Max(260, maxName * 7);
            int colStb = 90;
            int colUnit = 70;
            int colQty = 90;
            int colSupplier = Math.Max(220, maxSupplier * 7);
            int colPassport = Math.Max(260, maxPassport * 7);
           


            int maxTotalWidth = 1400;
            int total = colDate + colTtn + colName + colStb + colUnit + colQty + colSupplier + colPassport;

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
                    BorderBrush = new SolidColorBrush(Color.FromRgb(220, 223, 227)), // тот же тон что в таблице
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
                        AddCell(grid, r, 1, name, wrap: true, bg: bg);
                        AddCell(grid, r, 2, stb, bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 3, unit, bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 4, qty, bg: bg, align: TextAlignment.Right);
                        AddCell(grid, r, 5, supplier, wrap: true, bg: bg);
                        AddCell(grid, r, 6, passport, wrap: true, bg: bg);
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

                        AddCell(dayGrid, rowIndex, 1, name, wrap: true, bg: bg);
                        
                        AddCell(dayGrid, rowIndex, 4, qty, bg: bg, align: TextAlignment.Right);
        


                        rowIndex++;
                    }
                    AddCell(dayGrid, start, 2, stbMerged, rowspan: rows, bg: bg, align: TextAlignment.Center);
                    AddCell(dayGrid, start, 3, unitMerged, rowspan: rows, bg: bg, align: TextAlignment.Center);
                    AddCell(dayGrid, start, 5, supplierMerged, rowspan: rows, wrap: true, bg: bg);
                    AddCell(dayGrid, start, 6, passportMerged, rowspan: rows, wrap: true, bg: bg);


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
                    selectedArrivalTypes.Add(g);
                    RefreshArrivalNames();
                    ApplyAllFilters();
                };
                chip.Unchecked += (_, _) =>
                {
                    selectedArrivalTypes.Remove(g);
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