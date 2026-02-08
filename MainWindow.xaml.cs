using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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


        public MainWindow()
        {
            InitializeComponent();
            ArrivalLiveTable.IsReadOnly = true;
            ArrivalLiveTable.CanUserAddRows = false;
            ArrivalLiveTable.CanUserDeleteRows = false;


            // ===== БЛОКИРОВКА ВКЛЮЧЕНА ПО УМОЛЧАНИЮ =====
            isLocked = true;

            LoadState();
            ApplyAllFilters();

            ArrivalPanel.ArrivalAdded += OnArrivalAdded;

            PushUndo();
            UpdateUndoRedoButtons();

            if (currentObject != null)
                ArrivalPanel.SetObject(currentObject, journal);
            RefreshArrivalTypes();
            RefreshArrivalNames();

            RefreshTreePreserveState();

        }
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // гарантированно после создания всех контролов

        }
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source is TabControl tab &&
                tab.SelectedItem is TabItem item &&
                item.Header.ToString() == "Приход")
            {
                ArrivalPopup.Visibility = arrivalPanelVisible
                    ? Visibility.Visible
                    : Visibility.Collapsed;

                ShowArrivalButton.Visibility = arrivalPanelVisible
                    ? Visibility.Collapsed
                    : Visibility.Visible;
            }
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
                .GroupBy(j => j.MaterialGroup);

            foreach (var g in mainGroups)
            {
                var groupNode = new TreeViewItem
                {
                    Header = g.Key,
                    Tag = "Group",
                    IsExpanded = true
                };

                foreach (var m in g.Select(x => x.MaterialName).Distinct())
                {
                    groupNode.Items.Add(new TreeViewItem
                    {
                        Header = m,
                        Tag = "Material"
                    });
                }

                mainNode.Items.Add(groupNode);
            }

            // ===== ДОПЫ =====
            var extraGroups = journal
                .Where(j => j.Category == "Допы")
                .GroupBy(j => j.SubCategory);

            foreach (var g in extraGroups)
            {
                var subNode = new TreeViewItem
                {
                    Header = g.Key,
                    Tag = "SubCategory",
                    IsExpanded = true
                };

                foreach (var m in g.Select(x => x.MaterialName).Distinct())
                {
                    subNode.Items.Add(new TreeViewItem
                    {
                        Header = m,
                        Tag = "Material"
                    });
                }

                extraNode.Items.Add(subNode);
            }

            newRoot.Items.Add(mainNode);
            newRoot.Items.Add(extraNode);

            ObjectsTree.Items.Add(newRoot);
        }


        private void ObjectsTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            ApplyAllFilters();
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

            if (node.Tag as string == "Object")
                return;

            var oldName = node.Header.ToString();

            var input = Microsoft.VisualBasic.Interaction.InputBox(
                "Новое название:",
                "Переименование",
                oldName);

            if (string.IsNullOrWhiteSpace(input) || input == oldName)
                return;
            PushUndo(); // ⬅️ ВАЖНО: сохраняем состояние ДО переименования

            if (node.Tag as string == "Group")
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

            if (node.Tag as string == "Material")
            {
                foreach (var kv in currentObject.MaterialNamesByGroup)
                {
                    var idx = kv.Value.IndexOf(oldName);
                    if (idx >= 0)
                        kv.Value[idx] = input;
                }

                foreach (var j in journal.Where(x => x.MaterialName == oldName))
                    j.MaterialName = input;
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

            if (node.Tag as string == "Object")
                return;

            var name = node.Header.ToString();

            if (MessageBox.Show($"Удалить \"{name}\"?",
                "Подтверждение",
                MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;

            if (node.Tag as string == "Group")
            {
                currentObject.MaterialGroups.RemoveAll(g => g.Name == name);
                currentObject.MaterialNamesByGroup.Remove(name);
                journal.RemoveAll(j => j.MaterialGroup == name);
            }

            if (node.Tag as string == "Material")
            {
                foreach (var kv in currentObject.MaterialNamesByGroup)
                    kv.Value.Remove(name);

                journal.RemoveAll(j => j.MaterialName == name);
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




            if (ObjectsTree.SelectedItem is TreeViewItem node &&
                node.Tag is string tag)
            {
                var value = node.Header.ToString();

                if (tag == "Group")
                    data = data.Where(j => j.MaterialGroup == value);
                else if (tag == "Material")
                    data = data.Where(j => j.MaterialName == value);
            }


            // === ПРИХОД: ДАТЫ ===
            if (ArrivalDateFrom?.SelectedDate != null)
                data = data.Where(j => j.Date >= ArrivalDateFrom.SelectedDate);

            if (ArrivalDateTo?.SelectedDate != null)
                data = data.Where(j => j.Date <= ArrivalDateTo.SelectedDate);




       



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
                // после изменений — обновляем всё
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
                Text = "Формат ячейки: план / пришло",
                Foreground = new SolidColorBrush(Color.FromRgb(107, 114, 128)),
                Margin = new Thickness(0, 0, 0, 8)
            };

            SummaryPanel.Items.Add(note);
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

            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });
            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(240) });

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
                    summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });
                    colIndex++;
                }

                summaryColumns.Add(new SummaryColumnInfo
                {
                    ColumnIndex = colIndex,
                    Block = block.Block,
                    IsBlockTotal = true
                });
                summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });
                colIndex++;
            }

            summaryTotalColumn = colIndex++;
            summaryNotArrivedColumn = colIndex++;
            summaryArrivedColumn = colIndex++;

            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
            summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });

            var headerBg = new SolidColorBrush(Color.FromRgb(243, 244, 246));

            AddCell(summaryGrid, 0, 0, "Позиция", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold);
            AddCell(summaryGrid, 0, 1, "Наименование", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold);

            int blockStart = 2;
            foreach (var block in summaryBlocks)
            {
                int blockColumns = block.Floors.Count + 1;

                AddCell(summaryGrid, 0, blockStart, $"Блок {block.Block}", bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold, colspan: blockColumns);

                int floorCol = blockStart;
                foreach (var floor in block.Floors)
                {
                    AddCell(summaryGrid, 1, floorCol, GetFloorLabel(floor), bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold);
                    floorCol++;
                }

                AddCell(summaryGrid, 1, floorCol, "Итого", bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold);
                blockStart += blockColumns;
            }

            AddCell(summaryGrid, 0, summaryTotalColumn, "Всего на здание", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold);
            AddCell(summaryGrid, 0, summaryNotArrivedColumn, "Не доехало", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold);
            AddCell(summaryGrid, 0, summaryArrivedColumn, "Пришло", rowspan: 2, bg: headerBg, align: TextAlignment.Center, fontWeight: FontWeights.SemiBold);

            summaryRowIndex = 2;
        }

        void RenderMaterialRow(string group, string mat, string unit, double totalArrival, string position)
        {
            if (summaryGrid == null)
                return;

            summaryGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            string demandKey = BuildDemandKey(group, mat);
            var demand = GetOrCreateDemand(demandKey, unit);
            var allocations = AllocateArrival(demand, totalArrival);

            double totalPlanned = 0;
            var blockTotals = new Dictionary<int, double>();

            foreach (var block in summaryBlocks)
            {
                double blockTotal = 0;
                foreach (var floor in block.Floors)
                {
                    blockTotal += GetDemandValue(demand, block.Block, floor);
                }
                blockTotals[block.Block] = blockTotal;
                totalPlanned += blockTotal;
            }

            AddCell(summaryGrid, summaryRowIndex, 0, position, align: TextAlignment.Center);
            AddCell(summaryGrid, summaryRowIndex, 1, mat, wrap: true);

            foreach (var col in summaryColumns)
            {
                if (col.IsBlockTotal)
                {
                    double blockTotal = blockTotals.TryGetValue(col.Block, out var val) ? val : 0;
                    AddCell(summaryGrid, summaryRowIndex, col.ColumnIndex, FormatNumber(blockTotal), align: TextAlignment.Right);
                }
                else if (col.Floor.HasValue)
                {
                    double plan = GetDemandValue(demand, col.Block, col.Floor.Value);
                    double arrived = allocations.TryGetValue(col.Block, out var blockDict)
                        && blockDict.TryGetValue(col.Floor.Value, out var arr)
                        ? arr
                        : 0;

                    AddDiagonalDemandCell(summaryGrid, summaryRowIndex, col.ColumnIndex, plan, arrived, demandKey, col.Block, col.Floor.Value, unit);
                }
            }

            double notArrived = Math.Max(0, totalPlanned - totalArrival);

            AddCell(summaryGrid, summaryRowIndex, summaryTotalColumn, FormatNumber(totalPlanned), align: TextAlignment.Right);
            AddCell(summaryGrid, summaryRowIndex, summaryNotArrivedColumn, FormatNumber(notArrived), align: TextAlignment.Right);
            AddCell(summaryGrid, summaryRowIndex, summaryArrivedColumn, FormatNumber(totalArrival), align: TextAlignment.Right);

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
            SummaryTypeSelector.ItemsSource = null;
            SummaryTypeSelector.ItemsSource = groups;

            if (groups.Count == 0)
            {
                SummaryTypeSelector.SelectedItem = null;
                summaryFilterUpdating = false;
                return;
            }

            string selectedGroup = currentObject.SummaryVisibleGroups.FirstOrDefault();

            if (!summaryFilterInitialized)
            {
                if (string.IsNullOrWhiteSpace(selectedGroup))
                {
                    selectedGroup = groups[0];
                    currentObject.SummaryVisibleGroups = new List<string> { selectedGroup };
                }

                summaryFilterInitialized = true;
            }
            if (!groups.Contains(selectedGroup))
            {
                selectedGroup = groups[0];
                currentObject.SummaryVisibleGroups = new List<string> { selectedGroup };

            }

            currentObject.SummaryVisibleGroups = new List<string> { selectedGroup };
            SummaryTypeSelector.SelectedItem = selectedGroup;
            summaryFilterUpdating = false;
        }

        private void SummaryTypeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (summaryFilterUpdating || currentObject == null)
                return;

            if (SummaryTypeSelector.SelectedItem is not string group)
                return;
            currentObject.SummaryVisibleGroups = new List<string> { group };

            RefreshSummaryTable();
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
        void AddCell(Grid g, int r, int c, string text, int rowspan = 1, bool wrap = false, Brush bg = null, TextAlignment align = TextAlignment.Left, FontWeight? fontWeight = null, int colspan = 1)
        {
            var tb = new TextBlock
            {
                Text = text,
                Margin = new Thickness(6, 4, 6, 4),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
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

            border.Child = tb;

            Grid.SetRow(border, r);
            Grid.SetColumn(border, c);

            if (rowspan > 1)
                Grid.SetRowSpan(border, rowspan);

            if (colspan > 1)
                Grid.SetColumnSpan(border, colspan);


            g.Children.Add(border);
        }
        void AddDiagonalDemandCell(Grid g, int r, int c, double plan, double arrived, string demandKey, int block, int floor, string unit)
        {
            var container = new Grid
            {
                SnapsToDevicePixels = true,
                UseLayoutRounding = true
            };

            var line = new Line
            {
                X1 = 0,
                Y2 = 0,
                Stroke = new SolidColorBrush(Color.FromRgb(209, 213, 219)),
                StrokeThickness = 1,
                SnapsToDevicePixels = true,
                IsHitTestVisible = false
            };

            line.SetBinding(Line.X2Property, new Binding("ActualWidth")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(Grid), 1)
            });

            line.SetBinding(Line.Y1Property, new Binding("ActualHeight")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(Grid), 1)
            });

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
                Background = Brushes.White,
                MinHeight = 24,
                Child = container
            };

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