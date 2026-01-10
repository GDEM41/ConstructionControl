using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ConstructionControl
{
    public partial class MainWindow : Window
    {
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
        private bool mergeEnabled = false;

        private bool isLocked;

        public MainWindow()
        {
            InitializeComponent();

            // ===== БЛОКИРОВКА ВКЛЮЧЕНА ПО УМОЛЧАНИЮ =====
            isLocked = true;

            LoadState();
            ArrivalPanel.ArrivalAdded += OnArrivalAdded;

            PushUndo();
            UpdateUndoRedoButtons();

            if (currentObject != null)
                ArrivalPanel.SetObject(currentObject, journal);

            RefreshTreePreserveState();
            RefreshFilters();
            ApplyAllFilters();
        }
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // гарантированно после создания всех контролов
            
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

                ArrivalPanel.SetObject(currentObject, journal);



                SaveState();
                RefreshTreePreserveState();
                RefreshFilters();
                ApplyAllFilters();
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

        private void ToggleFilters_Click(object sender, RoutedEventArgs e)
        {
            FiltersPanel.Visibility =
                FiltersPanel.Visibility == Visibility.Visible
                    ? Visibility.Collapsed
                    : Visibility.Visible;
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (!filteredJournal.Any())
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "CSV (*.csv)|*.csv",
                FileName = "ЖВК.csv"
            };

            if (dlg.ShowDialog() != true)
                return;

            using var sw = new StreamWriter(dlg.FileName, false, System.Text.Encoding.UTF8);
            sw.WriteLine("Дата;Тип;Наименование;Кол-во;Ед.;ТТН;Паспорт;СТБ;Поставщик");

            foreach (var j in filteredJournal)
            {
                sw.WriteLine(
                    $"{j.Date:dd.MM.yyyy};{j.MaterialGroup};{j.MaterialName};{j.Quantity};{j.Unit};{j.Ttn};{j.Passport};{j.Stb};{j.Supplier}");
            }
        }

        private void LockButton_Checked(object sender, RoutedEventArgs e)
        {
            isLocked = true;


        }

        private void LockButton_Unchecked(object sender, RoutedEventArgs e)
        {
            isLocked = false;


        }


        // ================= ПРИХОД =================

        private void OnArrivalAdded(Arrival arrival)
        {
            PushUndo(); // ⬅️ ВОТ ЭТОГО НЕ ХВАТАЛО

            PushUndo();

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
                // ===== ТОЛЬКО ДЛЯ ОСНОВНЫХ =====
                if (arrival.Category == "Основные")
                {
                    if (!currentObject.MaterialNamesByGroup[arrival.MaterialGroup]
                            .Contains(i.MaterialName))
                    {
                        currentObject.MaterialNamesByGroup[arrival.MaterialGroup]
                            .Add(i.MaterialName);
                    }
                }

                // ===== ОБЩЕЕ: добавляем запись =====
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
            RefreshFilters();
            ApplyAllFilters();

            // важно: обновляем панель прихода
            ArrivalPanel.SetObject(currentObject, journal);


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
            RefreshFilters();
            ApplyAllFilters();
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
            RefreshFilters();
            ApplyAllFilters();
        }

        // ================= ФИЛЬТРЫ =================

        private void RefreshFilters()
        {
            if (currentObject == null)
                return;

            FilterGroupsList.ItemsSource =
                currentObject.MaterialGroups
                    .Select(g => g.Name)
                    .OrderBy(x => x)
                    .ToList();
        }

        private void Filters_Changed(object sender, EventArgs e)
        {
            ApplyAllFilters();
        }

        private void SelectAllGroups_Click(object sender, RoutedEventArgs e)
        {
            FilterGroupsList.SelectAll();
            ApplyAllFilters();
        }

        private void ClearGroups_Click(object sender, RoutedEventArgs e)
        {
            FilterGroupsList.UnselectAll();
            ApplyAllFilters();
        }

        private void ApplyAllFilters()
        {
            IEnumerable<JournalRecord> data = journal;


            // ===== ДОПОЛНИТЕЛЬНЫЕ ФИЛЬТРЫ =====
            // ===== ДОПОЛНИТЕЛЬНЫЕ ФИЛЬТРЫ (ДОПЫ ПО УМОЛЧАНИЮ СКРЫТЫ) =====
            bool showLowCost = LowCostCheckBox.IsChecked == true;
            bool showInternal = InternalCheckBox.IsChecked == true;

            data = data.Where(j =>
                // основные всегда видны
                j.Category == "Основные"

                // допы — только если включены галочки
                || (
                    j.Category == "Допы" &&
                    (
                        // если включены ОБЕ — показываем ВСЕ допы
                        (showLowCost && showInternal)

                        // если включена только одна
                        || (showLowCost && j.SubCategory == "Малоценка")
                        || (showInternal && j.SubCategory == "Внутренние")
                    )
                )
            );



            if (ObjectsTree.SelectedItem is TreeViewItem node &&
                node.Tag is string tag)
            {
                var value = node.Header.ToString();

                if (tag == "Group")
                    data = data.Where(j => j.MaterialGroup == value);
                else if (tag == "Material")
                    data = data.Where(j => j.MaterialName == value);
            }

            if (FilterGroupsList.SelectedItems.Count > 0)
            {
                var groups = FilterGroupsList.SelectedItems.Cast<string>().ToList();
                data = data.Where(j => groups.Contains(j.MaterialGroup));
            }

            if (DateFromPicker.SelectedDate != null)
                data = data.Where(j => j.Date >= DateFromPicker.SelectedDate);

            if (DateToPicker.SelectedDate != null)
                data = data.Where(j => j.Date <= DateToPicker.SelectedDate);
            // ===== ГЛОБАЛЬНЫЙ ПОИСК =====
            if (!string.IsNullOrWhiteSpace(GlobalSearchBox.Text))
            {
                var text = GlobalSearchBox.Text.Trim();

                data = data.Where(j =>
                    (j.MaterialName != null &&
                     j.MaterialName.Contains(text, StringComparison.OrdinalIgnoreCase))

                    || (j.MaterialGroup != null &&
                        j.MaterialGroup.Contains(text, StringComparison.OrdinalIgnoreCase))

                    || (j.Ttn != null &&
                        j.Ttn.Contains(text, StringComparison.OrdinalIgnoreCase))

                    || (j.Passport != null &&
                        j.Passport.Contains(text, StringComparison.OrdinalIgnoreCase))
                );
            }



            filteredJournal = data
                .OrderByDescending(j => j.Date) // ⬅️ СОРТИРОВКА ПО ДАТЕ (НОВЫЕ СВЕРХУ)
                .ToList();

            RenderJvk();

            RefreshSummaryTable();


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
            journal = state.Journal ?? new();

            ArrivalPanel.SetObject(currentObject, journal);

            RefreshTreePreserveState();
            RefreshFilters();
            ApplyAllFilters();
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
            journal = state?.Journal ?? new();
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

        private void ExtraFiltersToggle_Checked(object sender, RoutedEventArgs e)
        {
            ExtraFiltersPanel.Visibility = Visibility.Visible;
        }

        private void ExtraFiltersToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            ExtraFiltersPanel.Visibility = Visibility.Collapsed;

            LowCostCheckBox.IsChecked = false;
            InternalCheckBox.IsChecked = false;

            ApplyAllFilters();
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
            var importWindow = new ExcelImportWindow(dlg.FileName, sheetNames)
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
            RefreshFilters();
            ApplyAllFilters();
            RefreshSummaryTable();
            ArrivalPanel.SetObject(currentObject, journal);






        }


        private void CloseFilters_Click(object sender, RoutedEventArgs e)
        {
            FiltersPanel.Visibility = Visibility.Collapsed;
        }


        private void RefreshSummaryTable()
        {
            if (currentObject == null)
                return;

            var result = new List<SummaryRow>();

            var materials = journal
                .GroupBy(j => new { j.MaterialName, j.Unit });

            foreach (var mat in materials)
            {
                var row = new SummaryRow
                {
                    MaterialName = mat.Key.MaterialName,
                    Unit = mat.Key.Unit
                };

                // ❗ ГАРАНТИРУЕМ ХОТЯ БЫ 1 БЛОК
                int blocks = Math.Max(1, currentObject.BlocksCount);

                for (int block = 1; block <= blocks; block++)
                    row.ByBlocks[block] = 0.0;

                foreach (var rec in mat)
                {
                    row.ByBlocks[1] += rec.Quantity;
                    row.Total += rec.Quantity;
                }

                result.Add(row);
            }

            SummaryGrid.ItemsSource = result;
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
        void AddCell(Grid g, int r, int c, string text, int rowspan = 1, bool wrap = false, Brush bg = null, TextAlignment align = TextAlignment.Left)
        {
            var tb = new TextBlock
            {
                Text = text,
                Margin = new Thickness(6, 2, 6, 2),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = wrap ? TextWrapping.Wrap : TextWrapping.NoWrap,
                TextAlignment = align
            };

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

            g.Children.Add(border);
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
                return;

            // ===== авто размер колонок =====
            int maxName = filteredJournal.Max(j => j.MaterialName?.Length ?? 0);
            int maxPassport = filteredJournal.Max(j => j.Passport?.Length ?? 0);
            int maxSupplier = filteredJournal.Max(j => j.Supplier?.Length ?? 0);
            int maxTtn = filteredJournal.Max(j => j.Ttn?.Length ?? 0);

            int colDate = 95;
            int colTtn = Math.Max(120, maxTtn * 7);
            int colName = Math.Max(250, maxName * 7);
            int colStb = 70;
            int colUnit = 45;
            int colQty = 70;
            int colSupplier = Math.Max(180, maxSupplier * 7);
            int colPassport = Math.Max(260, maxPassport * 7);
            RenderHeader(colTtn, colName, colStb, colUnit, colQty, colSupplier, colPassport);


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

            // ===== шапка =====
            RenderHeader(colTtn, colName, colStb, colUnit, colQty, colSupplier, colPassport);

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

                    AddCell(grid, 0, 0, "", rows, bg: bg, align: TextAlignment.Center);


                    AddCell(grid, 0, 0, ttn.Key.Ttn, rows, bg: bg, align: TextAlignment.Center);


                    for (int r = 0; r < rows; r++)
                    {
                        var x = items[r];
                        AddCell(grid, r, 1, x.MaterialName, wrap: true, bg: bg, align: TextAlignment.Left);
                        AddCell(grid, r, 2, x.Stb ?? "—", bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 3, x.Unit ?? "—", bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 4, x.Quantity.ToString(), bg: bg, align: TextAlignment.Right);
                        AddCell(grid, r, 5, x.Supplier ?? "—", wrap: true, bg: bg, align: TextAlignment.Left);
                        AddCell(grid, r, 6, x.Passport ?? "—", wrap: true, bg: bg, align: TextAlignment.Left);

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
        void RenderHeader(int colTtn, int colName, int colStb, int colUnit, int colQty, int colSupplier, int colPassport)
        {
            HeaderAutoPanel.Children.Clear();

            var grid = new Grid
            {
                Margin = new Thickness(0, 0, 0, 4)
            };

            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colTtn) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colName) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colStb) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colUnit) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colQty) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colSupplier) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colPassport) });

            void Add(string text, int col, TextAlignment align = TextAlignment.Left)
            {
                var tb = new TextBlock
                {
                    Text = text,
                    FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(6, 2, 6, 2),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextAlignment = align
                };

                Grid.SetColumn(tb, col);
                grid.Children.Add(tb);
            }

            Add("ТТН", 0, TextAlignment.Center);
            Add("Наименование", 1);
            Add("СТБ", 2, TextAlignment.Center);
            Add("Ед.", 3, TextAlignment.Center);
            Add("Кол-во", 4, TextAlignment.Right);
            Add("Поставщик", 5);
            Add("Паспорт", 6);

            HeaderAutoPanel.Children.Add(grid);
        }
        void RenderMerged(
    IEnumerable<dynamic> merged,
    int colTtn, int colName, int colStb, int colUnit, int colQty, int colSupplier, int colPassport)
        {
            foreach (var day in merged)
            {
                var daySeparator = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(220, 223, 227)),
                    BorderThickness = new Thickness(0, 1, 0, 0),
                    Margin = new Thickness(0, 12, 0, 8)
                };
                JvkPanel.Children.Add(daySeparator);

                var dateHeader = new TextBlock
                {
                    Text = day.Date.ToString("dd.MM.yyyy"),
                    FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(0, 0, 0, 6),
                    FontSize = 15
                };
                JvkPanel.Children.Add(dateHeader);

                foreach (var grp in day.Groups)
                {
                    var items = grp.Items;
                    int rows = items.Count;

                    var grid = new Grid { Margin = new Thickness(0, 0, 0, 4) };
                    var bg = new SolidColorBrush(GetSoftColor(grp.Ttn ?? ""));

                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colTtn) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colName) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colStb) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colUnit) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colQty) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colSupplier) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(colPassport) });

                    for (int i = 0; i < rows; i++)
                        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                    AddCell(grid, 0, 0, grp.Ttn, rows, bg: bg, align: TextAlignment.Center);

                    for (int r = 0; r < rows; r++)
                    {
                        var x = items[r];
                        AddCell(grid, r, 1, x.Name, wrap: true, bg: bg);
                        AddCell(grid, r, 2, x.Stb, bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 3, x.Unit, bg: bg, align: TextAlignment.Center);
                        AddCell(grid, r, 4, x.Qty.ToString(), bg: bg, align: TextAlignment.Right);
                        AddCell(grid, r, 5, x.Supplier, wrap: true, bg: bg);
                        AddCell(grid, r, 6, x.Passport, wrap: true, bg: bg);
                    }

                    JvkPanel.Children.Add(grid);
                }
            }
        }











    }
}