using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;

namespace ConstructionControl
{
    public partial class MainWindow : Window
    {
        private const string SaveFileName = "data.json";

        private ProjectObject currentObject;
        private List<JournalRecord> journal = new();
        private List<JournalRecord> filteredJournal = new();

        private bool isLocked;

        public MainWindow()
        {
            InitializeComponent();

            LoadState();

            ArrivalPanel.ArrivalAdded += OnArrivalAdded;

            if (currentObject != null)
                ArrivalPanel.SetObject(currentObject);

            RefreshTreePreserveState();
            RefreshFilters();
            ApplyAllFilters();
        }

        // ================= МЕНЮ =================

        private void CreateObject_Click(object sender, RoutedEventArgs e)
        {
            var w = new CreateObjectWindow { Owner = this };
            if (w.ShowDialog() == true)
            {
                currentObject = new ProjectObject { Name = w.ObjectName };
                journal.Clear();

                ArrivalPanel.SetObject(currentObject);

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
            JournalGrid.IsReadOnly = true;
        }

        private void LockButton_Unchecked(object sender, RoutedEventArgs e)
        {
            isLocked = false;
            JournalGrid.IsReadOnly = false;
        }

        // ================= ПРИХОД =================

        private void OnArrivalAdded(Arrival arrival)
        {
            foreach (var i in arrival.Items)
            {
                journal.Add(new JournalRecord
                {
                    Date = i.Date,
                    ObjectName = currentObject.Name,
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
        }

        // ================= ДЕРЕВО =================

        private void RefreshTreePreserveState()
        {
            var expandedGroups = new HashSet<string>();
            string selectedGroup = null;
            string selectedMaterial = null;

            if (ObjectsTree.Items.Count > 0 &&
                ObjectsTree.Items[0] is TreeViewItem root)
            {
                foreach (TreeViewItem group in root.Items)
                {
                    if (group.IsExpanded)
                        expandedGroups.Add(group.Header.ToString());

                    foreach (TreeViewItem mat in group.Items)
                    {
                        if (mat.IsSelected)
                        {
                            selectedGroup = group.Header.ToString();
                            selectedMaterial = mat.Header.ToString();
                        }
                    }

                    if (group.IsSelected)
                        selectedGroup = group.Header.ToString();
                }
            }

            // пересборка
            ObjectsTree.Items.Clear();
            if (currentObject == null)
                return;

            var newRoot = new TreeViewItem
            {
                Header = currentObject.Name,
                Tag = "Object",
                IsExpanded = true
            };

            foreach (var g in currentObject.MaterialGroups)
            {
                var groupNode = new TreeViewItem
                {
                    Header = g.Name,
                    Tag = "Group",
                    IsExpanded = expandedGroups.Contains(g.Name)
                };

                if (currentObject.MaterialNamesByGroup.TryGetValue(g.Name, out var names))
                {
                    foreach (var m in names)
                    {
                        var matNode = new TreeViewItem
                        {
                            Header = m,
                            Tag = "Material",
                            IsSelected = g.Name == selectedGroup && m == selectedMaterial
                        };

                        groupNode.Items.Add(matNode);
                    }
                }

                if (g.Name == selectedGroup && selectedMaterial == null)
                    groupNode.IsSelected = true;

                newRoot.Items.Add(groupNode);
            }

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

            if (!string.IsNullOrWhiteSpace(SupplierFilterBox.Text))
                data = data.Where(j =>
                    j.Supplier != null &&
                    j.Supplier.Contains(SupplierFilterBox.Text,
                        StringComparison.OrdinalIgnoreCase));

                     filteredJournal = data.ToList();
                     JournalGrid.ItemsSource = filteredJournal;

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

            importWindow.ShowDialog();

        }




        private void RefreshSummaryTable()
        {
            if (currentObject == null)
                return;

            var result = new List<SummaryRow>();

            // группируем ЖВК по материалу
            var materials = journal
                .GroupBy(j => new { j.MaterialName, j.Unit });

            foreach (var mat in materials)
            {
                var row = new SummaryRow
                {
                    MaterialName = mat.Key.MaterialName,
                    Unit = mat.Key.Unit
                };

                // инициализируем блоки
                for (int block = 1; block <= currentObject.BlocksCount; block++)
                    row.ByBlocks[block] = 0;

                // суммируем
                foreach (var rec in mat)
                {
                    // ПОКА: считаем всё как блок 1
                    // (этажи и реальное распределение сделаем позже)
                    row.ByBlocks[1] += rec.Quantity;
                    row.Total += rec.Quantity;
                }

                result.Add(row);
            }

            SummaryGrid.ItemsSource = result;
        }




    }
}