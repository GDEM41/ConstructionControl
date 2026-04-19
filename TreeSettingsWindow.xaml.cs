using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace ConstructionControl
{
    public partial class TreeSettingsWindow : Window
    {
        public class MaterialSplitRuleSource
        {
            public string CategoryName { get; set; }
            public string TypeName { get; set; }
            public string SubTypeName { get; set; }
            public string MaterialName { get; set; }
            public string Level4Name { get; set; }
            public string Level5Name { get; set; }
            public string Level6Name { get; set; }
            public bool IsAutoSplitEnabled { get; set; }
        }

        public class MaterialSplitRuleRow : INotifyPropertyChanged
        {
            private readonly string[] segments = new string[10];
            private bool isAutoSplitEnabled;
            private string categoryName;
            private string typeName;
            private string subTypeName;
            private string materialName;
            private string level4Name;
            private string level5Name;
            private string level6Name;
            public string CategoryName { get; set; }
            public string TypeName { get; set; }
            public string SubTypeName { get; set; }
            public string MaterialName { get; set; }
            public string Level4Name { get; set; }
            public string Level5Name { get; set; }
            public string Level6Name { get; set; }
            public string OriginalCategoryName { get; set; }
            public string OriginalTypeName { get; set; }
            public string OriginalSubTypeName { get; set; }
            public string OriginalMaterialName { get; set; }
            public string OriginalLevel4Name { get; set; }
            public string OriginalLevel5Name { get; set; }
            public string OriginalLevel6Name { get; set; }
            public string EditableCategoryName
            {
                get => categoryName;
                set => SetField(ref categoryName, value, nameof(EditableCategoryName));
            }

            public string EditableTypeName
            {
                get => typeName;
                set => SetField(ref typeName, value, nameof(EditableTypeName));
            }

            public string EditableSubTypeName
            {
                get => subTypeName;
                set => SetField(ref subTypeName, value, nameof(EditableSubTypeName));
            }

            public string EditableMaterialName
            {
                get => materialName;
                set
                {
                    if (SetField(ref materialName, value, nameof(EditableMaterialName)) && IsAutoSplitEnabled)
                        ApplyAutomaticSplit();
                }
            }
            public string EditableLevel4Name
            {
                get => level4Name;
                set => SetField(ref level4Name, value, nameof(EditableLevel4Name));
            }

            public string EditableLevel5Name
            {
                get => level5Name;
                set => SetField(ref level5Name, value, nameof(EditableLevel5Name));
            }

            public string EditableLevel6Name
            {
                get => level6Name;
                set => SetField(ref level6Name, value, nameof(EditableLevel6Name));
            }
            public string Segment1 { get => segments[0]; set => SetSegment(0, value); }
            public string Segment2 { get => segments[1]; set => SetSegment(1, value); }
            public string Segment3 { get => segments[2]; set => SetSegment(2, value); }
            public string Segment4 { get => segments[3]; set => SetSegment(3, value); }
            public string Segment5 { get => segments[4]; set => SetSegment(4, value); }
            public string Segment6 { get => segments[5]; set => SetSegment(5, value); }
            public string Segment7 { get => segments[6]; set => SetSegment(6, value); }
            public string Segment8 { get => segments[7]; set => SetSegment(7, value); }
            public string Segment9 { get => segments[8]; set => SetSegment(8, value); }
            public string Segment10 { get => segments[9]; set => SetSegment(9, value); }
            public bool IsAutoSplitEnabled
            {
                get => isAutoSplitEnabled;
                set
                {
                    if (isAutoSplitEnabled == value)
                        return;

                    isAutoSplitEnabled = value;
                    OnPropertyChanged(nameof(IsAutoSplitEnabled));
                    if (isAutoSplitEnabled)
                        ApplyAutomaticSplit();
                }
            }

            public void SetRule(string rule)
            {

                var parts = NormalizeRule(rule)
                    .Split('|', System.StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .ToList();

                for (var i = 0; i < segments.Length; i++)
                    segments[i] = i < parts.Count ? parts[i] : string.Empty;

                OnPropertyChanged(nameof(Segment1));
                OnPropertyChanged(nameof(Segment2));
                OnPropertyChanged(nameof(Segment3));
                OnPropertyChanged(nameof(Segment4));
                OnPropertyChanged(nameof(Segment5));
                OnPropertyChanged(nameof(Segment6));
                OnPropertyChanged(nameof(Segment7));
                OnPropertyChanged(nameof(Segment8));
                OnPropertyChanged(nameof(Segment9));
                OnPropertyChanged(nameof(Segment10));
            }
            public string GetRule() => NormalizeRule(string.Join("|", segments));

            public void ApplyAutomaticSplit()
            {
                SetRule(BuildAutomaticRule(EditableMaterialName));
            }

            private void SetSegment(int index, string value)
            {
                var normalized = (value ?? string.Empty).Trim();
                if (string.Equals(segments[index], normalized, System.StringComparison.CurrentCulture))
                    return;

                segments[index] = normalized;
                OnPropertyChanged($"Segment{index + 1}");
            }

            public event PropertyChangedEventHandler PropertyChanged;
            private bool SetField(ref string field, string value, string propertyName)
            {
                var normalized = NormalizeMetaValue(value);
                if (string.Equals(field, normalized, System.StringComparison.CurrentCulture))
                    return false;

                field = normalized;
                OnPropertyChanged(propertyName);
                return true;
            }


            private void OnPropertyChanged([CallerMemberName] string propertyName = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public class MaterialBindingChange
        {
            public string MaterialName { get; set; }
            public string OldCategoryName { get; set; }
            public string OldTypeName { get; set; }
            public string OldSubTypeName { get; set; }
            public string NewCategoryName { get; set; }
            public string NewTypeName { get; set; }
            public string NewSubTypeName { get; set; }
            public string OldLevel4Name { get; set; }
            public string OldLevel5Name { get; set; }
            public string OldLevel6Name { get; set; }
            public string NewLevel4Name { get; set; }
            public string NewLevel5Name { get; set; }
            public string NewLevel6Name { get; set; }
            public string OldMaterialName { get; set; }
            public string NewMaterialName { get; set; }
        }

        private readonly ObservableCollection<MaterialSplitRuleRow> rows;
        private bool isBulkUpdating;
        private int visibleCatalogColumns = 5;
        private int visibleLevelColumns = 6;
        private ICollectionView rulesView;
        public Dictionary<string, string> ResultRules { get; private set; } = new();
        public List<string> ResultAutoSplitMaterials { get; private set; } = new();
        public List<MaterialBindingChange> ResultBindingChanges { get; private set; } = new();
        public List<MaterialCatalogItem> ResultCatalog { get; private set; } = new();

        public TreeSettingsWindow(IEnumerable<MaterialSplitRuleSource> materials, Dictionary<string, string> existingRules, IEnumerable<string> existingAutoSplitMaterials)
        {
            InitializeComponent();

            var autoSplitSet = new HashSet<string>(
                existingAutoSplitMaterials ?? Enumerable.Empty<string>(),
                System.StringComparer.CurrentCultureIgnoreCase);

            rows = new ObservableCollection<MaterialSplitRuleRow>(
                      materials
                    .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName)
                    && string.Equals(NormalizeMetaValue(x.CategoryName), "Основные", System.StringComparison.CurrentCultureIgnoreCase))
                                        .GroupBy(x => new
                                        {
                                            Material = x.MaterialName,
                                            Category = NormalizeMetaValue(x.CategoryName),
                                            Type = NormalizeMetaValue(x.TypeName),
                                            SubType = NormalizeMetaValue(x.SubTypeName),
                                            Level4 = NormalizeMetaValue(x.Level4Name),
                                            Level5 = NormalizeMetaValue(x.Level5Name),
                                            Level6 = NormalizeMetaValue(x.Level6Name)
                                        })
                    .Select(g => g.First())
                    .OrderBy(x => x.CategoryName)
                    .ThenBy(x => x.TypeName)
                    .ThenBy(x => x.SubTypeName)
                    .ThenBy(x => x.MaterialName)
                    .Select(x => new MaterialSplitRuleRow
                    {
                        CategoryName = "Основные",
                        TypeName = NormalizeMetaValue(x.TypeName),
                        SubTypeName = NormalizeMetaValue(x.SubTypeName),
                        Level4Name = NormalizeMetaValue(x.Level4Name),
                        Level5Name = NormalizeMetaValue(x.Level5Name),
                        Level6Name = NormalizeMetaValue(x.Level6Name),
                        MaterialName = x.MaterialName
                    }));
            foreach (var row in rows)
            {
                row.OriginalCategoryName = row.CategoryName;
                row.OriginalTypeName = row.TypeName;
                row.OriginalSubTypeName = row.SubTypeName;
                row.EditableCategoryName = row.CategoryName;
                row.EditableTypeName = row.TypeName;
                row.OriginalMaterialName = row.MaterialName;
                row.OriginalLevel4Name = row.Level4Name;
                row.OriginalLevel5Name = row.Level5Name;
                row.OriginalLevel6Name = row.Level6Name;
                row.EditableSubTypeName = row.SubTypeName;
                row.EditableLevel4Name = row.Level4Name;
                row.EditableLevel5Name = row.Level5Name;
                row.EditableLevel6Name = row.Level6Name;
                row.EditableMaterialName = row.MaterialName;
                row.IsAutoSplitEnabled = autoSplitSet.Contains(row.MaterialName);

                visibleCatalogColumns = rows.Any(x => !string.IsNullOrWhiteSpace(x.Level6Name)) ? 6 : rows.Any(x => !string.IsNullOrWhiteSpace(x.Level5Name)) ? 5 : rows.Any(x => !string.IsNullOrWhiteSpace(x.Level4Name)) ? 4 : 5;
                if (row.IsAutoSplitEnabled)
                    row.ApplyAutomaticSplit();
                else
                    row.SetRule(existingRules != null && existingRules.TryGetValue(row.MaterialName, out var rule)
                        ? rule
                        : string.Empty);
            }

            var cvs = new CollectionViewSource { Source = rows };
            cvs.GroupDescriptions.Add(new PropertyGroupDescription(nameof(MaterialSplitRuleRow.EditableCategoryName)));
            cvs.GroupDescriptions.Add(new PropertyGroupDescription(nameof(MaterialSplitRuleRow.EditableTypeName)));
            cvs.GroupDescriptions.Add(new PropertyGroupDescription(nameof(MaterialSplitRuleRow.EditableSubTypeName)));
            cvs.IsLiveGroupingRequested = true;
            cvs.LiveGroupingProperties.Add(nameof(MaterialSplitRuleRow.EditableCategoryName));
            cvs.LiveGroupingProperties.Add(nameof(MaterialSplitRuleRow.EditableTypeName));
            cvs.LiveGroupingProperties.Add(nameof(MaterialSplitRuleRow.EditableSubTypeName));
            rulesView = cvs.View;
            RulesGrid.ItemsSource = rulesView;
            visibleLevelColumns = rows.Any() ? System.Math.Max(6, rows.Max(GetUsedSegmentCount)) : 6;
            ApplyCatalogColumnVisibility();
            ApplyLevelColumnVisibility();
            Closing += TreeSettingsWindow_Closing;
        }

        private bool isAutoSaving;
        private void TreeSettingsWindow_Closing(object sender, CancelEventArgs e)
        {
            if (isAutoSaving)
                return;

            isAutoSaving = true;
            SaveChanges(closeWindow: false);
        }

        private static int GetUsedSegmentCount(MaterialSplitRuleRow row)
        {
            for (var i = 10; i >= 1; i--)
            {
                var value = i switch
                {
                    1 => row.Segment1,
                    2 => row.Segment2,
                    3 => row.Segment3,
                    4 => row.Segment4,
                    5 => row.Segment5,
                    6 => row.Segment6,
                    7 => row.Segment7,
                    8 => row.Segment8,
                    9 => row.Segment9,
                    _ => row.Segment10
                };

                if (!string.IsNullOrWhiteSpace(value))
                    return i;
            }

            return 0;
        }

        private List<MaterialSplitRuleRow> GetSelectedRows()
        {
            var selected = RulesGrid.SelectedItems
                .Cast<object>()
                .OfType<MaterialSplitRuleRow>()
                .Distinct()
                .ToList();

            if (selected.Count == 0 && RulesGrid.SelectedItem is MaterialSplitRuleRow current)
                selected.Add(current);

            return selected;
        }

        private List<MaterialSplitRuleRow> GetSelectedTargets(MaterialSplitRuleRow source)
        {
            var selected = GetSelectedRows();
            if (selected.Count < 2)
                return new List<MaterialSplitRuleRow>();

            return selected
                .Where(x => !ReferenceEquals(x, source))
                .ToList();
        }

        private bool ApplyToSelectedRows(MaterialSplitRuleRow source, System.Action<MaterialSplitRuleRow> apply)
        {
            var targets = GetSelectedTargets(source);
            if (targets.Count == 0)
                return false;

            isBulkUpdating = true;
            try
            {
                foreach (var target in targets)
                    apply(target);
            }
            finally
            {
                isBulkUpdating = false;
            }

            return true;
        }

        private static string GetColumnBindingPath(DataGridColumn column)
        {
            if (column is DataGridBoundColumn boundColumn && boundColumn.Binding is Binding binding)
                return binding.Path?.Path ?? string.Empty;

            return string.Empty;
        }

        private static bool IsGroupColumn(string path)
        {
            return string.Equals(path, nameof(MaterialSplitRuleRow.EditableCategoryName), System.StringComparison.Ordinal)
                   || string.Equals(path, nameof(MaterialSplitRuleRow.EditableTypeName), System.StringComparison.Ordinal)
                   || string.Equals(path, nameof(MaterialSplitRuleRow.EditableSubTypeName), System.StringComparison.Ordinal);
        }

        private void RefreshRulesView()
        {
            rulesView?.Refresh();
        }

        private bool ApplyCellValueToSelectedRows(MaterialSplitRuleRow source, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            bool changed = path switch
            {
                nameof(MaterialSplitRuleRow.EditableCategoryName) => ApplyToSelectedRows(source, target => target.EditableCategoryName = source.EditableCategoryName),
                nameof(MaterialSplitRuleRow.EditableTypeName) => ApplyToSelectedRows(source, target => target.EditableTypeName = source.EditableTypeName),
                nameof(MaterialSplitRuleRow.EditableSubTypeName) => ApplyToSelectedRows(source, target => target.EditableSubTypeName = source.EditableSubTypeName),
                nameof(MaterialSplitRuleRow.EditableLevel4Name) => ApplyToSelectedRows(source, target => target.EditableLevel4Name = source.EditableLevel4Name),
                nameof(MaterialSplitRuleRow.EditableLevel5Name) => ApplyToSelectedRows(source, target => target.EditableLevel5Name = source.EditableLevel5Name),
                nameof(MaterialSplitRuleRow.EditableLevel6Name) => ApplyToSelectedRows(source, target => target.EditableLevel6Name = source.EditableLevel6Name),
                nameof(MaterialSplitRuleRow.EditableMaterialName) => ApplyToSelectedRows(source, target => target.EditableMaterialName = source.EditableMaterialName),
                nameof(MaterialSplitRuleRow.IsAutoSplitEnabled) => ApplyToSelectedRows(source, target => target.IsAutoSplitEnabled = source.IsAutoSplitEnabled),
                nameof(MaterialSplitRuleRow.Segment1) => ApplyToSelectedRows(source, target => target.Segment1 = source.Segment1),
                nameof(MaterialSplitRuleRow.Segment2) => ApplyToSelectedRows(source, target => target.Segment2 = source.Segment2),
                nameof(MaterialSplitRuleRow.Segment3) => ApplyToSelectedRows(source, target => target.Segment3 = source.Segment3),
                nameof(MaterialSplitRuleRow.Segment4) => ApplyToSelectedRows(source, target => target.Segment4 = source.Segment4),
                nameof(MaterialSplitRuleRow.Segment5) => ApplyToSelectedRows(source, target => target.Segment5 = source.Segment5),
                nameof(MaterialSplitRuleRow.Segment6) => ApplyToSelectedRows(source, target => target.Segment6 = source.Segment6),
                nameof(MaterialSplitRuleRow.Segment7) => ApplyToSelectedRows(source, target => target.Segment7 = source.Segment7),
                nameof(MaterialSplitRuleRow.Segment8) => ApplyToSelectedRows(source, target => target.Segment8 = source.Segment8),
                nameof(MaterialSplitRuleRow.Segment9) => ApplyToSelectedRows(source, target => target.Segment9 = source.Segment9),
                nameof(MaterialSplitRuleRow.Segment10) => ApplyToSelectedRows(source, target => target.Segment10 = source.Segment10),
                _ => false
            };

            if (changed && IsGroupColumn(path))
                RefreshRulesView();

            return changed;
        }

        private void AddEntry_Click(object sender, RoutedEventArgs e)
        {
            var selected = RulesGrid.SelectedItem as MaterialSplitRuleRow;
            var newRow = new MaterialSplitRuleRow
            {
                EditableCategoryName = "Основные",
                EditableTypeName = selected?.EditableTypeName ?? string.Empty,
                EditableSubTypeName = selected?.EditableSubTypeName ?? string.Empty,
                EditableMaterialName = string.Empty,
                EditableLevel4Name = selected?.EditableLevel4Name ?? string.Empty,
                EditableLevel5Name = selected?.EditableLevel5Name ?? string.Empty,
                EditableLevel6Name = selected?.EditableLevel6Name ?? string.Empty,
                CategoryName = "Основные",
                TypeName = selected?.EditableTypeName ?? string.Empty,
                SubTypeName = selected?.EditableSubTypeName ?? string.Empty,
                Level4Name = selected?.EditableLevel4Name ?? string.Empty,
                Level5Name = selected?.EditableLevel5Name ?? string.Empty,
                Level6Name = selected?.EditableLevel6Name ?? string.Empty,
                MaterialName = string.Empty,
                OriginalCategoryName = string.Empty,
                OriginalTypeName = string.Empty,
                OriginalSubTypeName = string.Empty,
                OriginalMaterialName = string.Empty,
                OriginalLevel4Name = string.Empty,
                OriginalLevel5Name = string.Empty,
                OriginalLevel6Name = string.Empty,
                IsAutoSplitEnabled = false
            };

            var index = selected != null ? rows.IndexOf(selected) + 1 : rows.Count;
            rows.Insert(index, newRow);
            RulesGrid.SelectedItem = newRow;
            RulesGrid.ScrollIntoView(newRow);
        }

        private void RemoveEntry_Click(object sender, RoutedEventArgs e)
        {
            if (RulesGrid.SelectedItem is not MaterialSplitRuleRow selected)
                return;

            rows.Remove(selected);
        }

        private void ApplyTypeToSelection_Click(object sender, RoutedEventArgs e)
        {
            if (RulesGrid.SelectedItem is not MaterialSplitRuleRow selected)
                return;

            if (!ApplyToSelectedRows(selected, row => row.EditableTypeName = selected.EditableTypeName))
            {
                MessageBox.Show("Выделите несколько строк (Shift + клик), чтобы применить тип к диапазону.");
                return;
            }

            RefreshRulesView();
        }

        private void ApplySubTypeToSelection_Click(object sender, RoutedEventArgs e)
        {
            if (RulesGrid.SelectedItem is not MaterialSplitRuleRow selected)
                return;

            if (!ApplyToSelectedRows(selected, row => row.EditableSubTypeName = selected.EditableSubTypeName))
            {
                MessageBox.Show("Выделите несколько строк (Shift + клик), чтобы применить подтип к диапазону.");
                return;
            }

            RefreshRulesView();
        }

        private void ApplyAutoSplitToSelection_Click(object sender, RoutedEventArgs e)
        {
            if (RulesGrid.SelectedItem is not MaterialSplitRuleRow selected)
                return;

            if (!ApplyToSelectedRows(selected, row => row.IsAutoSplitEnabled = selected.IsAutoSplitEnabled))
            {
                MessageBox.Show("Выделите несколько строк (Shift + клик), чтобы применить авторазбиение к диапазону.");
            }
        }

        private void RenameTypeForAll_Click(object sender, RoutedEventArgs e)
        {
            if (RulesGrid.SelectedItem is not MaterialSplitRuleRow selected)
                return;

            var oldType = NormalizeMetaValue(selected.EditableTypeName);
            if (string.IsNullOrWhiteSpace(oldType))
                return;

            var dialog = new Window
            {
                Title = "Переименовать тип",
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                SizeToContent = SizeToContent.WidthAndHeight
            };

            var box = new TextBox { Text = oldType, MinWidth = 260, Margin = new Thickness(0, 8, 0, 0) };
            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(new TextBlock { Text = $"Новое название типа для всех материалов с типом \"{oldType}\":" });
            root.Children.Add(box);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            var cancel = new Button { Content = "Отмена", Width = 90, Margin = new Thickness(0, 0, 8, 0), IsCancel = true };
            var ok = new Button { Content = "Применить", Width = 90, IsDefault = true };
            ok.Click += (_, _) => dialog.DialogResult = true;
            buttons.Children.Add(cancel);
            buttons.Children.Add(ok);
            root.Children.Add(buttons);
            dialog.Content = root;

            if (dialog.ShowDialog() != true)
                return;

            var newType = NormalizeMetaValue(box.Text);
            if (string.IsNullOrWhiteSpace(newType) || string.Equals(newType, oldType, System.StringComparison.CurrentCultureIgnoreCase))
                return;

            foreach (var row in rows.Where(r => string.Equals(NormalizeMetaValue(r.EditableTypeName), oldType, System.StringComparison.CurrentCultureIgnoreCase)))
                row.EditableTypeName = newType;
        }
        private void ApplySubTypeToOthers_Click(object sender, RoutedEventArgs e)
        {
            if (RulesGrid.SelectedItem is not MaterialSplitRuleRow selected)
                return;

            var subtype = NormalizeMetaValue(selected.EditableSubTypeName);
            var candidates = rows
                .Where(r => !ReferenceEquals(r, selected)
                    && string.Equals(NormalizeMetaValue(r.EditableTypeName), NormalizeMetaValue(selected.EditableTypeName), System.StringComparison.CurrentCultureIgnoreCase))
                .OrderBy(r => r.EditableMaterialName)
                .ToList();

            if (candidates.Count == 0)
                return;

            var targets = PromptRowsSelection(candidates, "Выберите материалы для применения подтипа");
            foreach (var row in targets)
                row.EditableSubTypeName = subtype;
        }

        private List<MaterialSplitRuleRow> PromptRowsSelection(List<MaterialSplitRuleRow> candidates, string title)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock { Text = title, Margin = new Thickness(0, 0, 0, 8) });
            var checks = new List<(MaterialSplitRuleRow Row, CheckBox Check)>();
            var list = new StackPanel();
            foreach (var c in candidates)
            {
                var cb = new CheckBox { Content = c.EditableMaterialName, Margin = new Thickness(0, 2, 0, 2) };
                list.Children.Add(cb);
                checks.Add((c, cb));
            }
            panel.Children.Add(new ScrollViewer { Height = 240, Content = list });
            var wnd = new Window { Title = "Выбор материалов", Owner = this, Content = panel, SizeToContent = SizeToContent.WidthAndHeight, WindowStartupLocation = WindowStartupLocation.CenterOwner };
            var ok = new Button { Content = "Применить", Width = 90, Margin = new Thickness(0, 8, 0, 0), IsDefault = true, HorizontalAlignment = HorizontalAlignment.Right };
            ok.Click += (_, __) => wnd.DialogResult = true;
            panel.Children.Add(ok);
            if (wnd.ShowDialog() != true)
                return new List<MaterialSplitRuleRow>();
            return checks.Where(x => x.Check.IsChecked == true).Select(x => x.Row).ToList();
        }
        private void AddLevelColumn_Click(object sender, RoutedEventArgs e)
        {
            if (visibleCatalogColumns < 6)
            {
                visibleCatalogColumns++;
                ApplyCatalogColumnVisibility();
                return;
            }
            if (visibleLevelColumns >= 10)
                return;

            visibleLevelColumns++;
            ApplyLevelColumnVisibility();
        }

        private void RemoveLevelColumn_Click(object sender, RoutedEventArgs e)
        {
            if (visibleLevelColumns > 6)
            {
                visibleLevelColumns--;
                ApplyLevelColumnVisibility();
                return;
            }

            if (visibleCatalogColumns <= 3)
                return;

            visibleCatalogColumns--;
            ApplyCatalogColumnVisibility();
        }


        private void ApplyCatalogColumnVisibility()
        {
            CatalogLevel4Column.Visibility = visibleCatalogColumns >= 4 ? Visibility.Visible : Visibility.Collapsed;
            CatalogLevel5Column.Visibility = visibleCatalogColumns >= 5 ? Visibility.Visible : Visibility.Collapsed;
            CatalogLevel6Column.Visibility = visibleCatalogColumns >= 6 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ApplyLevelColumnVisibility()
        {
            LevelColumn7.Visibility = visibleLevelColumns >= 7 ? Visibility.Visible : Visibility.Collapsed;
            LevelColumn8.Visibility = visibleLevelColumns >= 8 ? Visibility.Visible : Visibility.Collapsed;
            LevelColumn9.Visibility = visibleLevelColumns >= 9 ? Visibility.Visible : Visibility.Collapsed;
            LevelColumn10.Visibility = visibleLevelColumns >= 10 ? Visibility.Visible : Visibility.Collapsed;
        }
        private void RulesGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (isBulkUpdating || e.EditAction != DataGridEditAction.Commit)
                return;

            if (e.Row?.Item is not MaterialSplitRuleRow edited)
                return;

            var bindingPath = GetColumnBindingPath(e.Column);
            var selectedRows = GetSelectedRows();
            if (selectedRows.Count > 1)
                ApplyCellValueToSelectedRows(edited, bindingPath);

            if (IsGroupColumn(bindingPath))
                RefreshRulesView();

            if (string.IsNullOrWhiteSpace(bindingPath) || !bindingPath.StartsWith("Segment", System.StringComparison.Ordinal))
                return;

            var rowsToNormalize = selectedRows.Count > 0
                ? selectedRows
                : new List<MaterialSplitRuleRow> { edited };

            foreach (var row in rowsToNormalize)
            {
                var normalizedRule = row.GetRule();
                if (!string.IsNullOrWhiteSpace(normalizedRule))
                    row.SetRule(normalizedRule);
            }
        }

        private void ApplySplitToOthers_Click(object sender, RoutedEventArgs e)
        {
            if (RulesGrid.SelectedItem is not MaterialSplitRuleRow edited)
                return;

            var normalizedRule = edited.GetRule();
            if (string.IsNullOrWhiteSpace(normalizedRule))
            {
                MessageBox.Show("Сначала задайте уровни разбиения для выбранного материала.");
                return;
            }

            var targets = GetSelectedTargets(edited);
            if (targets.Count == 0)
                targets = PromptRuleTargets(edited);

            if (targets == null || targets.Count == 0)
                return;

            var sourcePattern = BuildRulePattern(edited.EditableMaterialName, normalizedRule);
            if (sourcePattern == null || sourcePattern.Count == 0)
            {
                MessageBox.Show("Не удалось построить шаблон разбиения для выбранного материала.");
                return;
            }


            isBulkUpdating = true;
            try
            {
                foreach (var target in targets)
                {
                    if (ReferenceEquals(target, edited))
                        continue;

                    var convertedRule = ApplyRuleByPattern(target.EditableMaterialName, sourcePattern);
                    if (!string.IsNullOrWhiteSpace(convertedRule))
                        target.SetRule(convertedRule);
                }
            }
            finally
            {
                isBulkUpdating = false;
            }
        }
        private List<MaterialSplitRuleRow> PromptRuleTargets(MaterialSplitRuleRow edited)
        {
            var candidates = rows
                .Where(x => !ReferenceEquals(x, edited)
                            && string.Equals(x.EditableCategoryName, edited.EditableCategoryName, System.StringComparison.CurrentCultureIgnoreCase)
                            && string.Equals(x.EditableTypeName, edited.EditableTypeName, System.StringComparison.CurrentCultureIgnoreCase)
                            && string.Equals(x.EditableSubTypeName, edited.EditableSubTypeName, System.StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            if (candidates.Count == 0)
                return new List<MaterialSplitRuleRow>();

            var panel = new StackPanel();

            panel.Children.Add(new TextBlock
            {
                Text = $"Применить разбиение к другим элементам ({edited.EditableCategoryName} / {edited.EditableTypeName} / {edited.EditableSubTypeName})?",
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap
            });

            var checks = new List<(MaterialSplitRuleRow Row, CheckBox Check)>();

            var scroll = new ScrollViewer
            {
                Height = 240,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = new StackPanel()
            };

            foreach (var candidate in candidates.OrderBy(x => x.MaterialName))
            {
                var check = new CheckBox
                {
                    Content = candidate.MaterialName,
                    Margin = new Thickness(0, 2, 0, 2),
                    IsChecked = false
                };

                ((StackPanel)scroll.Content).Children.Add(check);
                checks.Add((candidate, check));
            }

            panel.Children.Add(scroll);

            var selectionWindow = new Window
            {
                Title = "Применение правила",
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                SizeToContent = SizeToContent.WidthAndHeight,
                Content = new DockPanel
                {
                    Margin = new Thickness(12)
                }
            };

            var dock = (DockPanel)selectionWindow.Content;
            DockPanel.SetDock(panel, Dock.Top);
            dock.Children.Add(panel);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };

            var cancel = new Button { Content = "Отмена", Width = 95, Margin = new Thickness(0, 0, 8, 0), IsCancel = true };
            var ok = new Button { Content = "Применить", Width = 95, IsDefault = true };

            ok.Click += (_, _) => selectionWindow.DialogResult = true;

            buttons.Children.Add(cancel);
            buttons.Children.Add(ok);
            DockPanel.SetDock(buttons, Dock.Bottom);
            dock.Children.Add(buttons);

            if (selectionWindow.ShowDialog() != true)
                return new List<MaterialSplitRuleRow>();

            return checks
                .Where(x => x.Check.IsChecked == true)
                .Select(x => x.Row)
                .ToList();
        }

        private static List<int> BuildRulePattern(string sourceMaterialName, string normalizedRule)
        {
            var materialTokens = GetPatternTokens(sourceMaterialName);
            if (materialTokens.Count == 0)
                return null;

            var segmentDefinitions = normalizedRule
                .Split('|', System.StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            if (segmentDefinitions.Count == 0)
                return null;

            var tokenIndex = 0;
            var pattern = new List<int>();

            foreach (var segment in segmentDefinitions)
            {
                var segmentTokens = GetPatternTokens(segment);
                if (segmentTokens.Count == 0)
                    return null;

                var segmentCanonical = string.Concat(segmentTokens);
                var consumed = 0;
                var assembled = string.Empty;

                while (tokenIndex + consumed < materialTokens.Count)
                {
                    assembled += materialTokens[tokenIndex + consumed];
                    consumed++;

                    if (string.Equals(assembled, segmentCanonical, System.StringComparison.CurrentCultureIgnoreCase))
                        break;
                }

                if (!string.Equals(assembled, segmentCanonical, System.StringComparison.CurrentCultureIgnoreCase))
                    return null;

                pattern.Add(consumed);
                tokenIndex += consumed;
            }

            return pattern;
        }

        private static string ApplyRuleByPattern(string targetMaterialName, List<int> pattern)
        {
            if (pattern == null || pattern.Count == 0)
                return string.Empty;

            var targetTokens = GetPatternTokens(targetMaterialName);
            if (targetTokens.Count == 0)
                return string.Empty;

            var consumed = 0;
            var segments = new List<string>();

            foreach (var chunkSize in pattern)
            {
                if (chunkSize <= 0 || consumed + chunkSize > targetTokens.Count)
                    return string.Empty;

                segments.Add(string.Concat(targetTokens.Skip(consumed).Take(chunkSize)));
                consumed += chunkSize;
            }

            return string.Join("|", segments);
        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveChanges(closeWindow: true);
        }

        private void SaveChanges(bool closeWindow)
        {
            foreach (var row in rows)
            {
                row.CategoryName = "Основные";
                row.EditableCategoryName = "Основные";
                row.TypeName = NormalizeMetaValue(row.EditableTypeName);
                row.SubTypeName = NormalizeMetaValue(row.EditableSubTypeName);
                row.Level4Name = NormalizeMetaValue(row.EditableLevel4Name);
                row.Level5Name = NormalizeMetaValue(row.EditableLevel5Name);
                row.Level6Name = NormalizeMetaValue(row.EditableLevel6Name);
                row.MaterialName = NormalizeMetaValue(row.EditableMaterialName);
            }


            var validRows = rows
                    .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName)
                    && string.Equals(NormalizeMetaValue(x.CategoryName), "Основные", System.StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            foreach (var row in validRows.Where(x => x.IsAutoSplitEnabled))
                row.ApplyAutomaticSplit();

            var ruleRows = validRows
                .Select(x => new { MaterialName = x.MaterialName, Rule = x.GetRule() })
                .Where(x => !string.IsNullOrWhiteSpace(x.Rule))
                .ToList();

            var duplicateRuleMaterials = ruleRows
                .GroupBy(x => x.MaterialName, System.StringComparer.CurrentCultureIgnoreCase)
                .Where(x => x.Count() > 1)
                .Select(x => x.Key)
                .OrderBy(x => x, System.StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (duplicateRuleMaterials.Count > 0)
            {
                var preview = string.Join(Environment.NewLine, duplicateRuleMaterials.Take(8));
                var suffix = duplicateRuleMaterials.Count > 8
                    ? $"{Environment.NewLine}... и еще {duplicateRuleMaterials.Count - 8}"
                    : string.Empty;

                MessageBox.Show(
                    $"Найдены дубли материалов в правилах разбиения.{Environment.NewLine}{Environment.NewLine}{preview}{suffix}{Environment.NewLine}{Environment.NewLine}Уберите дубли и сохраните снова.",
                    "Дубли материалов",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            ResultRules = ruleRows.ToDictionary(
                x => x.MaterialName,
                x => x.Rule,
                System.StringComparer.CurrentCultureIgnoreCase);
            ResultAutoSplitMaterials = validRows
                .Where(x => x.IsAutoSplitEnabled)
                .Select(x => x.MaterialName)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(System.StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, System.StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            ResultBindingChanges = validRows
                .Where(x => !string.Equals(x.OriginalCategoryName, x.CategoryName, System.StringComparison.CurrentCulture)
                         || !string.Equals(x.OriginalTypeName, x.TypeName, System.StringComparison.CurrentCulture)
                         || !string.Equals(x.OriginalSubTypeName, x.SubTypeName, System.StringComparison.CurrentCulture)
                         || !string.Equals(x.OriginalLevel4Name, x.Level4Name, System.StringComparison.CurrentCulture)
                         || !string.Equals(x.OriginalLevel5Name, x.Level5Name, System.StringComparison.CurrentCulture)
                         || !string.Equals(x.OriginalLevel6Name, x.Level6Name, System.StringComparison.CurrentCulture)
                         || !string.Equals(x.OriginalMaterialName, x.MaterialName, System.StringComparison.CurrentCulture))
                .Select(x => new MaterialBindingChange
                {
                    MaterialName = x.MaterialName,
                    OldCategoryName = x.OriginalCategoryName,
                    OldTypeName = x.OriginalTypeName,
                    OldSubTypeName = x.OriginalSubTypeName,
                    OldLevel4Name = x.OriginalLevel4Name,
                    OldLevel5Name = x.OriginalLevel5Name,
                    OldLevel6Name = x.OriginalLevel6Name,
                    NewCategoryName = x.CategoryName,
                    NewTypeName = x.TypeName,
                    NewSubTypeName = x.SubTypeName,
                    NewLevel4Name = x.Level4Name,
                    NewLevel5Name = x.Level5Name,
                    NewLevel6Name = x.Level6Name,
                    OldMaterialName = x.OriginalMaterialName,
                    NewMaterialName = x.MaterialName
                })
                .ToList();

            ResultCatalog = validRows
                                      .Select(x => new MaterialCatalogItem
                                      {
                                          CategoryName = x.CategoryName,
                                          TypeName = x.TypeName,
                                          SubTypeName = x.SubTypeName,
                                          ExtraLevels = new List<string> { x.Level4Name, x.Level5Name, x.Level6Name }
                                              .Where(v => !string.IsNullOrWhiteSpace(v)).ToList(),
                                          MaterialName = x.MaterialName
                                      })
                     .GroupBy(x => new
                     {
                         Category = x.CategoryName ?? string.Empty,
                         Type = x.TypeName ?? string.Empty,
                         SubType = x.SubTypeName ?? string.Empty,
                         Level4 = x.ExtraLevels != null && x.ExtraLevels.Count > 0 ? x.ExtraLevels[0] : string.Empty,
                         Level5 = x.ExtraLevels != null && x.ExtraLevels.Count > 1 ? x.ExtraLevels[1] : string.Empty,
                         Level6 = x.ExtraLevels != null && x.ExtraLevels.Count > 2 ? x.ExtraLevels[2] : string.Empty,
                         Material = x.MaterialName ?? string.Empty
                     })
                     .Select(x => x.First())
                     .OrderBy(x => x.CategoryName)
                     .ThenBy(x => x.TypeName)
                     .ThenBy(x => x.SubTypeName)
                     .ThenBy(x => x.MaterialName)
                     .ToList();

            if (!closeWindow)
                return;

            DialogResult = true;
            Close();
        }

        private static string NormalizeMetaValue(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
                return string.Empty;

            var value = rawValue.Trim();
            if (value.StartsWith("(без ", true, CultureInfo.CurrentCulture))
                return string.Empty;

            return value;
        }

        private static string NormalizeRule(string rawRule)
        {
            if (string.IsNullOrWhiteSpace(rawRule))
                return string.Empty;

            var parts = rawRule
                .Split('|', System.StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x));

            return string.Join("|", parts);
        }

        private static string BuildAutomaticRule(string materialName)
        {
            if (string.IsNullOrWhiteSpace(materialName))
                return string.Empty;

            var parts = Regex.Matches(materialName, @"[A-Za-zА-Яа-яЁё]+|\d+")
                .Select(x => x.Value.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            return parts.Count == 0 ? string.Empty : string.Join("|", parts);
        }

        private static List<string> GetPatternTokens(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return new List<string>();

            return Regex.Matches(value, @"[A-Za-zА-Яа-яЁё]+|\d+(?:[\.,]\d+)?|[^A-Za-zА-Яа-яЁё0-9\s]")
                .Select(x => x.Value)
                .ToList();
        }
    }
}
