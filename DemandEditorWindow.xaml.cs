using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ConstructionControl
{
    public partial class DemandEditorWindow : Window
    {
        public class DemandMaterialRow
        {
            public string Group { get; set; }
            public string Material { get; set; }
            public string Unit { get; set; }
        }

        private readonly ProjectObject currentObject;
        private readonly List<DemandMaterialRow> rows;
        private readonly Dictionary<string, List<(int Block, string Mark)>> markColumnsByGroup = new();
        private readonly Dictionary<string, List<TextBox>> editorsByGroup = new();
        private readonly Dictionary<string, int> columnsCountByGroup = new();

        public DemandEditorWindow(ProjectObject projectObject, IEnumerable<DemandMaterialRow> sourceRows)
        {
            InitializeComponent();
            currentObject = projectObject;
            rows = sourceRows.OrderBy(x => x.Group).ThenBy(x => x.Material).ToList();
            RenderTabs();
        }

        private void RenderTabs()
        {
            TypeTabs.Items.Clear();
            editorsByGroup.Clear();
            columnsCountByGroup.Clear();

            foreach (var group in rows.Select(r => r.Group).Distinct().OrderBy(x => x))
            {
                var grid = BuildGroupGrid(group);
                var tab = new TabItem
                {
                    Header = group,
                    Content = new ScrollViewer
                    {
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                        HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                        Content = grid
                    }
                };
                TypeTabs.Items.Add(tab);
            }

            if (TypeTabs.Items.Count > 0)
                TypeTabs.SelectedIndex = 0;
        }

        private Grid BuildGroupGrid(string group)
        {
            var groupRows = rows.Where(x => x.Group == group).ToList();
            var markColumns = BuildColumns(group);
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(280) });
            foreach (var _ in markColumns)
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 64 });

            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            AddCell(grid, 0, 0, "Материал", true, "#E2E8F0", true);
            Grid.SetRowSpan(grid.Children[^1], 2);

            int startColumn = 1;
            foreach (var blockGroup in markColumns.GroupBy(x => x.Block))
            {
                AddCell(grid, 0, startColumn, $"Блок {blockGroup.Key}", true, "#CBD5E1", true, HorizontalAlignment.Center);
                Grid.SetColumnSpan(grid.Children[^1], blockGroup.Count());

                int offset = 0;
                foreach (var column in blockGroup)
                {
                    AddCell(grid, 1, startColumn + offset, column.Mark, true, "#E2E8F0", true, HorizontalAlignment.Center);
                    offset++;
                }

                startColumn += blockGroup.Count();
            }

            var editors = new List<TextBox>();
            for (int rowIndex = 0; rowIndex < groupRows.Count; rowIndex++)
            {
                var row = groupRows[rowIndex];
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var bg = rowIndex % 2 == 0 ? "White" : "#F8FAFC";
                AddCell(grid, rowIndex + 2, 0, row.Material, false, bg, false);

                string key = $"{row.Group}::{row.Material}";
                var demand = GetOrCreateDemand(key, row.Unit);
                for (int i = 0; i < markColumns.Count; i++)
                {
                    var col = markColumns[i];
                    var value = GetValue(demand.Levels, col.Block, col.Mark);
                    var box = new TextBox
                    {
                        Text = Format(value),
                        MinWidth = 58,
                        Margin = new Thickness(2),
                        Padding = new Thickness(4, 2, 4, 2),
                        HorizontalContentAlignment = HorizontalAlignment.Right,
                        Tag = (key, row.Unit, col.Block, col.Mark),
                        BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1")),
                        BorderThickness = new Thickness(1)
                    };
                    box.PreviewKeyDown += Cell_PreviewKeyDown;
                    Grid.SetRow(box, rowIndex + 2);
                    Grid.SetColumn(box, i + 1);
                    grid.Children.Add(box);
                    editors.Add(box);
                }
            }

            editorsByGroup[group] = editors;
            columnsCountByGroup[group] = markColumns.Count;
            return grid;
        }

        private List<(int Block, string Mark)> BuildColumns(string group)
        {
            if (markColumnsByGroup.TryGetValue(group, out var existing))
                return existing;

            var columns = new List<(int Block, string Mark)>();
            var marks = LevelMarkHelper.GetMarksForGroup(currentObject, group);

            for (int b = 1; b <= currentObject.BlocksCount; b++)
            {
                foreach (var mark in marks)
                    columns.Add((b, mark));
            }

            markColumnsByGroup[group] = columns;
            return columns;
        }
        private void Cell_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (sender is not TextBox current)
                return;

            int index = -1;
            List<TextBox> list = null;
            string groupName = null;
            foreach (var kv in editorsByGroup)
            {
                index = kv.Value.IndexOf(current);
                if (index >= 0)
                {
                    list = kv.Value;
                    groupName = kv.Key;
                    break;
                }
            }

            if (list == null || groupName == null || !columnsCountByGroup.TryGetValue(groupName, out int columnsCount) || columnsCount <= 0)
                return;

            int currentRow = index / columnsCount;
            int currentCol = index % columnsCount;
            int rowsCount = (int)Math.Ceiling((double)list.Count / columnsCount);

            int targetRow = currentRow;
            int targetCol = currentCol;

            if (e.Key == Key.Right)
                targetCol = Math.Min(columnsCount - 1, currentCol + 1);
            else if (e.Key == Key.Left)
                targetCol = Math.Max(0, currentCol - 1);
            else if (e.Key == Key.Down || e.Key == Key.Enter)
                targetRow = Math.Min(rowsCount - 1, currentRow + 1);
            else if (e.Key == Key.Up)
                targetRow = Math.Max(0, currentRow - 1);
            else
                return;

            int target = targetRow * columnsCount + targetCol;
            if (target >= list.Count)
                target = list.Count - 1;

            if (target != index)
            {
                e.Handled = true;
                list[target].Focus();
                list[target].SelectAll();
            }
        }




        private MaterialDemand GetOrCreateDemand(string key, string unit)
        {
            if (!currentObject.Demand.TryGetValue(key, out var demand))
            {
                demand = new MaterialDemand
                {
                    Unit = unit,
                    Levels = new Dictionary<int, Dictionary<string, double>>(),
                    MountedLevels = new Dictionary<int, Dictionary<string, double>>(),
                    Floors = new Dictionary<int, Dictionary<int, double>>(),
                    MountedFloors = new Dictionary<int, Dictionary<int, double>>()
                };
                currentObject.Demand[key] = demand;
            }

            demand.Levels ??= new Dictionary<int, Dictionary<string, double>>();
            demand.MountedLevels ??= new Dictionary<int, Dictionary<string, double>>();
            demand.Floors ??= new Dictionary<int, Dictionary<int, double>>();
            demand.MountedFloors ??= new Dictionary<int, Dictionary<int, double>>();
            if (string.IsNullOrWhiteSpace(demand.Unit))
                demand.Unit = unit;

            return demand;
        }

        private static double GetValue(Dictionary<int, Dictionary<string, double>> map, int block, string mark)
        {
            if (map != null && map.TryGetValue(block, out var levels) && levels.TryGetValue(mark, out var value))
                return value;
            return 0;
        }

        private static string Format(double value)
            => Math.Abs(value % 1) < 0.0001 ? value.ToString("0", CultureInfo.CurrentCulture) : value.ToString("0.##", CultureInfo.CurrentCulture);

        private static double Parse(string text)
        {
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out var value))
                return value;
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                return value;
            return 0;
        }

        private void AddCell(Grid grid, int row, int col, string text, bool bold, string bg, bool header, HorizontalAlignment alignment = HorizontalAlignment.Left)
        {
            var border = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225)),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Background = (Brush)new BrushConverter().ConvertFromString(bg),
                Padding = new Thickness(6, 5, 6, 5)
            };
            border.Child = new TextBlock
            {
                Text = text,
                FontWeight = bold ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = header ? new SolidColorBrush(Color.FromRgb(30, 41, 59)) : new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                HorizontalAlignment = alignment,
                TextAlignment = alignment == HorizontalAlignment.Center ? TextAlignment.Center : TextAlignment.Left
            };
           
            Grid.SetRow(border, row);
            Grid.SetColumn(border, col);
            grid.Children.Add(border);
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            foreach (var tb in editorsByGroup.Values.SelectMany(x => x))
            {
                var (key, unit, block, mark) = ((string, string, int, string))tb.Tag;
                var demand = GetOrCreateDemand(key, unit);
                demand.Levels.TryAdd(block, new Dictionary<string, double>());
                demand.Levels[block][mark] = Parse(tb.Text);
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
