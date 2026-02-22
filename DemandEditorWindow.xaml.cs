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
        private readonly List<(int Block, int Floor)> floorColumns = new();
        private readonly Dictionary<string, List<TextBox>> editorsByGroup = new();

        public DemandEditorWindow(ProjectObject projectObject, IEnumerable<DemandMaterialRow> sourceRows)
        {
            InitializeComponent();
            currentObject = projectObject;
            rows = sourceRows.OrderBy(x => x.Group).ThenBy(x => x.Material).ToList();
            BuildColumns();
            RenderTabs();
        }

        private void BuildColumns()
        {
            floorColumns.Clear();
            for (int b = 1; b <= currentObject.BlocksCount; b++)
            {
                int floors = currentObject.SameFloorsInBlocks
                    ? currentObject.FloorsPerBlock
                    : (currentObject.FloorsByBlock.TryGetValue(b, out var value) ? value : 0);

                if (currentObject.HasBasement)
                    floorColumns.Add((b, 0));

                for (int f = 1; f <= floors; f++)
                    floorColumns.Add((b, f));
            }
        }

        private void RenderTabs()
        {
            TypeTabs.Items.Clear();
            editorsByGroup.Clear();

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
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(280) });
            foreach (var _ in floorColumns)
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 64 });

            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            AddCell(grid, 0, 0, "Материал", true, "#E2E8F0", true);
            for (int i = 0; i < floorColumns.Count; i++)
            {
                var col = floorColumns[i];
                AddCell(grid, 0, i + 1, $"Б{col.Block}-{(col.Floor == 0 ? "П" : col.Floor.ToString())}", true, "#E2E8F0", true);
            }

            var editors = new List<TextBox>();
            for (int rowIndex = 0; rowIndex < groupRows.Count; rowIndex++)
            {
                var row = groupRows[rowIndex];
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var bg = rowIndex % 2 == 0 ? "White" : "#F8FAFC";
                AddCell(grid, rowIndex + 1, 0, row.Material, false, bg, false);

                string key = $"{row.Group}::{row.Material}";
                var demand = GetOrCreateDemand(key, row.Unit);
                for (int i = 0; i < floorColumns.Count; i++)
                {
                    var col = floorColumns[i];
                    var value = GetValue(demand.Floors, col.Block, col.Floor);
                    var box = new TextBox
                    {
                        Text = Format(value),
                        MinWidth = 58,
                        Margin = new Thickness(2),
                        Padding = new Thickness(4, 2, 4, 2),
                        HorizontalContentAlignment = HorizontalAlignment.Right,
                        Tag = (key, row.Unit, col.Block, col.Floor),
                        BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1")),
                        BorderThickness = new Thickness(1)
                    };
                    box.PreviewKeyDown += Cell_PreviewKeyDown;
                    Grid.SetRow(box, rowIndex + 1);
                    Grid.SetColumn(box, i + 1);
                    grid.Children.Add(box);
                    editors.Add(box);
                }
            }

            editorsByGroup[group] = editors;
            return grid;
        }
        private void Cell_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (sender is not TextBox current)
                return;

            int index = -1;
            List<TextBox> list = null;
            foreach (var kv in editorsByGroup)
            {
                index = kv.Value.IndexOf(current);
                if (index >= 0)
                {
                    list = kv.Value;
                    break;
                }
            }

            if (list == null)
                return;

            int target = index;
            if (e.Key == Key.Right || e.Key == Key.Enter || e.Key == Key.Down)
                target = Math.Min(list.Count - 1, index + 1);
            else if (e.Key == Key.Left || e.Key == Key.Up)
                target = Math.Max(0, index - 1);

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
                    Floors = new Dictionary<int, Dictionary<int, double>>(),
                    MountedFloors = new Dictionary<int, Dictionary<int, double>>()
                };
                currentObject.Demand[key] = demand;
            }

            demand.Floors ??= new Dictionary<int, Dictionary<int, double>>();
            demand.MountedFloors ??= new Dictionary<int, Dictionary<int, double>>();
            if (string.IsNullOrWhiteSpace(demand.Unit))
                demand.Unit = unit;

            return demand;
        }

        private static double GetValue(Dictionary<int, Dictionary<int, double>> map, int block, int floor)
        {
            if (map != null && map.TryGetValue(block, out var floors) && floors.TryGetValue(floor, out var value))
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

        private void AddCell(Grid grid, int row, int col, string text, bool bold, string bg, bool header)
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
                Foreground = header ? new SolidColorBrush(Color.FromRgb(30, 41, 59)) : new SolidColorBrush(Color.FromRgb(51, 65, 85))
            };
           
            Grid.SetRow(border, row);
            Grid.SetColumn(border, col);
            grid.Children.Add(border);
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            foreach (var tb in editorsByGroup.Values.SelectMany(x => x))
            {
                var (key, unit, block, floor) = ((string, string, int, int))tb.Tag;
                var demand = GetOrCreateDemand(key, unit);
                demand.Floors.TryAdd(block, new Dictionary<int, double>());
                demand.Floors[block][floor] = Parse(tb.Text);
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