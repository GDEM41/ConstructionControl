using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
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

        public DemandEditorWindow(ProjectObject projectObject, IEnumerable<DemandMaterialRow> sourceRows)
        {
            InitializeComponent();
            currentObject = projectObject;
            rows = sourceRows.OrderBy(x => x.Group).ThenBy(x => x.Material).ToList();
            BuildColumns();
            RenderTable();
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

        private void RenderTable()
        {
            DemandGrid.Children.Clear();
            DemandGrid.RowDefinitions.Clear();
            DemandGrid.ColumnDefinitions.Clear();

            DemandGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(280) });
            foreach (var _ in floorColumns)
                DemandGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 60 });

            DemandGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            AddCell(0, 0, "Материал", true, "#F1F5F9");
            for (int i = 0; i < floorColumns.Count; i++)
            {
                var col = floorColumns[i];
                var caption = $"Б{col.Block}-{(col.Floor == 0 ? "П" : col.Floor.ToString())}";
                AddCell(0, i + 1, caption, true, "#F1F5F9");
            }

            int rowIndex = 1;
            foreach (var row in rows)
            {
                DemandGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                AddCell(rowIndex, 0, $"{row.Group} / {row.Material}", false, "White");

                string key = $"{row.Group}::{row.Material}";
                var demand = GetOrCreateDemand(key, row.Unit);
                for (int i = 0; i < floorColumns.Count; i++)
                {
                    var col = floorColumns[i];
                    var value = GetValue(demand.Floors, col.Block, col.Floor);
                    var box = new TextBox
                    {
                        Text = Format(value),
                        MinWidth = 52,
                        Margin = new Thickness(2),
                        HorizontalContentAlignment = HorizontalAlignment.Right,
                        Tag = (key, row.Unit, col.Block, col.Floor)
                    };
                    Grid.SetRow(box, rowIndex);
                    Grid.SetColumn(box, i + 1);
                    DemandGrid.Children.Add(box);
                }

                rowIndex++;
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

        private void AddCell(int row, int col, string text, bool bold, string bg)
        {
            var border = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Background = (Brush)new BrushConverter().ConvertFromString(bg),
                Padding = new Thickness(6, 4, 6, 4)
            };
            border.Child = new TextBlock { Text = text, FontWeight = bold ? FontWeights.SemiBold : FontWeights.Normal };
            Grid.SetRow(border, row);
            Grid.SetColumn(border, col);
            DemandGrid.Children.Add(border);
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            foreach (var tb in DemandGrid.Children.OfType<TextBox>())
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