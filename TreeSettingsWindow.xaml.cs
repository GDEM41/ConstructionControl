using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;

namespace ConstructionControl
{
    public partial class TreeSettingsWindow : Window
    {
        public class MaterialSplitRuleSource
        {
            public string Category { get; set; }
            public string TypeName { get; set; }
            public string MaterialName { get; set; }
        }

        public class MaterialSplitRuleRow : INotifyPropertyChanged
        {
            private string splitPath;

            public string Category { get; set; }
            public string TypeName { get; set; }
            public string MaterialName { get; set; }
            public string SplitPath
            {
                get => splitPath;
                set
                {
                    if (splitPath == value)
                        return;

                    splitPath = value;
                    OnPropertyChanged();
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            private void OnPropertyChanged([CallerMemberName] string propertyName = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private readonly ObservableCollection<MaterialSplitRuleRow> rows;
        private bool isBulkUpdating;
        public Dictionary<string, string> ResultRules { get; private set; } = new();

        public TreeSettingsWindow(IEnumerable<MaterialSplitRuleSource> materials, Dictionary<string, string> existingRules)
        {
            InitializeComponent();

            rows = new ObservableCollection<MaterialSplitRuleRow>(
                                materials
                    .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName))
                    .GroupBy(x => x.MaterialName)
                    .Select(g => g.First())
                    .OrderBy(x => x.Category)
                    .ThenBy(x => x.TypeName)
                    .ThenBy(x => x.MaterialName)
                    .Select(x => new MaterialSplitRuleRow
                    {
                        Category = x.Category,
                        TypeName = x.TypeName,
                        MaterialName = x.MaterialName,
                        SplitPath = existingRules != null && existingRules.TryGetValue(x.MaterialName, out var rule)
                            ? rule
                            : string.Empty
                    }));

            RulesGrid.ItemsSource = rows;
        }
        private void RulesGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (isBulkUpdating || e.EditAction != DataGridEditAction.Commit)
                return;

            if (e.Row?.Item is not MaterialSplitRuleRow edited)
                return;

            var raw = edited.SplitPath?.Trim();
            if (string.IsNullOrWhiteSpace(raw))
                return;

            var editedSegments = raw
                .Split('|', System.StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            if (editedSegments.Count == 0)
                return;

            var typeSegment = editedSegments[0];

            isBulkUpdating = true;
            try
            {
                foreach (var row in rows.Where(x => x.Category == edited.Category && x.TypeName == edited.TypeName))
                {
                    var materialParts = MainWindow.GetSegmentsFromText(row.MaterialName);
                    row.SplitPath = materialParts.Count <= 1
                        ? typeSegment
                        : string.Join("|", new[] { typeSegment }.Concat(materialParts.Skip(1)));
                }
            }
            finally
            {
                isBulkUpdating = false;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            ResultRules = rows
                .Where(x => !string.IsNullOrWhiteSpace(x.SplitPath))
                .ToDictionary(
                    x => x.MaterialName,
                    x => x.SplitPath.Trim());

            DialogResult = true;
            Close();
        }
    }
}