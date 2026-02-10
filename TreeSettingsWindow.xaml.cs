using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace ConstructionControl
{
    public partial class TreeSettingsWindow : Window
    {
        public class MaterialSplitRuleRow
        {
            public string MaterialName { get; set; }
            public string SplitPath { get; set; }
        }

        private readonly ObservableCollection<MaterialSplitRuleRow> rows;
        public Dictionary<string, string> ResultRules { get; private set; } = new();

        public TreeSettingsWindow(IEnumerable<string> materialNames, Dictionary<string, string> existingRules)
        {
            InitializeComponent();

            rows = new ObservableCollection<MaterialSplitRuleRow>(
                materialNames
                    .Distinct()
                    .OrderBy(x => x)
                    .Select(m => new MaterialSplitRuleRow
                    {
                        MaterialName = m,
                        SplitPath = existingRules != null && existingRules.TryGetValue(m, out var rule)
                            ? rule
                            : string.Empty
                    }));

            RulesGrid.ItemsSource = rows;
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