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
            public string CategoryName { get; set; }
            public string TypeName { get; set; }
            public string SubTypeName { get; set; }
            public string MaterialName { get; set; }
        }

        public class MaterialSplitRuleRow : INotifyPropertyChanged
        {
            private readonly string[] segments = new string[6];
            public string CategoryName { get; set; }
            public string TypeName { get; set; }
            public string SubTypeName { get; set; }
            public string MaterialName { get; set; }

            public string Segment1 { get => segments[0]; set => SetSegment(0, value); }
            public string Segment2 { get => segments[1]; set => SetSegment(1, value); }
            public string Segment3 { get => segments[2]; set => SetSegment(2, value); }
            public string Segment4 { get => segments[3]; set => SetSegment(3, value); }
            public string Segment5 { get => segments[4]; set => SetSegment(4, value); }
            public string Segment6 { get => segments[5]; set => SetSegment(5, value); }

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
            }
            public string GetRule() => NormalizeRule(string.Join("|", segments));

            private void SetSegment(int index, string value)
            {
                var normalized = (value ?? string.Empty).Trim();
                if (string.Equals(segments[index], normalized, System.StringComparison.CurrentCulture))
                    return;

                segments[index] = normalized;
                OnPropertyChanged($"Segment{index + 1}");
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
                    .OrderBy(x => x.CategoryName)
                    .ThenBy(x => x.TypeName)
                    .ThenBy(x => x.SubTypeName)
                    .ThenBy(x => x.MaterialName)
                    .Select(x => new MaterialSplitRuleRow
                    {
                        CategoryName = x.CategoryName,
                        TypeName = x.TypeName,
                        SubTypeName = x.SubTypeName,
                        MaterialName = x.MaterialName
                    }));
            foreach (var row in rows)
            {
                row.SetRule(existingRules != null && existingRules.TryGetValue(row.MaterialName, out var rule)
                    ? rule
                    : string.Empty);
            }

            RulesGrid.ItemsSource = rows;
        }
        private void RulesGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (isBulkUpdating || e.EditAction != DataGridEditAction.Commit)
                return;

            if (e.Row?.Item is not MaterialSplitRuleRow edited)
                return;

            var normalizedRule = edited.GetRule();
            if (string.IsNullOrWhiteSpace(normalizedRule))
                return;
            edited.SetRule(normalizedRule);

            var targets = PromptRuleTargets(edited);
            if (targets == null || targets.Count == 0)
                return;

            var sourcePattern = BuildRulePattern(edited.MaterialName, normalizedRule);



            isBulkUpdating = true;
            try
            {
                foreach (var target in targets)
                {
                    if (ReferenceEquals(target, edited))
                        continue;

                    var convertedRule = ApplyRuleByPattern(target.MaterialName, sourcePattern);
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
                            && string.Equals(x.CategoryName, edited.CategoryName, System.StringComparison.CurrentCultureIgnoreCase)
                            && string.Equals(x.TypeName, edited.TypeName, System.StringComparison.CurrentCultureIgnoreCase)
                            && string.Equals(x.SubTypeName, edited.SubTypeName, System.StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            if (candidates.Count == 0)
                return new List<MaterialSplitRuleRow>();

            var panel = new StackPanel();

            panel.Children.Add(new TextBlock
            {
                Text = $"Применить разбиение к другим элементам ({edited.CategoryName} / {edited.TypeName} / {edited.SubTypeName})?",
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
            var materialTokens = MainWindow.GetSegmentsFromText(sourceMaterialName);
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
                var segmentTokens = MainWindow.GetSegmentsFromText(segment);
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

            var targetTokens = MainWindow.GetSegmentsFromText(targetMaterialName);
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
            ResultRules = rows
                .Select(x => new { x.MaterialName, Rule = x.GetRule() })
                .Where(x => !string.IsNullOrWhiteSpace(x.Rule))
                .ToDictionary(x => x.MaterialName, x => x.Rule);


            DialogResult = true;
            Close();
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
    }
}