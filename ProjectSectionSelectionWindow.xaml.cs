using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ConstructionControl
{
    public partial class ProjectSectionSelectionWindow : Window
    {
        private readonly Dictionary<ProjectTransferSection, CheckBox> sectionCheckBoxes = new();
        private readonly HashSet<ProjectTransferSection> enabledSections;

        private const ProjectTransferSection DataSections =
            ProjectTransferSection.ObjectSettings
            | ProjectTransferSection.MaterialsAndSummary
            | ProjectTransferSection.Arrival
            | ProjectTransferSection.Ot
            | ProjectTransferSection.Timesheet
            | ProjectTransferSection.Production
            | ProjectTransferSection.HiddenWorkActs
            | ProjectTransferSection.Inspection
            | ProjectTransferSection.Notes;

        public ProjectTransferSection SelectedSections { get; private set; }

        public ProjectSectionSelectionWindow(
            string title,
            string subtitle,
            string hint,
            string confirmText,
            IReadOnlyList<ProjectSectionSelectionOption> options)
        {
            InitializeComponent();

            DialogTitleText.Text = title ?? string.Empty;
            DialogSubtitleText.Text = subtitle ?? string.Empty;
            DialogHintText.Text = hint ?? string.Empty;
            ConfirmButton.Content = string.IsNullOrWhiteSpace(confirmText) ? "Применить" : confirmText;

            enabledSections = options?
                .Where(x => x != null && x.IsEnabled)
                .Select(x => x.Section)
                .ToHashSet()
                ?? new HashSet<ProjectTransferSection>();

            BuildSectionRows(options ?? Array.Empty<ProjectSectionSelectionOption>());
            UpdateButtons();
        }

        private void BuildSectionRows(IEnumerable<ProjectSectionSelectionOption> options)
        {
            foreach (var option in options.Where(x => x != null))
            {
                var card = new Border
                {
                    Margin = new Thickness(0, 0, 0, 10),
                    Padding = new Thickness(14, 12, 14, 12),
                    Background = FindResource("SurfaceBrush") as Brush ?? Brushes.White,
                    BorderBrush = FindResource("StrokeBrush") as Brush ?? Brushes.Gainsboro,
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(12),
                    Opacity = option.IsEnabled ? 1.0 : 0.58
                };

                var stack = new StackPanel();
                var checkBox = new CheckBox
                {
                    Content = option.Title ?? string.Empty,
                    IsChecked = option.IsSelected && option.IsEnabled,
                    IsEnabled = option.IsEnabled,
                    FontSize = 15,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = FindResource("TextBrush") as Brush ?? Brushes.Black
                };
                checkBox.Checked += SectionCheckBox_Changed;
                checkBox.Unchecked += SectionCheckBox_Changed;

                var description = new TextBlock
                {
                    Margin = new Thickness(28, 6, 0, 0),
                    Text = option.Description ?? string.Empty,
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = FindResource("TextSecondaryBrush") as Brush ?? Brushes.DimGray
                };

                stack.Children.Add(checkBox);
                stack.Children.Add(description);
                card.Child = stack;
                SectionsHost.Children.Add(card);
                sectionCheckBoxes[option.Section] = checkBox;
            }
        }

        private void SectionCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            UpdateButtons();
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            SetSections(enabledSections, true);
        }

        private void SelectNoneButton_Click(object sender, RoutedEventArgs e)
        {
            SetSections(enabledSections, false);
        }

        private void DataOnlyButton_Click(object sender, RoutedEventArgs e)
        {
            SetSections(enabledSections, false);
            SetSections(enabledSections.Where(x => (x & DataSections) == x), true);
        }

        private void DataWithPdfButton_Click(object sender, RoutedEventArgs e)
        {
            DataOnlyButton_Click(sender, e);
            SetSections(new[] { ProjectTransferSection.Pdf }, true);
        }

        private void DataWithPdfEstimateButton_Click(object sender, RoutedEventArgs e)
        {
            DataOnlyButton_Click(sender, e);
            SetSections(new[] { ProjectTransferSection.Pdf, ProjectTransferSection.Estimates }, true);
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = ProjectTransferSection.None;
            foreach (var pair in sectionCheckBoxes)
            {
                if (pair.Value.IsChecked == true)
                    selected |= pair.Key;
            }

            if (selected == ProjectTransferSection.None)
            {
                MessageBox.Show(
                    "Выберите хотя бы один раздел.",
                    "Разделы проекта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            SelectedSections = selected;
            DialogResult = true;
        }

        private void SetSections(IEnumerable<ProjectTransferSection> sections, bool value)
        {
            foreach (var section in sections.Distinct())
            {
                if (sectionCheckBoxes.TryGetValue(section, out var checkBox) && checkBox.IsEnabled)
                    checkBox.IsChecked = value;
            }

            UpdateButtons();
        }

        private void UpdateButtons()
        {
            var hasEnabled = sectionCheckBoxes.Values.Any(x => x.IsEnabled);
            var anySelected = sectionCheckBoxes.Values.Any(x => x.IsEnabled && x.IsChecked == true);
            ConfirmButton.IsEnabled = anySelected;
            SelectAllButton.IsEnabled = hasEnabled;
            SelectNoneButton.IsEnabled = hasEnabled;
            DataOnlyButton.IsEnabled = hasEnabled;
            DataWithPdfButton.IsEnabled = hasEnabled;
            DataWithPdfEstimateButton.IsEnabled = hasEnabled;
        }
    }
}
