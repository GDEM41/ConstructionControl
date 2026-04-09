using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;

namespace ConstructionControl
{
    public partial class ObjectSettingsWindow : Window
    {
        private readonly ProjectObject _object;
        private readonly ObservableCollection<BlockFloors> floorsRows = new();
        private readonly ObservableCollection<BlockAxisRow> blockAxisRows = new();
        private readonly ObservableCollection<NamedPersonRow> masterRows = new();
        private readonly ObservableCollection<NamedPersonRow> foremanRows = new();

        public ObjectSettingsWindow(ProjectObject obj)
        {
            InitializeComponent();
            _object = obj;

            NameBox.Text = obj.Name ?? string.Empty;
            FullObjectNameBox.Text = obj.FullObjectName ?? string.Empty;
            BlocksBox.Text = Math.Max(1, obj.BlocksCount).ToString();
            SameFloorsCheck.IsChecked = obj.SameFloorsInBlocks;
            FloorsPerBlockBox.Text = Math.Max(1, obj.FloorsPerBlock).ToString();
            BasementCheck.IsChecked = obj.HasBasement;

            GeneralContractorRepresentativeBox.Text = obj.GeneralContractorRepresentative ?? string.Empty;
            TechnicalSupervisorRepresentativeBox.Text = obj.TechnicalSupervisorRepresentative ?? string.Empty;
            ProjectOrganizationRepresentativeBox.Text = obj.ProjectOrganizationRepresentative ?? string.Empty;
            ProjectDocumentationNameBox.Text = obj.ProjectDocumentationName ?? string.Empty;
            SiteManagerNameBox.Text = obj.SiteManagerName ?? string.Empty;

            for (var i = 0; i <= 12; i++)
            {
                MasterCountComboBox.Items.Add(i);
                ForemanCountComboBox.Items.Add(i);
            }

            BuildFloorsGrid();
            BuildBlockAxesGrid();
            LoadMasters();
            LoadForemen();
            UpdateVisibility();
        }

        private void SameFloorsChanged(object sender, RoutedEventArgs e)
        {
            UpdateVisibility();
        }

        private void UpdateVisibility()
        {
            var same = SameFloorsCheck.IsChecked == true;
            SameFloorsPanel.Visibility = same ? Visibility.Visible : Visibility.Collapsed;
            FloorsByBlockGrid.Visibility = same ? Visibility.Collapsed : Visibility.Visible;
        }

        private int ParseBlocksCount()
        {
            if (!int.TryParse(BlocksBox.Text, out var blocks))
                blocks = 1;
            return Math.Max(1, blocks);
        }

        private void BuildFloorsGrid()
        {
            var blocks = ParseBlocksCount();
            floorsRows.Clear();

            for (var i = 1; i <= blocks; i++)
            {
                floorsRows.Add(new BlockFloors
                {
                    Block = i,
                    Floors = _object.FloorsByBlock.ContainsKey(i)
                        ? _object.FloorsByBlock[i]
                        : Math.Max(1, _object.FloorsPerBlock)
                });
            }

            FloorsByBlockGrid.ItemsSource = floorsRows;
        }

        private void BuildBlockAxesGrid()
        {
            var blocks = ParseBlocksCount();
            _object.BlockAxesByNumber ??= new Dictionary<int, string>();
            blockAxisRows.Clear();

            for (var i = 1; i <= blocks; i++)
            {
                blockAxisRows.Add(new BlockAxisRow
                {
                    Block = i,
                    Axes = _object.BlockAxesByNumber.TryGetValue(i, out var axes)
                        ? axes ?? string.Empty
                        : string.Empty
                });
            }

            BlockAxesGrid.ItemsSource = blockAxisRows;
        }

        private void LoadMasters()
        {
            var names = (_object.MasterNames ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .ToList();

            if (names.Count == 0)
                names.Add(string.Empty);

            masterRows.Clear();
            for (var i = 0; i < names.Count; i++)
            {
                masterRows.Add(new NamedPersonRow
                {
                    Index = i + 1,
                    FullName = names[i]
                });
            }

            MastersGrid.ItemsSource = masterRows;
            MasterCountComboBox.SelectedItem = masterRows.Count;
        }

        private void LoadForemen()
        {
            var names = (_object.ForemanNames ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .ToList();

            if (names.Count == 0)
                names.Add(string.Empty);

            foremanRows.Clear();
            for (var i = 0; i < names.Count; i++)
            {
                var row = new NamedPersonRow
                {
                    Index = i + 1,
                    FullName = names[i]
                };
                row.PropertyChanged += ForemanRow_PropertyChanged;
                foremanRows.Add(row);
            }

            ForemenGrid.ItemsSource = foremanRows;
            ForemanCountComboBox.SelectedItem = foremanRows.Count;
            RefreshResponsibleForemanList();

            if (!string.IsNullOrWhiteSpace(_object.ResponsibleForeman))
            {
                var selected = foremanRows.FirstOrDefault(x =>
                    string.Equals(x.FullName?.Trim(), _object.ResponsibleForeman.Trim(), StringComparison.CurrentCultureIgnoreCase));
                if (selected != null)
                    ResponsibleForemanComboBox.SelectedItem = selected;
            }
        }

        private void ForemanRow_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(NamedPersonRow.FullName))
                RefreshResponsibleForemanList();
        }

        private void RefreshResponsibleForemanList()
        {
            var selectedName = (ResponsibleForemanComboBox.SelectedItem as NamedPersonRow)?.FullName?.Trim();
            var list = foremanRows.Where(x => !string.IsNullOrWhiteSpace(x.FullName)).ToList();
            ResponsibleForemanComboBox.ItemsSource = null;
            ResponsibleForemanComboBox.ItemsSource = list;

            if (!string.IsNullOrWhiteSpace(selectedName))
            {
                var selected = list.FirstOrDefault(x => string.Equals(x.FullName?.Trim(), selectedName, StringComparison.CurrentCultureIgnoreCase));
                if (selected != null)
                    ResponsibleForemanComboBox.SelectedItem = selected;
            }
        }

        private static void ResizeNamedRows(ObservableCollection<NamedPersonRow> rows, int targetCount)
        {
            var count = Math.Max(0, targetCount);
            while (rows.Count < count)
            {
                rows.Add(new NamedPersonRow
                {
                    Index = rows.Count + 1,
                    FullName = string.Empty
                });
            }

            while (rows.Count > count)
                rows.RemoveAt(rows.Count - 1);

            for (var i = 0; i < rows.Count; i++)
                rows[i].Index = i + 1;
        }

        private void MasterCountComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MasterCountComboBox.SelectedItem is int count)
                ResizeNamedRows(masterRows, count);
        }

        private void ForemanCountComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ForemanCountComboBox.SelectedItem is not int count)
                return;

            foreach (var row in foremanRows)
                row.PropertyChanged -= ForemanRow_PropertyChanged;

            ResizeNamedRows(foremanRows, count);

            foreach (var row in foremanRows)
                row.PropertyChanged += ForemanRow_PropertyChanged;

            RefreshResponsibleForemanList();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            _object.Name = NameBox.Text.Trim();
            _object.FullObjectName = FullObjectNameBox.Text?.Trim() ?? string.Empty;
            _object.BlocksCount = ParseBlocksCount();
            _object.HasBasement = BasementCheck.IsChecked == true;
            _object.SameFloorsInBlocks = SameFloorsCheck.IsChecked == true;

            _object.GeneralContractorRepresentative = GeneralContractorRepresentativeBox.Text?.Trim() ?? string.Empty;
            _object.TechnicalSupervisorRepresentative = TechnicalSupervisorRepresentativeBox.Text?.Trim() ?? string.Empty;
            _object.ProjectOrganizationRepresentative = ProjectOrganizationRepresentativeBox.Text?.Trim() ?? string.Empty;
            _object.ProjectDocumentationName = ProjectDocumentationNameBox.Text?.Trim() ?? string.Empty;
            _object.SiteManagerName = SiteManagerNameBox.Text?.Trim() ?? string.Empty;
            _object.MasterNames = masterRows.Select(x => x.FullName?.Trim()).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            _object.ForemanNames = foremanRows.Select(x => x.FullName?.Trim()).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            _object.ResponsibleForeman = (ResponsibleForemanComboBox.SelectedItem as NamedPersonRow)?.FullName?.Trim() ?? string.Empty;
            _object.BlockAxesByNumber = blockAxisRows.ToDictionary(x => x.Block, x => x.Axes?.Trim() ?? string.Empty);

            if (_object.SameFloorsInBlocks)
            {
                _object.FloorsPerBlock = Math.Max(1, int.TryParse(FloorsPerBlockBox.Text, out var floors) ? floors : 1);
            }
            else
            {
                _object.FloorsByBlock.Clear();
                foreach (var row in floorsRows)
                    _object.FloorsByBlock[row.Block] = Math.Max(0, row.Floors);
            }

            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void BlocksBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            BuildFloorsGrid();
            BuildBlockAxesGrid();
        }

        private sealed class BlockFloors
        {
            public int Block { get; set; }
            public int Floors { get; set; }
        }

        private sealed class BlockAxisRow
        {
            public int Block { get; set; }
            public string Axes { get; set; } = string.Empty;
        }

        private sealed class NamedPersonRow : INotifyPropertyChanged
        {
            private string fullName = string.Empty;

            public int Index { get; set; }

            public string FullName
            {
                get => fullName;
                set => SetField(ref fullName, value ?? string.Empty);
            }

            public event PropertyChangedEventHandler PropertyChanged;

            private bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
            {
                if (EqualityComparer<T>.Default.Equals(field, value))
                    return false;

                field = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                return true;
            }
        }
    }
}
