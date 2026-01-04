using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace ConstructionControl
{
    public partial class ObjectSettingsWindow : Window
    {
        private ProjectObject _object;

        public ObjectSettingsWindow(ProjectObject obj)
        {
            InitializeComponent();
            _object = obj;

            // Заполнение полей
            NameBox.Text = obj.Name;
            BlocksBox.Text = obj.BlocksCount.ToString();
            SameFloorsCheck.IsChecked = obj.SameFloorsInBlocks;
            FloorsPerBlockBox.Text = obj.FloorsPerBlock.ToString();
            BasementCheck.IsChecked = obj.HasBasement;

            BuildFloorsGrid();
            UpdateVisibility();
        }

        private void SameFloorsChanged(object sender, RoutedEventArgs e)
        {
            UpdateVisibility();
        }

        private void UpdateVisibility()
        {
            bool same = SameFloorsCheck.IsChecked == true;
            SameFloorsPanel.Visibility = same ? Visibility.Visible : Visibility.Collapsed;
            FloorsByBlockGrid.Visibility = same ? Visibility.Collapsed : Visibility.Visible;
        }

        private void BuildFloorsGrid()
        {
            int blocks = _object.BlocksCount;
            var list = new List<BlockFloors>();

            for (int i = 1; i <= blocks; i++)
            {
                list.Add(new BlockFloors
                {
                    Block = i,
                    Floors = _object.FloorsByBlock.ContainsKey(i)
                        ? _object.FloorsByBlock[i]
                        : 0
                });
            }

            FloorsByBlockGrid.ItemsSource = list;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            _object.Name = NameBox.Text.Trim();
            _object.BlocksCount = int.TryParse(BlocksBox.Text, out var b) ? b : 0;
            _object.HasBasement = BasementCheck.IsChecked == true;
            _object.SameFloorsInBlocks = SameFloorsCheck.IsChecked == true;

            if (_object.SameFloorsInBlocks)
            {
                _object.FloorsPerBlock =
                    int.TryParse(FloorsPerBlockBox.Text, out var f) ? f : 0;
            }
            else
            {
                _object.FloorsByBlock.Clear();
                foreach (BlockFloors row in FloorsByBlockGrid.ItemsSource)
                    _object.FloorsByBlock[row.Block] = row.Floors;
            }

            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private class BlockFloors
        {
            public int Block { get; set; }
            public int Floors { get; set; }
        }
        private void BlocksBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (!int.TryParse(BlocksBox.Text, out var blocks) || blocks < 0)
                return;

            _object.BlocksCount = blocks;

            BuildFloorsGrid();
        }

    }
}
