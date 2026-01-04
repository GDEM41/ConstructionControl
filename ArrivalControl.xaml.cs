using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace ConstructionControl
{
    public partial class ArrivalControl : UserControl
    {
        private ProjectObject currentObject;
        private ObservableCollection<ArrivalItem> items = new();

        public event System.Action<Arrival> ArrivalAdded;

        public ArrivalControl()
        {
            InitializeComponent();
            ItemsGrid.ItemsSource = items;
        }

        // ================= ХАК ДЛЯ ComboBox.Text =================

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            MaterialGroupBox.AddHandler(
                TextBoxBase.TextChangedEvent,
                new TextChangedEventHandler(MaterialGroupTextChanged));
        }

        // ================= ИНИЦИАЛИЗАЦИЯ =================

        public void SetObject(ProjectObject obj)
        {
            currentObject = obj;

            MaterialGroupBox.ItemsSource =
                currentObject.MaterialGroups.Select(g => g.Name).ToList();

            items.Clear();
            AddRow();
        }

        // ================= ТИП МАТЕРИАЛА =================

        private void MaterialGroupTextChanged(object sender, TextChangedEventArgs e)
        {
            var group = MaterialGroupBox.Text?.Trim();

            foreach (var item in items)
            {
                item.MaterialName = null;
                item.AvailableNames.Clear();

                if (string.IsNullOrWhiteSpace(group))
                    continue;

                if (currentObject.MaterialNamesByGroup.TryGetValue(group, out var names))
                {
                    foreach (var n in names)
                        item.AvailableNames.Add(n);
                }
            }
        }

        // ================= СТРОКИ =================

        private void AddRow()
        {
            var item = new ArrivalItem
            {
                Date = System.DateTime.Today,
                AvailableNames = new ObservableCollection<string>(),
                AvailableUnits = new ObservableCollection<string>
                {
                    "шт", "м3", "т", "кг"
                }
            };

            items.Add(item);
        }

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            AddRow();
        }

        // ================= ДОБАВЛЕНИЕ ПРИХОДА =================

        private void AddArrival_Click(object sender, RoutedEventArgs e)
        {
            var groupName = MaterialGroupBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(groupName))
            {
                MessageBox.Show("Укажите тип материала");
                return;
            }

            if (!currentObject.MaterialGroups.Any(g => g.Name == groupName))
                currentObject.MaterialGroups.Add(new MaterialGroup { Name = groupName });

            if (!currentObject.MaterialNamesByGroup.ContainsKey(groupName))
                currentObject.MaterialNamesByGroup[groupName] = new();

            foreach (var i in items)
            {
                if (!string.IsNullOrWhiteSpace(i.MaterialName) &&
                    !currentObject.MaterialNamesByGroup[groupName].Contains(i.MaterialName))
                {
                    currentObject.MaterialNamesByGroup[groupName].Add(i.MaterialName);
                }
            }

            ArrivalAdded?.Invoke(new Arrival
            {
                MaterialGroup = groupName,
                TtnNumber = TtnBox.Text,
                Items = items.ToList()
            });

            TtnBox.Clear();
            items.Clear();
            AddRow();
        }
    }
}
