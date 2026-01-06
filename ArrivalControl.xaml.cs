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
        private List<JournalRecord> journal;
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

        private void ArrivalTypeChanged(object sender, RoutedEventArgs e)
        {
            // защита от вызова во время InitializeComponent
            if (ExtraTypeBox == null || MaterialGroupBox == null)
                return;

            if (ExtraRadio.IsChecked == true)
            {
                ExtraTypeBox.Visibility = Visibility.Visible;
                ExtraTypeBox.ItemsSource = new[] { "Внутренние", "Малоценка" };
                ExtraTypeBox.SelectedIndex = 0;

                MaterialGroupBox.Visibility = Visibility.Collapsed;
            }
            else
            {
                ExtraTypeBox.Visibility = Visibility.Collapsed;
                MaterialGroupBox.Visibility = Visibility.Visible;
            }
        }



        public void SetObject(ProjectObject obj, List<JournalRecord> journalRecords)
        {
            currentObject = obj;
            journal = journalRecords;

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
                AvailableNames = new ObservableCollection<string>()
            };

            items.Add(item);

            item.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(ArrivalItem.MaterialName))
                {
                    TryAutofillUnitAndStb(item);
                }
            };
        }



        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            AddRow();
        }

        // ================= ДОБАВЛЕНИЕ ПРИХОДА =================

        private void AddArrival_Click(object sender, RoutedEventArgs e)
        {
            // ===== ОБЯЗАТЕЛЬНАЯ ЗАЩИТА =====
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

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

                // ===== ШАГ 6: ЗАПОМИНАЕМ ЕД. ИЗМ И СТБ =====
                if (!string.IsNullOrWhiteSpace(i.MaterialName))
                {
                  
                }
            }


            ArrivalAdded?.Invoke(new Arrival
            {
                Category = MainRadio.IsChecked == true ? "Основные" : "Допы",
                SubCategory = ExtraRadio.IsChecked == true
                    ? ExtraTypeBox.SelectedItem?.ToString()
                    : null,

                MaterialGroup = MainRadio.IsChecked == true ? groupName : null,
                TtnNumber = TtnBox.Text,
                Items = items.ToList()
            });


            TtnBox.Clear();
            items.Clear();
            AddRow();
        }
        private void TryAutofillUnitAndStb(ArrivalItem item)
        {
            if (item == null)
                return;

            // очищаем при смене материала
            item.Unit = null;
            item.Stb = null;

            if (string.IsNullOrWhiteSpace(item.MaterialName))
                return;

            if (journal == null || journal.Count == 0)
                return;

            var last = journal
                .Where(j => j.MaterialName == item.MaterialName)
                .OrderByDescending(j => j.Date)
                .FirstOrDefault();

            if (last == null)
                return;

            item.Unit = last.Unit;
            item.Stb = last.Stb;
        }







        private void Material_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox cb && cb.DataContext is ArrivalItem item)
            {
                TryAutofillUnitAndStb(item);
            }
        }



    }
}

