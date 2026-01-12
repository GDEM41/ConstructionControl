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
            if (MaterialGroupPanel == null || ExtraTypeBox == null)
                return;

            if (ExtraRadio.IsChecked == true)
            {
                // Допы
                MaterialGroupPanel.Visibility = Visibility.Hidden;

                ExtraTypeBox.Visibility = Visibility.Visible;

                if (ExtraTypeBox.ItemsSource == null)
                {
                    ExtraTypeBox.ItemsSource = new[] { "Внутренние", "Малоценка" };
                    ExtraTypeBox.SelectedIndex = 0;
                }
            }
            else
            {
                // Основные
                MaterialGroupPanel.Visibility = Visibility.Visible;
                ExtraTypeBox.Visibility = Visibility.Hidden;
            }
        }






        public void SetObject(ProjectObject obj, List<JournalRecord> journalRecords)
        {
            currentObject = obj;
            journal = journalRecords;

            MaterialGroupBox.ItemsSource = currentObject.Archive.Groups;


            items.Clear();
            AddRow();
            foreach (var item in items)
            {
                item.AvailableUnits = new ObservableCollection<string>(currentObject.Archive.Units);
            }

        }



        // ================= ТИП МАТЕРИАЛА =================

        private void MaterialGroupTextChanged(object sender, TextChangedEventArgs e)
        {
            var group = MaterialGroupBox.Text?.Trim();

            foreach (var item in items)
            {
                item.AvailableNames.Clear();

                if (string.IsNullOrWhiteSpace(group))
                    continue;

                if (currentObject.Archive.Materials.TryGetValue(group, out var names))
                {
                    foreach (var n in names)
                        item.AvailableNames.Add(n);
                }

            }
            // ДОБАВЛЯЕМ ВРЕМЕННО ВВЕДЁННЫЕ МАТЕРИАЛЫ ИЗ ТЕКУЩИХ СТРОК
            foreach (var it in items)
            {
                if (!string.IsNullOrWhiteSpace(it.MaterialName))
                {
                    foreach (var row in items)
                    {
                        if (!row.AvailableNames.Contains(it.MaterialName))
                            row.AvailableNames.Add(it.MaterialName);
                    }
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

            // заполняем Units из архива
            item.AvailableUnits = new ObservableCollection<string>(currentObject.Archive.Units);

            items.Add(item);
            // КОПИРУЕМ СПИСОК МАТЕРИАЛОВ ИЗ ПЕРВОЙ СТРОКИ В НОВУЮ
            if (items.Count > 1)
            {
                var first = items[0];
                foreach (var n in first.AvailableNames)
                    item.AvailableNames.Add(n);
            }


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

            // === ПОПОЛНЕНИЕ АРХИВА ===
            var archive = currentObject.Archive;

            if (!archive.Groups.Contains(groupName))
                archive.Groups.Add(groupName);

            if (!archive.Materials.ContainsKey(groupName))
                archive.Materials[groupName] = new();

            foreach (var i in items)
            {
                if (!string.IsNullOrWhiteSpace(i.MaterialName) && !archive.Materials[groupName].Contains(i.MaterialName))
                    archive.Materials[groupName].Add(i.MaterialName);

                if (!string.IsNullOrWhiteSpace(i.Unit) && !archive.Units.Contains(i.Unit))
                    archive.Units.Add(i.Unit);

                if (!string.IsNullOrWhiteSpace(i.Supplier) && !archive.Suppliers.Contains(i.Supplier))
                    archive.Suppliers.Add(i.Supplier);

                if (!string.IsNullOrWhiteSpace(i.Passport) && !archive.Passports.Contains(i.Passport))
                    archive.Passports.Add(i.Passport);

                if (!string.IsNullOrWhiteSpace(i.Stb) && !archive.Stb.Contains(i.Stb))
                    archive.Stb.Add(i.Stb);
            }


            // === ПОПОЛНЕНИЕ АРХИВА ===

            foreach (var item in items)
            {
                item.AvailableUnits = new ObservableCollection<string>(archive.Units);
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
            MaterialGroupTextChanged(null, null);

        }
        private void TryAutofillUnitAndStb(ArrivalItem item)
        {
            if (item == null)
                return;

            // очищаем при смене материала
            item.Unit = null;
            item.Stb = null;

            var archive = currentObject.Archive;

            // если в архиве единица одна — ставим автоматом
            if (archive.Units.Count == 1)
            {
                item.Unit = archive.Units[0];
                return;
            }

            // если несколько — формируем список выбора
            item.AvailableUnits = new ObservableCollection<string>(archive.Units);


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

