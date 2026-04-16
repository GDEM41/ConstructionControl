using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ConstructionControl
{
    public partial class ArrivalControl : UserControl
    {
        private ProjectObject currentObject;
        private List<JournalRecord> journal;
        private readonly ObservableCollection<ArrivalItem> items = new();

        public event System.Action<Arrival> ArrivalAdded;

        public ArrivalControl()
        {
            InitializeComponent();
            ItemsGrid.ItemsSource = items;
        }

        private string SelectedCategory
        {
            get
            {
                if (MainRadio?.IsChecked == true)
                    return "Основные";
                if (LowCostRadio?.IsChecked == true)
                    return "Малоценка";
                return "Внутренние";
            }
        }

        private void ArrivalTypeChanged(object sender, RoutedEventArgs e)
        {
            RefreshAllRowLookups();
        }

        public void SetObject(ProjectObject obj, List<JournalRecord> journalRecords)
        {
            currentObject = obj;
            journal = journalRecords;

            items.Clear();
            AddRow();
            RefreshAllRowLookups();
        }

        private void RefreshAllRowLookups()
        {
            foreach (var item in items)
            {
                item.AvailableUnits = new ObservableCollection<string>(currentObject?.Archive?.Units ?? new List<string>());
                RefreshAvailableGroups(item);
                RefreshAvailableNames(item);
            }
        }

        private void RefreshAvailableGroups(ArrivalItem item)
        {
            item.AvailableGroups.Clear();
            if (currentObject?.Archive?.Groups == null)
                return;

            foreach (var group in currentObject.Archive.Groups
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct()
                .OrderBy(x => x))
            {
                item.AvailableGroups.Add(group);
            }
        }

        private void RefreshAvailableNames(ArrivalItem item)
        {
            if (item == null)
                return;

            item.AvailableNames.Clear();
            var group = item.MaterialGroup?.Trim();
            if (string.IsNullOrWhiteSpace(group))
                return;

            if (currentObject?.Archive?.Materials != null
                && currentObject.Archive.Materials.TryGetValue(group, out var archiveNames))
            {
                foreach (var name in archiveNames
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct()
                    .OrderBy(x => x))
                {
                    if (!item.AvailableNames.Contains(name))
                        item.AvailableNames.Add(name);
                }
            }

            if (currentObject?.MaterialCatalog != null)
            {
                var names = currentObject.MaterialCatalog
                    .Where(x => !string.IsNullOrWhiteSpace(x.MaterialName))
                    .Where(x => IsMatchingCatalogEntry(x, group))
                    .Select(x => x.MaterialName)
                    .Distinct()
                    .OrderBy(x => x);

                foreach (var name in names)
                {
                    if (!item.AvailableNames.Contains(name))
                        item.AvailableNames.Add(name);
                }
            }
        }

        private bool IsMatchingCatalogEntry(MaterialCatalogItem entry, string group)
        {
            var category = SelectedCategory;
            if (string.Equals(category, "Основные", System.StringComparison.CurrentCultureIgnoreCase))
            {
                return string.Equals(entry.CategoryName, "Основные", System.StringComparison.CurrentCultureIgnoreCase)
                    && string.Equals(entry.TypeName, group, System.StringComparison.CurrentCultureIgnoreCase);
            }

            var matchesNewCategory = string.Equals(entry.CategoryName, category, System.StringComparison.CurrentCultureIgnoreCase)
                && string.Equals(entry.TypeName, group, System.StringComparison.CurrentCultureIgnoreCase);

            var matchesLegacyCategory = string.Equals(entry.CategoryName, "Допы", System.StringComparison.CurrentCultureIgnoreCase)
                && string.Equals(entry.TypeName, category, System.StringComparison.CurrentCultureIgnoreCase)
                && string.Equals(entry.SubTypeName, group, System.StringComparison.CurrentCultureIgnoreCase);

            return matchesNewCategory || matchesLegacyCategory;
        }

        private void AddRow()
        {
            var defaultGroup = items.LastOrDefault(x => !string.IsNullOrWhiteSpace(x.MaterialGroup))?.MaterialGroup
                ?? items.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.MaterialGroup))?.MaterialGroup;

            var item = new ArrivalItem
            {
                Date = System.DateTime.Today,
                MaterialGroup = defaultGroup,
                AvailableGroups = new ObservableCollection<string>(),
                AvailableNames = new ObservableCollection<string>(),
                AvailableUnits = new ObservableCollection<string>(currentObject?.Archive?.Units ?? new List<string>())
            };

            items.Add(item);
            RefreshAvailableGroups(item);
            RefreshAvailableNames(item);

            item.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(ArrivalItem.MaterialGroup))
                    RefreshAvailableNames(item);

                if (e.PropertyName == nameof(ArrivalItem.MaterialName))
                    TryAutofillUnitAndStb(item);
            };
        }

        private void AddRow_Click(object sender, RoutedEventArgs e) => AddRow();

        private void AddArrival_Click(object sender, RoutedEventArgs e)
        {
            if (currentObject == null)
            {
                MessageBox.Show("Сначала создайте объект");
                return;
            }

            var rows = items.Where(x =>
                    !string.IsNullOrWhiteSpace(x.MaterialName)
                    || !string.IsNullOrWhiteSpace(x.MaterialGroup)
                    || x.Quantity > 0)
                .ToList();

            if (rows.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы одну заполненную строку.");
                return;
            }

            foreach (var row in rows)
            {
                if (string.IsNullOrWhiteSpace(row.MaterialGroup))
                {
                    MessageBox.Show("Укажите тип материала в каждой заполненной строке.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(row.MaterialName))
                {
                    MessageBox.Show("Укажите наименование в каждой заполненной строке.");
                    return;
                }
            }

            var archive = currentObject.Archive;
            foreach (var row in rows)
            {
                var groupName = row.MaterialGroup.Trim();

                if (!currentObject.MaterialGroups.Any(g => g.Name == groupName))
                    currentObject.MaterialGroups.Add(new MaterialGroup { Name = groupName });

                if (!currentObject.MaterialNamesByGroup.ContainsKey(groupName))
                    currentObject.MaterialNamesByGroup[groupName] = new();

                if (!currentObject.MaterialNamesByGroup[groupName].Contains(row.MaterialName))
                    currentObject.MaterialNamesByGroup[groupName].Add(row.MaterialName);

                if (!archive.Groups.Contains(groupName))
                    archive.Groups.Add(groupName);

                if (!archive.Materials.ContainsKey(groupName))
                    archive.Materials[groupName] = new List<string>();

                if (!archive.Materials[groupName].Contains(row.MaterialName))
                    archive.Materials[groupName].Add(row.MaterialName);

                if (!string.IsNullOrWhiteSpace(row.Unit) && !archive.Units.Contains(row.Unit))
                    archive.Units.Add(row.Unit);

                if (!string.IsNullOrWhiteSpace(row.Supplier) && !archive.Suppliers.Contains(row.Supplier))
                    archive.Suppliers.Add(row.Supplier);

                if (!string.IsNullOrWhiteSpace(row.Passport) && !archive.Passports.Contains(row.Passport))
                    archive.Passports.Add(row.Passport);

                if (!string.IsNullOrWhiteSpace(row.Stb) && !archive.Stb.Contains(row.Stb))
                    archive.Stb.Add(row.Stb);
            }

            ArrivalAdded?.Invoke(new Arrival
            {
                Category = SelectedCategory,
                SubCategory = null,
                TtnNumber = TtnBox.Text?.Trim(),
                Items = rows.ToList()
            });

            TtnBox.Clear();
            items.Clear();
            AddRow();
            RefreshAllRowLookups();
        }

        private void TryAutofillUnitAndStb(ArrivalItem item)
        {
            if (item == null || currentObject?.Archive == null)
                return;

            item.Unit = null;
            item.Stb = null;
            item.Supplier = null;
            item.AvailableUnits = new ObservableCollection<string>(currentObject.Archive.Units);

            if (string.IsNullOrWhiteSpace(item.MaterialName) || journal == null || journal.Count == 0)
                return;

            var last = journal
                .Where(j => string.Equals(j.MaterialName, item.MaterialName, System.StringComparison.CurrentCultureIgnoreCase))
                .Where(j => string.IsNullOrWhiteSpace(item.MaterialGroup)
                    || string.Equals(j.MaterialGroup, item.MaterialGroup, System.StringComparison.CurrentCultureIgnoreCase))
                .OrderByDescending(j => j.Date)
                .FirstOrDefault();

            if (last == null)
                return;

            item.Unit = last.Unit;
            item.Stb = last.Stb;
            item.Supplier = last.Supplier;
        }

        private void MaterialGroup_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox combo && combo.DataContext is ArrivalItem item)
                RefreshAvailableNames(item);
        }

        private void Material_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox cb && cb.DataContext is ArrivalItem item)
                TryAutofillUnitAndStb(item);
        }
    }
}
