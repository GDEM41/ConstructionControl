using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace ConstructionControl
{
    public partial class ArchiveWindow : Window
    {
        private ProjectObject obj;
        private List<JournalRecord> journal;




        private void LoadArchive()
        {
            var list = new List<ArchiveRecord>();

            foreach (var g in obj.Archive.Materials)
            {
                foreach (var mat in g.Value)
                {
                    var entries = journal
                        .Where(j => j.MaterialGroup == g.Key && j.MaterialName == mat)
                        .ToList();

                    list.Add(new ArchiveRecord
                    {
                        Group = g.Key,
                        Material = mat,
                        Unit = entries.Select(x => x.Unit).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)),
                        Suppliers = string.Join(", ", entries.Select(x => x.Supplier).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()),
                        Passports = string.Join(", ", entries.Select(x => x.Passport).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()),
                        Stb = string.Join(", ", entries.Select(x => x.Stb).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()),
                        LastArrival = entries.OrderByDescending(x => x.Date).Select(x => x.Date).FirstOrDefault(),
                        ArrivalsCount = entries.Count
                    });
                }
            }

            ArchiveGrid.ItemsSource = list;
        }




        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (ArchiveGrid.SelectedItem is not ArchiveRecord r)
                return;

            string group = r.Group;
            string material = r.Material;

            journal.RemoveAll(j => j.MaterialGroup == group && j.MaterialName == material);
            obj.Demand.Remove($"{group}::{material}");
            RebuildObjectFromJournal();

            LoadArchive();

            
            if (Owner is MainWindow mw)
            {
                mw.RefreshAfterArchiveChange();
            }
        }




        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Очистить архив? Будут удалены данные дерева и журналов.",
                     "Подтверждение", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;

            journal.Clear();
            obj.Demand.Clear();
            obj.MaterialGroups.Clear();
            obj.MaterialNamesByGroup.Clear();
            obj.SummaryVisibleGroups.Clear();
            obj.Archive = new ObjectArchive();


            LoadArchive();

            if (Owner is MainWindow mw)
            {
                mw.RefreshAfterArchiveChange();
            }
        }

        private void RebuildObjectFromJournal()
        {
            var mainRecords = journal
                .Where(j => j.Category == "Основные" && !string.IsNullOrWhiteSpace(j.MaterialGroup))
                .ToList();

            var groups = mainRecords
                .GroupBy(j => j.MaterialGroup)
                .ToList();

            obj.MaterialGroups = groups
                .Select(g => new MaterialGroup
                {
                    Name = g.Key,
                    Items = g.Select(j => j.MaterialName)
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct()
                        .OrderBy(x => x)
                        .ToList()
                })
                .ToList();

            obj.MaterialNamesByGroup = groups.ToDictionary(
                g => g.Key,
                g => g.Select(j => j.MaterialName)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList());

            obj.SummaryVisibleGroups = obj.SummaryVisibleGroups
                .Where(g => obj.MaterialNamesByGroup.ContainsKey(g))
                .ToList();

            var archive = new ObjectArchive();

            foreach (var record in journal)
            {
                if (!string.IsNullOrWhiteSpace(record.MaterialGroup))
                {
                    if (!archive.Groups.Contains(record.MaterialGroup))
                        archive.Groups.Add(record.MaterialGroup);

                    if (!archive.Materials.ContainsKey(record.MaterialGroup))
                        archive.Materials[record.MaterialGroup] = new List<string>();

                    if (!string.IsNullOrWhiteSpace(record.MaterialName)
                        && !archive.Materials[record.MaterialGroup].Contains(record.MaterialName))
                        archive.Materials[record.MaterialGroup].Add(record.MaterialName);
                }

                if (!string.IsNullOrWhiteSpace(record.Unit) && !archive.Units.Contains(record.Unit))
                    archive.Units.Add(record.Unit);

                if (!string.IsNullOrWhiteSpace(record.Supplier) && !archive.Suppliers.Contains(record.Supplier))
                    archive.Suppliers.Add(record.Supplier);

                if (!string.IsNullOrWhiteSpace(record.Passport) && !archive.Passports.Contains(record.Passport))
                    archive.Passports.Add(record.Passport);

                if (!string.IsNullOrWhiteSpace(record.Stb) && !archive.Stb.Contains(record.Stb))
                    archive.Stb.Add(record.Stb);
            }

            obj.Archive = archive;

            var validKeys = new HashSet<string>(
                mainRecords.Select(r => $"{r.MaterialGroup}::{r.MaterialName}"));

            foreach (var key in obj.Demand.Keys.Where(k => !validKeys.Contains(k)).ToList())
                obj.Demand.Remove(key);
        }

        public class ArchiveRecord
        {
            public string Group { get; set; }
            public string Material { get; set; }
            public string Unit { get; set; }
            public string Suppliers { get; set; }
            public string Passports { get; set; }
            public string Stb { get; set; }
            public DateTime? LastArrival { get; set; }
            public int ArrivalsCount { get; set; }
        }
        public ArchiveWindow(ProjectObject obj, List<JournalRecord> journal)
        {
            InitializeComponent();

            // сохраняем входящие параметры в поля
            this.obj = obj ?? throw new ArgumentNullException(nameof(obj));
            this.journal = journal ?? new List<JournalRecord>();

            // загружаем таблицу
            LoadArchive();
        }




    }
}
