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
                        Unit = entries.Select(x => x.Unit).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().FirstOrDefault(),
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
            var row = ArchiveGrid.SelectedItem;
            if (row == null) return;

            dynamic r = row;
            string group = r.Group;
            string material = r.Material;

            // 1. Удаляем из архива
            obj.Archive.Materials[group].Remove(material);
            if (obj.Archive.Materials[group].Count == 0)
                obj.Archive.Materials.Remove(group);

            if (!obj.Archive.Materials.Any())
                obj.Archive.Groups.Remove(group);

            // 2. Удаляем из MaterialGroups / MaterialNamesByGroup
            obj.MaterialNamesByGroup[group].Remove(material);
            if (obj.MaterialNamesByGroup[group].Count == 0)
            {
                obj.MaterialNamesByGroup.Remove(group);
                obj.MaterialGroups.RemoveAll(g => g.Name == group);
            }

            // 3. Удаляем из журнала
            var journal = Owner is MainWindow mw ? mw.GetJournal() : null;
            if (journal != null)
                journal.RemoveAll(j => j.MaterialName == material && j.MaterialGroup == group);

            DialogResult = true;
            Close();
        }


        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Очистить архив и удалить данные журнала?",
                "Подтверждение", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;

            // 1. Чистим архив
            obj.Archive = new ObjectArchive();

            // 2. Чистим группы/материалы
            obj.MaterialGroups.Clear();
            obj.MaterialNamesByGroup.Clear();

            // 3. Чистим журнал
            var journal = Owner is MainWindow mw ? mw.GetJournal() : null;
            journal?.Clear();

            DialogResult = true;
            Close();
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
                        Unit = entries.Select(x => x.Unit).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().FirstOrDefault(),
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



    }
}
