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

            // === 1. УДАЛЯЕМ ИЗ АРХИВА ===
            if (obj.Archive.Materials.ContainsKey(group))
            {
                obj.Archive.Materials[group].Remove(material);

                if (obj.Archive.Materials[group].Count == 0)
                    obj.Archive.Materials.Remove(group);
            }
            // === 2. ОБНОВЛЯЕМ ТАБЛИЦУ В АРХИВЕ ===

            LoadArchive();

            // === 3. ОБНОВЛЯЕМ MAINWINDOW ===
            if (Owner is MainWindow mw)
            {
                mw.RefreshTree();
                mw.RefreshJournal();
                mw.RefreshSummaryTable();
            }
        }




        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Очистить архив? Данные проекта и журнала останутся.",
                                 "Подтверждение", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;

            obj.Archive = new ObjectArchive();


            LoadArchive();
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
