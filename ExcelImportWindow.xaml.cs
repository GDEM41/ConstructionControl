using System.Collections.Generic;
using System.Windows;
using ClosedXML.Excel;
using System.Data;
using System.Linq;
using System.Windows.Controls;
using System.IO;

namespace ConstructionControl
{
    public partial class ExcelImportWindow : Window
    {
        private int? _dateRow;
        private int? _materialColumn;
        private int? _quantityStartColumn;
        private int? _positionColumn;
        private int? _unitColumn;
        private int? _volumeColumn;
        private int? _stbColumn;

        private int? _ttnRow;
        private int? _supplierRow;
        private int? _passportRow;
        private readonly Dictionary<string, ExcelImportTemplate> _appliedTemplates
    = new Dictionary<string, ExcelImportTemplate>();



        private readonly string _filePath;
        private readonly ProjectObject _currentObject;

        public bool DemandUpdated { get; private set; }

        public ExcelImportWindow(string filePath, List<string> sheets, ProjectObject currentObject)
        {
            InitializeComponent();
            LoadTemplatesList();

            MainRadio.Checked += ImportTypeChanged;
            ExtraRadio.Checked += ImportTypeChanged;


            _filePath = filePath;
            _currentObject = currentObject;
            FilePathText.Text = filePath;
            SheetsList.ItemsSource = sheets;
            PopulateBlocks();
        }

        private void ImportTypeChanged(object sender, RoutedEventArgs e)
        {
            ExtraTypeBox.Visibility = ExtraRadio.IsChecked == true
                ? Visibility.Visible
                : Visibility.Collapsed;
        }



        private void SheetsList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (SheetsList.SelectedItem == null)
                return;

            string sheetName = SheetsList.SelectedItem.ToString();
            LoadPreview(sheetName);
        }
        private void LoadPreview(string sheetName)
        {
            using var wb = new ClosedXML.Excel.XLWorkbook(_filePath);
            var ws = wb.Worksheet(sheetName);

            var range = ws.RangeUsed();
            if (range == null)
                return;

            var dt = new System.Data.DataTable();

            int maxRow = Math.Min(range.RowCount(), 100);
            int maxCol = range.ColumnCount();




            // колонки
            for (int c = 1; c <= maxCol; c++)
                dt.Columns.Add($"C{c}");

            // строки
            for (int r = 1; r <= maxRow; r++)
            {
                var row = dt.NewRow();
                for (int c = 1; c <= maxCol; c++)
                    row[c - 1] = ws.Cell(r, c).GetValue<string>();

                dt.Rows.Add(row);
            }

            PreviewGrid.ItemsSource = dt.DefaultView;
        }

        private void PreviewGrid_SelectedCellsChanged(object sender, System.Windows.Controls.SelectedCellsChangedEventArgs e)
        {
            if (PreviewGrid.SelectedCells.Count == 0)
                return;

            var cellInfo = PreviewGrid.SelectedCells[0];

            int rowIndex = PreviewGrid.Items.IndexOf(cellInfo.Item) + 1;
            int colIndex = cellInfo.Column.DisplayIndex + 1;

            string excelColumn = ToExcelColumn(colIndex);

            SelectedCellText.Text = $"Выбрана ячейка: {excelColumn}{rowIndex}";
        }

        private void PopulateBlocks()
        {
            if (_currentObject == null)
                return;

            var blocks = Enumerable.Range(1, Math.Max(0, _currentObject.BlocksCount)).ToList();
            BlockSelector.ItemsSource = blocks;
            if (blocks.Count > 0)
                BlockSelector.SelectedIndex = 0;
            PopulateFloors();
        }

        private void PopulateFloors()
        {
            if (_currentObject == null)
                return;

            if (BlockSelector.SelectedItem is not int block)
                return;

            var floors = GetFloorsForBlock(block);
            FloorSelector.ItemsSource = floors;
            if (floors.Count > 0)
                FloorSelector.SelectedIndex = 0;
        }

        private string ToExcelColumn(int columnNumber)
        {
            string columnName = string.Empty;

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        public List<JournalRecord> ImportedRecords { get; } = new();
        private void Import_Click(object sender, RoutedEventArgs e)
        {
            ImportedRecords.Clear();

            using var wb = new XLWorkbook(_filePath);

            // ===== РЕЖИМ 1: ИМПОРТ ПО ШАБЛОНАМ =====
            if (_appliedTemplates.Count > 0)
            {
                foreach (var pair in _appliedTemplates)
                {
                    ImportSheet(wb, pair.Key, pair.Value);
                }
            }
            // ===== РЕЖИМ 2: ИМПОРТ БЕЗ ШАБЛОНА (КАК РАНЬШЕ) =====
            else
            {
                if (_dateRow == null || _materialColumn == null || _quantityStartColumn == null)
                {
                    MessageBox.Show("Настройте импорт кнопками или примените шаблон");
                    return;
                }

                var tempTemplate = new ExcelImportTemplate
                {
                    DateRow = _dateRow.Value,
                    MaterialColumn = _materialColumn.Value,
                    QuantityStartColumn = _quantityStartColumn.Value,
                    PositionColumn = _positionColumn,
                    UnitColumn = _unitColumn,
                    VolumeColumn = _volumeColumn,
                    StbColumn = _stbColumn,
                    TtnRow = _ttnRow,
                    SupplierRow = _supplierRow,
                    PassportRow = _passportRow
                };

                ImportSheet(wb, SheetsList.SelectedItem.ToString(), tempTemplate);
            }

            MessageBox.Show($"Импортировано записей: {ImportedRecords.Count}");
            DialogResult = true;
        }
        private void ImportSheet(XLWorkbook wb, string sheetName, ExcelImportTemplate t)
        {
            var ws = wb.Worksheet(sheetName);
            var range = ws.RangeUsed();
            if (range == null)
                return;

            int lastRow = range.RowCount();
            int lastCol = range.ColumnCount();

            for (int r = t.DateRow + 1; r <= lastRow; r++)
            {
                string material = ws.Cell(r, t.MaterialColumn).GetValue<string>();
                if (string.IsNullOrWhiteSpace(material))
                    continue;

                for (int c = t.QuantityStartColumn; c <= lastCol; c++)
                {
                    if (!double.TryParse(ws.Cell(r, c).GetValue<string>(), out double qty))
                        continue;

                    if (qty <= 0)
                        continue;

                    if (!DateTime.TryParse(ws.Cell(t.DateRow, c).GetValue<string>(), out DateTime date))
                        continue;

                    ImportedRecords.Add(new JournalRecord
                    {
                        Date = date,
                        Category = MainRadio.IsChecked == true ? "Основные" : "Допы",
                        SubCategory = ExtraRadio.IsChecked == true
                             ? ExtraTypeBox.SelectedItem?.ToString()
                             : null,

                        MaterialGroup = MainRadio.IsChecked == true ? sheetName : null,
                        MaterialName = material,

                        Quantity = qty,
                        Unit = t.UnitColumn != null ? ws.Cell(r, t.UnitColumn.Value).GetValue<string>() : "шт",

                        Position = t.PositionColumn != null ? ws.Cell(r, t.PositionColumn.Value).GetValue<string>() : null,
                        Volume = t.VolumeColumn != null ? ws.Cell(r, t.VolumeColumn.Value).GetValue<string>() : null,

                        Passport = t.PassportRow != null ? ws.Cell(t.PassportRow.Value, c).GetValue<string>() : null,
                        Supplier = t.SupplierRow != null ? ws.Cell(t.SupplierRow.Value, c).GetValue<string>() : null,
                        Ttn = t.TtnRow != null ? ws.Cell(t.TtnRow.Value, c).GetValue<string>() : null,
                        Stb = t.StbColumn != null ? ws.Cell(r, t.StbColumn.Value).GetValue<string>() : null,
                    });



                }
            }
        }



        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
        private void LoadTemplate()
        {
            if (!File.Exists("excel_template.json"))
                return;

            try
            {
                var json = File.ReadAllText("excel_template.json");
                var template = System.Text.Json.JsonSerializer.Deserialize<ExcelImportTemplate>(json);

                if (template == null)
                    return;

                _dateRow = template.DateRow;
                _materialColumn = template.MaterialColumn;
                _quantityStartColumn = template.QuantityStartColumn;

                _positionColumn = template.PositionColumn;
                _unitColumn = template.UnitColumn;
                _volumeColumn = template.VolumeColumn;
                _stbColumn = template.StbColumn;

                _ttnRow = template.TtnRow;
                _supplierRow = template.SupplierRow;
                _passportRow = template.PassportRow;
            }
            catch
            {
                MessageBox.Show("Не удалось загрузить шаблон импорта");
            }
        }

        private void SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            

            if (_dateRow == null || _materialColumn == null || _quantityStartColumn == null)
            {
                MessageBox.Show("Сначала настройте импорт кнопками");
                return;
            }

            var template = new ExcelImportTemplate
            {
                DateRow = _dateRow.Value,
                MaterialColumn = _materialColumn.Value,
                QuantityStartColumn = _quantityStartColumn.Value,

                PositionColumn = _positionColumn,
                UnitColumn = _unitColumn,
                VolumeColumn = _volumeColumn,
                StbColumn = _stbColumn,

                TtnRow = _ttnRow,
                SupplierRow = _supplierRow,
                PassportRow = _passportRow
            };

            var json = System.Text.Json.JsonSerializer.Serialize(
                template,
                new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

            string name = Microsoft.VisualBasic.Interaction.InputBox(
    "Введите имя шаблона:",
    "Сохранение шаблона");

            if (string.IsNullOrWhiteSpace(name))
                return;

            if (!Directory.Exists(TemplatesFolder))
                Directory.CreateDirectory(TemplatesFolder);

            string path = Path.Combine(TemplatesFolder, name + ".json");
            File.WriteAllText(path, json);

            LoadTemplatesList();

            MessageBox.Show($"Шаблон «{name}» сохранён");

        }

        private const string TemplatesFolder = "Templates";

        private void LoadTemplatesList()
        {
            if (!Directory.Exists(TemplatesFolder))
                Directory.CreateDirectory(TemplatesFolder);

            TemplatesBox.Items.Clear();

            foreach (var file in Directory.GetFiles(TemplatesFolder, "*.json"))
                TemplatesBox.Items.Add(Path.GetFileNameWithoutExtension(file));
        }
        private void TemplatesBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TemplatesBox.SelectedItem == null)
                return;

            string name = TemplatesBox.SelectedItem.ToString();
            CurrentTemplateText.Text = $"Выбран: {name}";
        }
        private void DeleteTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (TemplatesBox.SelectedItem == null)
            {
                MessageBox.Show("Выберите шаблон для удаления.");
                return;
            }

            string name = TemplatesBox.SelectedItem.ToString();
            string path = Path.Combine(TemplatesFolder, name + ".json");

            if (File.Exists(path))
                File.Delete(path);

            LoadTemplatesList();
            CurrentTemplateText.Text = "";
        }


        private void RefreshTemplates_Click(object sender, RoutedEventArgs e)
        {
            LoadTemplatesList();
            CurrentTemplateText.Text = "";
        }

        private void BlockSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            PopulateFloors();
        }


        private void SelectCell_Click(object sender, RoutedEventArgs e)
        {
            if (PreviewGrid.SelectedCells.Count == 0)
            {
                MessageBox.Show("Сначала выберите ячейку в таблице");
                return;
            }

            var button = sender as Button;
            if (button == null)
                return;

            var cell = PreviewGrid.SelectedCells[0];

            if (cell.Column == null)
            {
                MessageBox.Show("Выберите ЯЧЕЙКУ, а не строку");
                return;
            }


            int row = PreviewGrid.Items.IndexOf(cell.Item) + 1;
            int col = cell.Column.DisplayIndex + 1;

            switch (button.Tag?.ToString())
            {
                case "Date":
                    _dateRow = row;
                    MessageBox.Show($"📅 Дата: строка {row} → вправо");
                    break;

                case "Material":
                    _materialColumn = col;
                    MessageBox.Show($"🧱 Наименование: колонка {ToExcelColumn(col)} → вниз");
                    break;

                case "Quantity":
                    _quantityStartColumn = col;
                    MessageBox.Show($"🔢 Кол-во: {ToExcelColumn(col)}{row}");
                    break;

                case "Position":
                    _positionColumn = col;
                    MessageBox.Show($"🧾 Позиция: колонка {ToExcelColumn(col)}");
                    break;

                case "Unit":
                    _unitColumn = col;
                    MessageBox.Show($"📐 Ед. изм: колонка {ToExcelColumn(col)}");
                    break;

                case "Volume":
                    _volumeColumn = col;
                    MessageBox.Show($"📦 Объем: колонка {ToExcelColumn(col)}");
                    break;

                case "Stb":
                    _stbColumn = col;
                    MessageBox.Show($"🏷 СТБ: колонка {ToExcelColumn(col)}");
                    break;

                case "Ttn":
                    _ttnRow = row;
                    MessageBox.Show($"🚚 ТТН: строка {row} → вправо");
                    break;

                case "Supplier":
                    _supplierRow = row;
                    MessageBox.Show($"🏭 Поставщик: строка {row} → вправо");
                    break;

                case "Passport":
                    _passportRow = row;
                    MessageBox.Show($"📄 Паспорт: строка {row} → вправо");
                    break;
                case "DemandRange":
                    ApplyDemandRange();
                    break;
                case "DemandColumn":
                    ApplyDemandColumn();
                    break;
            }
        }
        private void ApplyDemandColumn()
        {
            if (_currentObject == null)
            {
                MessageBox.Show("Сначала выберите объект.");
                return;
            }

            if (_materialColumn == null)
            {
                MessageBox.Show("Сначала укажите колонку «Наименование».");
                return;
            }

            if (BlockSelector.SelectedItem is not int block)
            {
                MessageBox.Show("Выберите блок для заполнения.");
                return;
            }

            if (FloorSelector.SelectedItem is not int floor)
            {
                MessageBox.Show("Выберите этаж для заполнения.");
                return;
            }

            if (SheetsList.SelectedItem == null)
            {
                MessageBox.Show("Выберите лист Excel.");
                return;
            }

            if (PreviewGrid.SelectedCells.Count == 0)
            {
                MessageBox.Show("Сначала выберите верхнюю ячейку столбца.");
                return;
            }

            var selectedCell = PreviewGrid.SelectedCells[0];
            if (selectedCell.Column == null)
            {
                MessageBox.Show("Выберите ЯЧЕЙКУ, а не строку.");
                return;
            }

            int startRow = PreviewGrid.Items.IndexOf(selectedCell.Item) + 1;
            int quantityColumn = selectedCell.Column.DisplayIndex + 1;

            string sheetName = SheetsList.SelectedItem.ToString();
            using var wb = new XLWorkbook(_filePath);
            var ws = wb.Worksheet(sheetName);
            var range = ws.RangeUsed();
            if (range == null)
                return;

            int lastRow = range.RowCount();
            string group = sheetName;

            for (int r = startRow; r <= lastRow; r++)
            {
                string material = ws.Cell(r, _materialColumn.Value).GetValue<string>().Trim();
                if (string.IsNullOrWhiteSpace(material))
                    continue;

                if (!double.TryParse(ws.Cell(r, quantityColumn).GetValue<string>(), out double value))
                    continue;

                string unit = _unitColumn != null
                    ? ws.Cell(r, _unitColumn.Value).GetValue<string>()
                    : "шт";

                string demandKey = $"{group}::{material}";
                if (!_currentObject.Demand.TryGetValue(demandKey, out var demand))
                {
                    demand = new MaterialDemand
                    {
                        Unit = unit,
                        Floors = new Dictionary<int, Dictionary<int, double>>()
                    };
                    _currentObject.Demand[demandKey] = demand;
                }

                if (string.IsNullOrWhiteSpace(demand.Unit))
                    demand.Unit = unit;

                if (!demand.Floors.ContainsKey(block))
                    demand.Floors[block] = new Dictionary<int, double>();

                demand.Floors[block][floor] = value;
                EnsureMaterialGroup(group, material);
            }

            DemandUpdated = true;
            MessageBox.Show($"Кол-во по этажу {floor} для блока {block} импортировано.");
        }

        private void ApplyDemandRange()
        {
            if (_currentObject == null)
            {
                MessageBox.Show("Сначала выберите объект.");
                return;
            }

            if (_materialColumn == null)
            {
                MessageBox.Show("Сначала укажите колонку «Наименование».");
                return;
            }

            if (BlockSelector.SelectedItem is not int block)
            {
                MessageBox.Show("Выберите блок для заполнения.");
                return;
            }

            if (SheetsList.SelectedItem == null)
            {
                MessageBox.Show("Выберите лист Excel.");
                return;
            }

            var selectedCells = PreviewGrid.SelectedCells
                .Where(c => c.Column != null)
                .OrderBy(c => c.Column.DisplayIndex)
                .ToList();

            if (selectedCells.Count == 0)
            {
                MessageBox.Show("Выберите диапазон значений для этажей.");
                return;
            }

            int firstRow = PreviewGrid.Items.IndexOf(selectedCells[0].Item);
            if (selectedCells.Any(c => PreviewGrid.Items.IndexOf(c.Item) != firstRow))
            {
                MessageBox.Show("Выберите значения в одной строке (один материал).");
                return;
            }

            var rowView = selectedCells[0].Item as DataRowView;
            if (rowView == null)
            {
                MessageBox.Show("Не удалось прочитать строку.");
                return;
            }

            string material = rowView.Row[_materialColumn.Value - 1]?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(material))
            {
                MessageBox.Show("В выбранной строке нет наименования материала.");
                return;
            }

            string unit = _unitColumn != null
                ? rowView.Row[_unitColumn.Value - 1]?.ToString()
                : "шт";

            var floors = GetFloorsForBlock(block);
            if (floors.Count == 0)
            {
                MessageBox.Show("В выбранном блоке нет этажей.");
                return;
            }

            if (selectedCells.Count != floors.Count)
            {
                MessageBox.Show($"Выберите {floors.Count} ячеек по этажам. Сейчас выбрано: {selectedCells.Count}.");
                return;
            }

            string group = SheetsList.SelectedItem.ToString();
            string demandKey = $"{group}::{material}";

            if (!_currentObject.Demand.TryGetValue(demandKey, out var demand))
            {
                demand = new MaterialDemand
                {
                    Unit = unit,
                    Floors = new Dictionary<int, Dictionary<int, double>>()
                };
                _currentObject.Demand[demandKey] = demand;
            }

            if (string.IsNullOrWhiteSpace(demand.Unit))
                demand.Unit = unit;

            if (!demand.Floors.ContainsKey(block))
                demand.Floors[block] = new Dictionary<int, double>();

            for (int i = 0; i < floors.Count; i++)
            {
                double value = ParseCellValue(selectedCells[i]);
                demand.Floors[block][floors[i]] = value;
            }
            EnsureMaterialGroup(group, material);
            DemandUpdated = true;

            MessageBox.Show($"Значения по этажам для блока {block} обновлены.");
        }

        private List<int> GetFloorsForBlock(int block)
        {
            var list = new List<int>();
            if (_currentObject.HasBasement)
                list.Add(0);

            int floors = _currentObject.SameFloorsInBlocks
                ? _currentObject.FloorsPerBlock
                : (_currentObject.FloorsByBlock.TryGetValue(block, out var f) ? f : 0);

            for (int floor = 1; floor <= floors; floor++)
                list.Add(floor);

            return list;
        }

        private double ParseCellValue(DataGridCellInfo cell)
        {
            if (cell.Item is not DataRowView rowView)
                return 0;

            int columnIndex = cell.Column.DisplayIndex;
            string text = rowView.Row[columnIndex]?.ToString() ?? string.Empty;

            if (double.TryParse(text, out var value))
                return value;

            return 0;
        }

        private void EnsureMaterialGroup(string group, string material)
        {
            if (!_currentObject.MaterialGroups.Any(g => g.Name == group))
            {
                _currentObject.MaterialGroups.Add(new MaterialGroup
                {
                    Name = group
                });
            }

            if (!_currentObject.MaterialNamesByGroup.ContainsKey(group))
                _currentObject.MaterialNamesByGroup[group] = new List<string>();

            if (!_currentObject.MaterialNamesByGroup[group].Contains(material))
                _currentObject.MaterialNamesByGroup[group].Add(material);
        }
        private void ApplyTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (SheetsList.SelectedItem == null)
            {
                MessageBox.Show("Выберите лист Excel.");
                return;
            }

            if (TemplatesBox.SelectedItem == null)
            {
                MessageBox.Show("Выберите шаблон.");
                return;
            }

            string sheetName = SheetsList.SelectedItem.ToString();
            string templateName = TemplatesBox.SelectedItem.ToString();
            string path = Path.Combine(TemplatesFolder, templateName + ".json");

            var json = File.ReadAllText(path);
            var template = System.Text.Json.JsonSerializer.Deserialize<ExcelImportTemplate>(json);

            if (template == null)
                return;

            _appliedTemplates[sheetName] = template;

            CurrentTemplateText.Text = $"Применён: {templateName}";
        }



    }

}
