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

        public ExcelImportWindow(string filePath, List<string> sheets)
        {
            InitializeComponent();
            LoadTemplatesList();

            MainRadio.Checked += ImportTypeChanged;
            ExtraRadio.Checked += ImportTypeChanged;


            _filePath = filePath;
            FilePathText.Text = filePath;
            SheetsList.ItemsSource = sheets;
        }

        private void ImportTypeChanged(object sender, RoutedEventArgs e)
        {
            if (ExtraTypeBox == null)
                return;

            if (ExtraRadio.IsChecked == true)
            {
                ExtraTypeBox.Visibility = Visibility.Visible;
                ExtraTypeBox.ItemsSource = new[] { "Внутренние", "Малоценка" };
                ExtraTypeBox.SelectedIndex = 0;
            }
            else
            {
                ExtraTypeBox.Visibility = Visibility.Collapsed;
            }
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
                        Unit = "шт"
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

            TemplatesCombo.Items.Clear();

            foreach (var file in Directory.GetFiles(TemplatesFolder, "*.json"))
            {
                TemplatesCombo.Items.Add(Path.GetFileNameWithoutExtension(file));
            }
        }
        private void TemplatesCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TemplatesCombo.SelectedItem == null)
                return;

            string name = TemplatesCombo.SelectedItem.ToString();
            string path = Path.Combine(TemplatesFolder, name + ".json");

            try
            {
                var json = File.ReadAllText(path);
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

                MessageBox.Show($"Шаблон «{name}» применён");
            }
            catch
            {
                MessageBox.Show("Ошибка загрузки шаблона");
            }
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
            }
        }
        private void ApplyTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (SheetsList.SelectedItem == null)
            {
                MessageBox.Show("Выберите лист Excel");
                return;
            }

            if (TemplatesCombo.SelectedItem == null)
            {
                MessageBox.Show("Выберите шаблон");
                return;
            }

            string sheetName = SheetsList.SelectedItem.ToString();
            string templateName = TemplatesCombo.SelectedItem.ToString();
            string path = Path.Combine(TemplatesFolder, templateName + ".json");

            var json = File.ReadAllText(path);
            var template = System.Text.Json.JsonSerializer.Deserialize<ExcelImportTemplate>(json);

            if (template == null)
                return;

            _appliedTemplates[sheetName] = template;

            SelectedCellText.Text = $"Лист «{sheetName}» → шаблон «{templateName}» применён";
        }


    }

}
