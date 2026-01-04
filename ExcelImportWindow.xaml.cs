using System.Collections.Generic;
using System.Windows;
using ClosedXML.Excel;
using System.Data;
using System.Linq;
using System.Windows.Controls;

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


        private readonly string _filePath;

        public ExcelImportWindow(string filePath, List<string> sheets)
        {
            InitializeComponent();

            _filePath = filePath;
            FilePathText.Text = filePath;
            SheetsList.ItemsSource = sheets;
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
            if (_dateRow == null || _materialColumn == null || _quantityStartColumn == null)
            {
                MessageBox.Show("Сначала выберите строку дат, колонку материалов и начало количеств");
                return;
            }

            using var wb = new ClosedXML.Excel.XLWorkbook(_filePath);
            var ws = wb.Worksheet(SheetsList.SelectedItem.ToString());

            var range = ws.RangeUsed();
            if (range == null)
                return;

            int lastRow = range.RowCount();
            int lastCol = range.ColumnCount();

            for (int r = _dateRow.Value + 1; r <= lastRow; r++)
            {
                string material = ws.Cell(r, _materialColumn.Value).GetValue<string>();
                if (string.IsNullOrWhiteSpace(material))
                    continue;

                for (int c = _quantityStartColumn.Value; c <= lastCol; c++)
                {
                    if (!double.TryParse(ws.Cell(r, c).GetValue<string>(), out double qty))
                        continue;

                    if (qty <= 0)
                        continue;

                    DateTime date;
                    if (!DateTime.TryParse(ws.Cell(_dateRow.Value, c).GetValue<string>(), out date))
                        continue;

                    ImportedRecords.Add(new JournalRecord
                    {
                        Date = date,
                        MaterialGroup = SheetsList.SelectedItem.ToString(), // ← ИМЯ ЛИСТА
                        MaterialName = material,
                        Quantity = (int)qty,
                        Unit = "шт", // позже сделаем выбор из Excel
                        ObjectName = ""
                    });

                }
            }

            MessageBox.Show($"Импортировано записей: {ImportedRecords.Count}");
            DialogResult = true;
        }


        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
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

    }

}
