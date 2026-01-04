using System.Collections.Generic;
using System.Windows;
using ClosedXML.Excel;
using System.Data;
using System.Linq;

namespace ConstructionControl
{
    public partial class ExcelImportWindow : Window
    {
       

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




        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }

}
