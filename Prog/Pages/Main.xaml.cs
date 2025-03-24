using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using Prog.Connection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prog.Pages
{
    public partial class Main : Page
    {
        private Excel.Application excelApp;
        private Excel.Workbook workbook;

        public Main()
        {
            InitializeComponent();
        }

        private void ChooseFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                Title = "Выберите файл"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                txtFilePath.Text = openFileDialog.FileName;
                LoadExcelFile(openFileDialog.FileName);
            }
        }

        private void LoadExcelFile(string filePath)
        {
            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                cbSheets.Items.Clear();

                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    cbSheets.Items.Add(sheet.Name);
                }

                cbSheets.SelectionChanged += (s, e) =>
                    LoadSheet(cbSheets.SelectedItem.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке файла: {ex.Message}");
                CloseExcel();
            }
        }

        private void LoadSheet(string sheetName)
        {
            try
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
                Excel.Range usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;
                var dataTable = new DataTable();
                dataTable.Columns.Add("ID", typeof(int));
                dataTable.Columns.Add("OP_KZR", typeof(string));
                dataTable.Columns.Add("OP_DSE", typeof(string));
                dataTable.Columns.Add("CNT", typeof(string));
                dataTable.Columns.Add("Flag", typeof(int)); // Для подсветки строк
                // Чтение данных (начиная со 2 строки, так как 1 строка - заголовки)
                for (int row = 2; row <= rowCount; row++)
                {
                    var newRow = dataTable.NewRow();
                    bool hasEmptyCell = false;
                    bool isCntNotNumber = false;
                    newRow["ID"] = row - 1;
                    newRow["OP_KZR"] = usedRange.Cells[row, 1]?.Value2?.ToString() ?? string.Empty;
                    newRow["OP_DSE"] = usedRange.Cells[row, 2]?.Value2?.ToString() ?? string.Empty;
                    newRow["CNT"] = usedRange.Cells[row, 3]?.Value2?.ToString() ?? string.Empty;
                    // Проверка пустых значений
                    if (string.IsNullOrEmpty(newRow["OP_KZR"].ToString())) hasEmptyCell = true;
                    if (string.IsNullOrEmpty(newRow["OP_DSE"].ToString())) hasEmptyCell = true;
                    if (string.IsNullOrEmpty(newRow["CNT"].ToString())) hasEmptyCell = true;
                    // Проверка, является ли CNT числом
                    if (!string.IsNullOrEmpty(newRow["CNT"].ToString()) && !int.TryParse(newRow["CNT"].ToString(), out _))
                    {
                        isCntNotNumber = true;
                    }
                    // Устанавливаем флаг (0 - OK, 1 - ошибка)
                    newRow["Flag"] = (hasEmptyCell || isCntNotNumber) ? 1 : 0;

                    dataTable.Rows.Add(newRow);
                }
                dataGrid.ItemsSource = dataTable.DefaultView;
                dataGrid.LoadingRow += DataGrid_LoadingRow;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке листа: {ex.Message}");
            }
        }

        private void DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            var row = e.Row;
            var dataRow = (e.Row.Item as DataRowView)?.Row;
            if (dataRow != null)
            {
                int flag = Convert.ToInt32(dataRow["Flag"]);
                if (flag == 1)
                {
                    row.Background = new SolidColorBrush(Color.FromRgb(255, 200, 200)); // Красный
                }
                else
                {
                    row.Background = new SolidColorBrush(Color.FromRgb(200, 255, 200)); // Зеленый
                }
            }
        }

        private void CloseExcel()
        {
            try
            {
                workbook?.Close(false);
                excelApp?.Quit();
                // Освобождаем ресурсы
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                workbook = null;
                excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при закрытии Excel: {ex.Message}");
            }
        }

        private void UploadData(object sender, RoutedEventArgs e)
        {
            var db = new ConnectionDB();
            db.ProcessGreenRecords();
            db.CloseConnection();
            MessageBox.Show("Обработка завершена!");
        }
    }
}
