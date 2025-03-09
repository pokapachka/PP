using System;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace Prog.Pages
{
    public partial class Main : Page
    {
        private XLWorkbook workbook;
        public Main()
        {
            InitializeComponent();
            dataGrid.LoadingRow += DataGrid_LoadingRow;
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
                workbook?.Dispose(); // Освобождаем предыдущую книгу, если она была открыта
                workbook = new XLWorkbook(filePath);
                cbSheets.Items.Clear();
                // Заполняем ComboBox листами
                foreach (var sheet in workbook.Worksheets)
                {
                    cbSheets.Items.Add(sheet.Name);
                }
                cbSheets.SelectionChanged += (s, e) =>
                {
                    if (cbSheets.SelectedItem != null)
                    {
                        LoadSheet(cbSheets.SelectedItem.ToString());
                    }
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке файла: {ex.Message}");
            }
        }

        private void LoadSheet(string sheetName)
        {
            try
            {
                var worksheet = workbook.Worksheet(sheetName);
                var dataTable = new DataTable();
                dataTable.Columns.Add("ID", typeof(int));
                dataTable.Columns.Add("OP_KZR", typeof(string));
                dataTable.Columns.Add("OP_DSE", typeof(string));
                dataTable.Columns.Add("CNT", typeof(string));
                dataTable.Columns.Add("Flag", typeof(int));      // Флаг для закрашивания строк
                // Получаем диапазон используемых ячеек
                var range = worksheet.RangeUsed();
                // Добавляем строки
                int rowIndex = 1;
                foreach (var row in range.Rows().Skip(1))
                {
                    var newRow = dataTable.NewRow();
                    bool hasEmptyCell = false;
                    bool isCntNotNumber = false;
                    // Заполняем ID
                    newRow["ID"] = rowIndex;
                    // Заполняем остальные столбцы
                    for (int i = 0; i < dataTable.Columns.Count - 2; i++)
                    {
                        var cell = row.Cell(i + 1);
                        var cellValue = cell.Value.ToString();
                        newRow[i + 1] = cellValue;
                        // Проверка на пустую ячейку
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            hasEmptyCell = true;
                        }
                        // Проверка, является ли значение в CNT числом
                        if (dataTable.Columns[i + 1].ColumnName == "CNT" && !string.IsNullOrEmpty(cellValue))
                        {
                            if (!int.TryParse(cellValue, out _))
                            {
                                isCntNotNumber = true;
                            }
                        }
                    }
                    // 0 — зеленый
                    // 1 — есть пустая ячейка или ячейка CNT не число (красный)
                    newRow["Flag"] = (hasEmptyCell || isCntNotNumber) ? 1 : 0;
                    dataTable.Rows.Add(newRow);
                    rowIndex++;
                }

                dataGrid.ItemsSource = dataTable.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке листа: {ex.Message}");
            }
        }
        private DataGridCell GetCell(DataGrid grid, int row, int column)
        {
            DataGridRow rowContainer = (DataGridRow)grid.ItemContainerGenerator.ContainerFromIndex(row);
            if (rowContainer != null)
            {
                DataGridCellsPresenter presenter = FindVisualChild<DataGridCellsPresenter>(rowContainer);
                return (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(column);
            }
            return null;
        }
        private void DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            var row = e.Row;
            // Получаем данные строки
            var dataRow = (e.Row.Item as DataRowView)?.Row;
            if (dataRow != null)
            {
                int flag = Convert.ToInt32(dataRow["Flag"]);
                // Закрашиваем строку в зависимости от значения Flag
                if (flag == 0)
                {
                    row.Background = new SolidColorBrush(Color.FromRgb(204, 255, 204)); // Светло-зелёный
                }
                else if (flag == 1)
                {
                    row.Background = new SolidColorBrush(Color.FromRgb(255, 204, 204)); // Светло-красный
                }
            }
        }

        private T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is T)
                {
                    return (T)child;
                }
                else
                {
                    T childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                    {
                        return childOfChild;
                    }
                }
            }
            return null;
        }
    }
}