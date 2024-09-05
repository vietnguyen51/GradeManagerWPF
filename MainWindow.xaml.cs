using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Windows.Data;

namespace GradingEditor
{
    public partial class MainWindow : Window
    {
        private Dictionary<string, DataTable> _sheetData = new Dictionary<string, DataTable>();
        private DataTable _dataTable;
        private string _filePath;
        private bool _isDataSaved = false;

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.DataContext = this;
        }

        private void OnSelectFileClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _filePath = openFileDialog.FileName;
                FilePathTextBox.Text = _filePath;
                LoadExcelSheets(_filePath);
            }
        }

        private void LoadExcelSheets(string filePath)
        {
            _sheetData.Clear();
            SheetSelectionComboBox.Items.Clear();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    SheetSelectionComboBox.Items.Add(worksheet.Name);
                    _sheetData[worksheet.Name] = CreateDataTableFromWorksheet(worksheet);
                }
            }

            if (SheetSelectionComboBox.Items.Count > 0)
            {
                SheetSelectionComboBox.SelectedIndex = 0;
            }
        }

        private DataTable CreateDataTableFromWorksheet(ExcelWorksheet worksheet)
{
    DataTable dataTable = new DataTable();

    // Add columns based on Excel headers
    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
    {
        string columnName = worksheet.Cells[1, col].Text.Trim();
        if (!string.IsNullOrEmpty(columnName))
        {
            dataTable.Columns.Add(columnName, typeof(object));
        }
    }

    // Ensure required columns exist
    if (!dataTable.Columns.Contains("ID"))
        dataTable.Columns.Add("ID", typeof(int));
    if (!dataTable.Columns.Contains("Role"))
        dataTable.Columns.Add("Role", typeof(string));
    if (!dataTable.Columns.Contains("Name"))
        dataTable.Columns.Add("Name", typeof(string));

    // Read rows from Excel and add to DataTable
    int lastRow = FindLastRow(worksheet);
    for (var rowNumber = 2; rowNumber <= lastRow; rowNumber++)
    {
        var newRow = dataTable.NewRow();
        bool hasData = false;
        
        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
        {
            string columnName = worksheet.Cells[1, col].Text.Trim();
            if (dataTable.Columns.Contains(columnName))
            {
                var cellValue = worksheet.Cells[rowNumber, col].Value;
                
                if (cellValue != null)
                {
                    hasData = true;
                    if (columnName == "ID")
                    {
                        if (int.TryParse(cellValue.ToString(), out int id))
                        {
                            newRow["ID"] = id;
                        }
                        else
                        {
                            newRow["ID"] = rowNumber - 1;
                        }
                    }
                    else
                    {
                        newRow[columnName] = cellValue;
                    }
                }
            }
        }

        if (hasData)
        {
            dataTable.Rows.Add(newRow);
        }
    }

    return dataTable;
}
        
        //Ham nay dung de tim row co id lon nhat de id + 1 khi add data moi vao
        private int FindLastRow(ExcelWorksheet worksheet)
        {
            int lastRow = worksheet.Dimension.End.Row;
            while (lastRow >= 1)
            {
                var row = worksheet.Cells[lastRow, 1, lastRow, worksheet.Dimension.End.Column];
                if (row.Any(c => !string.IsNullOrEmpty(c.Text)))
                {
                    break;
                }
                lastRow--;
            }
            return lastRow;
        }

        private void OnSheetSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SheetSelectionComboBox.SelectedItem != null)
            {
                string selectedSheet = SheetSelectionComboBox.SelectedItem.ToString();
                if (_sheetData.ContainsKey(selectedSheet))
                {
                    _dataTable = _sheetData[selectedSheet];
                    UpdateColumnListBox();
                    UpdateDataGridColumns();
                }
            }
        }
        // Ham nay dung de Update ListBox khi open excel, ListBox dung de hien dau diem co trong file excel
        private void UpdateColumnListBox()
        {
            ColumnListBox.Items.Clear();
            foreach (DataColumn column in _dataTable.Columns)
            {
                if (column.ColumnName != "ID" && column.ColumnName != "Role" && column.ColumnName != "Name")
                {
                    var checkBox = new CheckBox
                    {
                        Content = column.ColumnName,
                        IsChecked = false
                    };
                    checkBox.Checked += OnColumnCheckChanged;
                    checkBox.Unchecked += OnColumnCheckChanged;
                    ColumnListBox.Items.Add(checkBox);
                }
            }
        }
        //Hien data theo dau diem check trong list box
        private void OnColumnCheckChanged(object sender, RoutedEventArgs e)
        {
            UpdateDataGridColumns();
        }

        private void UpdateDataGridColumns()
        {
            GradingDataGrid.Columns.Clear();

            // Luôn thêm cột ID, Role và Name
            GradingDataGrid.Columns.Add(new DataGridTextColumn { Header = "ID", Binding = new Binding("ID") { Mode = BindingMode.OneWay } });
            GradingDataGrid.Columns.Add(new DataGridTextColumn { Header = "Role", Binding = new Binding("Role") });
            GradingDataGrid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new Binding("Name") });

            // Thêm các cột đã chọn
            foreach (CheckBox checkBox in ColumnListBox.Items)
            {
                if (checkBox.IsChecked == true)
                {
                    string columnName = checkBox.Content.ToString();
                    GradingDataGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = columnName,
                        Binding = new Binding(columnName)
                    });
                }
            }

            GradingDataGrid.ItemsSource = _dataTable.DefaultView;
        }
        //Add new component
        private void OnAddComponentClick(object sender, RoutedEventArgs e)
        {
            string newComponentName = NewComponentTextBox.Text.Trim();
            if (string.IsNullOrEmpty(newComponentName) || _dataTable.Columns.Contains(newComponentName))
            {
                MessageBox.Show("Component name is invalid or already exists.");
                return;
            }

            _dataTable.Columns.Add(newComponentName);

            var checkBox = new CheckBox
            {
                Content = newComponentName,
                IsChecked = false
            };
            checkBox.Checked += OnColumnCheckChanged;
            checkBox.Unchecked += OnColumnCheckChanged;
            ColumnListBox.Items.Add(checkBox);

            NewComponentTextBox.Clear();
        }
        //Remove Component
        private void OnRemoveComponentClick(object sender, RoutedEventArgs e)
        {
            var selectedCheckBox = ColumnListBox.SelectedItem as CheckBox;
            if (selectedCheckBox != null)
            {
                string columnName = selectedCheckBox.Content.ToString();
                if (_dataTable.Columns.Contains(columnName))
                {
                    _dataTable.Columns.Remove(columnName);
                }
                ColumnListBox.Items.Remove(selectedCheckBox);
                UpdateDataGridColumns();
            }
            else
            {
                MessageBox.Show("No component selected to remove.");
            }
        }
        //Add new Student
        private void OnAddStudentClick(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                MessageBox.Show("No data loaded. Please load an Excel file first.");
                return;
            }

            if (!_dataTable.Columns.Contains("ID") || !_dataTable.Columns.Contains("Role") || !_dataTable.Columns.Contains("Name"))
            {
                MessageBox.Show("Required columns are missing from the data table.");
                return;
            }

            string role = NewStudentRollTextBox.Text.Trim();
            string name = NewStudentNameTextBox.Text.Trim();

            if (string.IsNullOrEmpty(role) || string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Both Role and Name are required.");
                return;
            }

            try
            {
                // Find the current maximum ID
                int maxId = 0;
                if (_dataTable.Rows.Count > 0)
                {
                    maxId = _dataTable.AsEnumerable()
                        .Select(r => int.TryParse(r["ID"].ToString(), out int id) ? id : 0)
                        .Max();
                }

                // Create a new ID by incrementing the maxId
                int newId = maxId + 1;

                // Create a new row
                DataRow newRow = _dataTable.NewRow();
                newRow["ID"] = newId; // Assigning the integer value directly
                newRow["Role"] = role;
                newRow["Name"] = name;

                // Add the new row to the DataTable
                _dataTable.Rows.Add(newRow);

                // Display the new student's information in a MessageBox
                string studentInfo = $"ID: {newId}\nRole: {role}\nName: {name}";
                MessageBox.Show(studentInfo, "New Student Added");

                // Clear the TextBoxes after adding the student
                NewStudentRollTextBox.Clear();
                NewStudentNameTextBox.Clear();

                // Update the DataGrid
                UpdateDataGridColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding new student: {ex.Message}");
            }
        }


        private void OnDeleteSelectedRowsClick(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                MessageBox.Show("No data loaded. Please load an Excel file first.");
                return;
            }

            // Get selected rows from the DataGrid
            var selectedRows = GradingDataGrid.SelectedItems.Cast<DataRowView>().ToList();

            if (selectedRows.Count == 0)
            {
                MessageBox.Show("Please select one or more rows to delete.");
                return;
            }

            // Confirm deletion
            if (MessageBox.Show($"Are you sure you want to delete {selectedRows.Count} selected rows?", "Confirm Deletion", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                try
                {
                    // Delete selected rows from the DataTable
                    foreach (var row in selectedRows)
                    {
                        _dataTable.Rows.Remove(row.Row);
                    }

                    // Update the DataGrid
                    UpdateDataGridColumns();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting rows: {ex.Message}");
                }
            }
        }

        private void OnSaveClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_filePath) || SheetSelectionComboBox.SelectedItem == null) return;

            string selectedSheet = SheetSelectionComboBox.SelectedItem.ToString();

            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                var worksheet = package.Workbook.Worksheets[selectedSheet];
                worksheet.Cells.Clear(); // Clear existing data

                // Write new header
                for (int col = 0; col < _dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = _dataTable.Columns[col].ColumnName;
                }

                // Write new data
                for (int row = 0; row < _dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < _dataTable.Columns.Count; col++)
                    {
                        var value = _dataTable.Rows[row][col];

                        if (value != DBNull.Value && double.TryParse(value.ToString(), out double numericValue))
                        {
                            worksheet.Cells[row + 2, col + 1].Value = numericValue; // Save as a number
                        }
                        else
                        {
                            worksheet.Cells[row + 2, col + 1].Value = value; // Save as text or original value
                        }
                    }
                }

                package.Save();
                _isDataSaved = true;
            }

            MessageBox.Show("Data saved successfully.");
        }


        private void OnSearchRoleClick(object sender, RoutedEventArgs e)
        {
            string searchRole = SearchRoleTextBox.Text.Trim();
            if (string.IsNullOrEmpty(searchRole))
            {
                MessageBox.Show("Please enter a role to search.");
                return;
            }

            DataView dv = _dataTable.DefaultView;
            dv.RowFilter = $"Role LIKE '%{searchRole}%'";
            GradingDataGrid.ItemsSource = dv;
        }

        private void OnClearSearchClick(object sender, RoutedEventArgs e)
        {
            SearchRoleTextBox.Clear();
            _dataTable.DefaultView.RowFilter = string.Empty;
            GradingDataGrid.ItemsSource = _dataTable.DefaultView;
        }

        private void OnSelectAllChecked(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox checkBox in ColumnListBox.Items)
            {
                checkBox.IsChecked = true;
            }
            UpdateDataGridColumns();
        }

        private void OnSelectAllUnchecked(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox checkBox in ColumnListBox.Items)
            {
                checkBox.IsChecked = false;
            }
            UpdateDataGridColumns();
        }
        private void OnExitClick(object sender, RoutedEventArgs e)
        {
            if (!_isDataSaved)
            {
                MessageBoxResult result = MessageBox.Show(
                    "You have unsaved changes. Do you want to save before exiting?",
                    "Unsaved Changes",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Warning);

                switch (result)
                {
                    case MessageBoxResult.Yes:
                        OnSaveClick(null, null);  // Call the save method
                        if (_isDataSaved)  // Check if save was successful
                        {
                            Application.Current.Shutdown();
                        }
                        break;
                    case MessageBoxResult.No:
                        Application.Current.Shutdown();
                        break;
                    case MessageBoxResult.Cancel:
                        // Do nothing, just return to the application
                        break;
                }
            }
            else
            {
                Application.Current.Shutdown();
            }
        }
    }
}