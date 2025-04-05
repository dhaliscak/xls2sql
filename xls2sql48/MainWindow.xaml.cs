using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace xls2sql48
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xls *.xlsx *.csv)|*.xls; *.xlsx; *.csv";
            openFileDialog.Title = "Please select an excel to convert";

            if (openFileDialog.ShowDialog() == true)
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();

                var filePath = openFileDialog.FileName;
                txtFilepath.Text = filePath;

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                var readerConfig = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1250),
                    AutodetectSeparators = new char[] { ',', ';', '\t' },
                    LeaveOpen = false,
                };

                try
                {
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream, readerConfig))
                    {
                        DataSet dataSet = excelReader.AsDataSet();
                        int totalTables = dataSet.Tables.Count;

                        string[] tablesNames = new string[totalTables];
                        for (int tableid = 0; tableid < totalTables; tableid++)
                        {
                            tablesNames[tableid] = dataSet.Tables[tableid].TableName;
                        }

                        if (totalTables > 1)
                        {
                            cmbWorkbook.ItemsSource = tablesNames;
                            cmbWorkbook.Visibility = Visibility.Visible;
                            cmbWorkbook.SelectedIndex = -1;
                        }
                        else
                        {
                            cmbWorkbook.ItemsSource = tablesNames;
                            cmbWorkbook.Visibility = Visibility.Collapsed;
                            cmbWorkbook.SelectedIndex = 0;
                        }
                    }
                }
                catch (Exception er)
                {
                    ShowError(er);
                }

                timer.Stop();
                txtStatus.Text = $"File Loading Time: {timer.ElapsedMilliseconds}ms";
            }
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            string errors = ValidateInputs();

            if (string.IsNullOrEmpty(errors))
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();

                string query = GenerateSQLQuery();
                txtEditor.Text = query;

                timer.Stop();
                txtStatus.Text = $"Execution Time: {timer.ElapsedMilliseconds}ms";
            }
            else
            {
                MessageBox.Show(errors, "xls2sql", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        private string ValidateInputs()
        {
            string errors = "";
            string filePath = txtFilepath.Text.Trim();
            int workbookIndex = cmbWorkbook.SelectedIndex;

            if (string.IsNullOrEmpty(filePath))
            {
                errors += "File is not selected.\n";
            }

            if (workbookIndex == -1)
            {
                errors += "Workbook is not selected.\n";
            }

            return errors;
        }

        private void ShowError(Exception e)
        {
            MessageBox.Show(e.Message, "xls2sql", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        }

        private string GenerateSQLQuery()
        {
            string databaseName = string.IsNullOrEmpty(txtDatabaseName.Text.Trim()) ? "databasename" : txtDatabaseName.Text.Trim();
            string tableName = string.IsNullOrEmpty(txtTableName.Text.Trim()) ? "tablename" : txtTableName.Text.Trim();
            bool isCreateTable = ckbCreateTable.IsChecked ?? false;
            string filePath = txtFilepath.Text.Trim();

            int separator = (txtSeparator.Text == "" || Int32.Parse(txtSeparator.Text) < 1) ? 1000 : Int32.Parse(txtSeparator.Text);
            bool prefferNulls = ckbPrefferNulls.IsChecked ?? false;
            bool trimWhiteSpaces = ckbTrimWhiteSpaces.IsChecked ?? false;
            int workbookIndex = cmbWorkbook.SelectedIndex;
            int firstColumnId = cmbFirstColumn.SelectedIndex;

            StringBuilder queryBuilder = new StringBuilder();
            queryBuilder.AppendLine($"USE {databaseName};");
            queryBuilder.AppendLine("SET ANSI_NULLS ON;");
            queryBuilder.AppendLine("SET QUOTED_IDENTIFIER ON;");
            queryBuilder.AppendLine();

            var readerConfig = new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding(1250),
                AutodetectSeparators = new char[] { ',', ';', '\t' },
                LeaveOpen = false,
            };

            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream, readerConfig))
                {
                    var dataSet = excelReader.AsDataSet();
                    var dataTable = dataSet.Tables[workbookIndex];

                    List<string> headers = GetColumnNames(dataTable, trimWhiteSpaces);
                    string columnNames = GenerateColumnNames(headers);

                    if (isCreateTable)
                    {
                        string columnNamesCreateTable = GenerateColumnNamesForCreateTable(headers, firstColumnId);
                        queryBuilder.AppendLine($"CREATE TABLE {tableName} ({columnNamesCreateTable});");
                        queryBuilder.AppendLine();
                    }

                    var index = 0;
                    List<object> rowData = null;
                    List<string> tableData = new List<string>();

                    foreach (DataRow row in dataTable.Rows)
                    {
                        List<string> queryRowData = new List<string>();
                        rowData = row.ItemArray.ToList();

                        if (index > 0)
                        {
                            var j = 0;
                            foreach (var data in rowData)
                            {
                                if (j >= headers.Count())
                                {
                                    break;
                                }                                    

                                string value = data.ToString().Replace("'", "''");
                                value = trimWhiteSpaces ? value.Trim() : value;
                                value = value == "" && prefferNulls ? "NULL" : value;
                                queryRowData.Add(value);

                                j++;
                            }

                            var rowString = $"('{String.Join("', N'", queryRowData)}')";
                            rowString = prefferNulls ? rowString.ToString().Replace("N'NULL'", "NULL") : rowString;

                            tableData.Add(rowString);
                        }
                        index++;
                    }

                    StringBuilder valuesBuilder = new StringBuilder();
                    List<string> query = new List<string>();
                    valuesBuilder.Append($"INSERT INTO {tableName} ({columnNames}) ");
                    valuesBuilder.AppendLine();
                    valuesBuilder.Append($"VALUES ");

                    for (var i = 0; i < tableData.Count; i++)
                    {
                        if (i > 0 && i % separator == 0)
                        {
                            valuesBuilder.Append($"{String.Join(", ", query)}; ");
                            valuesBuilder.AppendLine();
                            query.Clear();
                            valuesBuilder.AppendLine();
                            valuesBuilder.Append($"INSERT INTO {tableName} ({columnNames}) ");
                            valuesBuilder.AppendLine();
                            valuesBuilder.Append($"VALUES ");
                        }
                        query.Add(tableData[i]);
                    }
                    valuesBuilder.Append($"{String.Join(", ", query)}; ");

                    return $"{queryBuilder}{valuesBuilder}";
                }
            }
            catch
            {
                MessageBox.Show("Error Occured", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return $"";
            }
        }

        private List<string> GetColumnNames(DataTable dataTable, bool trimWhiteSpaces)
        {
            var headers = new List<string>();

            if (dataTable.Rows.Count > 0)
            {
                DataRow firstRow = dataTable.Rows[0];

                foreach (var column in firstRow.ItemArray)
                {
                    string header = Convert.ToString(column);
                    if (trimWhiteSpaces)
                    {
                        header = header.Trim();
                    }
                    header = Regex.Replace(header, @"\t|\n|\r", "");

                    // Add header only if it's not empty
                    if (!string.IsNullOrEmpty(header))
                    {
                        headers.Add(header);
                    }
                }
            }

            return headers;
        }

        private string GenerateColumnNames(List<string> headers)
        {
            if (headers == null || headers.Count == 0)
            {
                return string.Empty;
            }

            StringBuilder columnNamesBuilder = new StringBuilder();
            columnNamesBuilder.Append($"[{string.Join("], [", headers)}]");

            return columnNamesBuilder.ToString();
        }

        private string GenerateColumnNamesForCreateTable(List<string> headers, int firstColumnId = 0)
        {
            if (headers == null || headers.Count == 0)
            {
                return string.Empty;
            }

            StringBuilder columnNamesBuilder = new StringBuilder();

            switch (firstColumnId)
            {
                case 1:
                    columnNamesBuilder.Append("Id INT IDENTITY (1, 1) NOT NULL");
                    break;
                case 2:
                    columnNamesBuilder.Append("Id UNIQUEIDENTIFIER DEFAULT NEWSEQUENTIALID() NOT NULL");
                    break;
            }

            if (firstColumnId != 0)
            {
                columnNamesBuilder.Append(", ");
            }

            columnNamesBuilder.Append($"[{string.Join("] varchar(max) NULL, [", headers)}] varchar(max) NULL");
            return columnNamesBuilder.ToString();
        }

        private void TxtFilepath_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void TxtFilepath_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1)
                {
                    Stopwatch timer = new Stopwatch();

                    string filePath = files[0];

                    // Ensure the file is an Excel file before setting the path
                    if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                        filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                        filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        txtFilepath.Text = filePath;

                        timer.Start();

                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                        var readerConfig = new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding(1250),
                            AutodetectSeparators = new char[] { ',', ';', '\t' },
                            LeaveOpen = false,
                        };

                        try
                        {
                            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream, readerConfig))
                            {
                                DataSet dataSet = excelReader.AsDataSet();
                                int totalTables = dataSet.Tables.Count;

                                string[] tablesNames = new string[totalTables];
                                for (int tableid = 0; tableid < totalTables; tableid++)
                                {
                                    tablesNames[tableid] = dataSet.Tables[tableid].TableName;
                                }

                                if (totalTables > 1)
                                {
                                    cmbWorkbook.ItemsSource = tablesNames;
                                    cmbWorkbook.Visibility = Visibility.Visible;
                                    cmbWorkbook.SelectedIndex = -1;
                                }
                                else
                                {
                                    cmbWorkbook.ItemsSource = tablesNames;
                                    cmbWorkbook.Visibility = Visibility.Collapsed;
                                    cmbWorkbook.SelectedIndex = 0;
                                }
                            }
                        }
                        catch (Exception er)
                        {
                            ShowError(er);
                        }

                        timer.Stop();
                        txtStatus.Text = $"File Loading Time: {timer.ElapsedMilliseconds}ms";

                    }
                    else
                    {
                        MessageBox.Show("Only Excel files (.xls, .xlsx, .csv) are supported.", "Invalid File", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
        }
    }
}
