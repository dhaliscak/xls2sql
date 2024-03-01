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

                // Configure the reader to treat the first row as headers
                var readerConfig = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1250),
                    AutodetectSeparators = new char[] { ',', ';', '\t' },
                    LeaveOpen = true,
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

            int separator = txtSeparator.Text == "" ? 1000 : Int32.Parse(txtSeparator.Text);
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
                LeaveOpen = true,
            };

            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream, readerConfig))
            {
                var dataSet = excelReader.AsDataSet();
                var dataTable = dataSet.Tables[workbookIndex];

                List<string> headers = GetColumnNames(excelReader);
                string columnNames = GenerateColumnNames(headers, trimWhiteSpaces);

                if (isCreateTable)
                {
                    string columnNamesCreateTable = GenerateColumnNamesForCreateTable(headers, trimWhiteSpaces, firstColumnId);
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
                        foreach (var data in rowData)
                        {
                            string value = data.ToString().Replace("'", "''");
                            value = trimWhiteSpaces ? value.Trim() : value;
                            queryRowData.Add(value);
                        }

                        var temp = $"('{String.Join("', N'", queryRowData)}')";
                        if (prefferNulls)
                        {
                            temp = temp.ToString().Replace("N'NULL'", "NULL").Replace("N''", "NULL");
                        }
                        tableData.Add(temp);
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
                        valuesBuilder.Append($"{String.Join("', ", query)}");
                        query.Clear();
                        valuesBuilder.AppendLine();
                        valuesBuilder.Append($"INSERT INTO {tableName} ({columnNames}) ");
                        valuesBuilder.AppendLine();
                        valuesBuilder.Append($"VALUES ");

                    }
                    query.Add(tableData[i]);
                }
                valuesBuilder.Append($"{String.Join(", ", query)} ");

                return $"{queryBuilder.ToString()}{valuesBuilder}";
            }

        }

        private List<string> GetColumnNames(IExcelDataReader excelReader)
        {
            var headers = new List<string>();

            if (excelReader.Read())
            {
                for (var i = 0; i < excelReader.FieldCount; i++)
                {
                    headers.Add(Regex.Replace(Convert.ToString(excelReader[i]), @"\t|\n|\r", ""));
                }
            }

            return headers;
        }

        private string GenerateColumnNames(List<string> headers, bool trimWhiteSpaces)
        {

            StringBuilder columnNamesBuilder = new StringBuilder();
            
            foreach (var column in headers)
            {
                string columnName = trimWhiteSpaces ? column.Trim() : column;
                columnNamesBuilder.Append($"[{columnName}], ");
            }

            return columnNamesBuilder.ToString().TrimEnd(' ', ',');
        }

        private string GenerateColumnNamesForCreateTable(List<string> headers, bool trimWhiteSpaces, int firstColumnId = 0)
        {
            StringBuilder columnNamesBuilder = new StringBuilder();

            switch (firstColumnId)
            {
                case 0:
                    break;
                case 1:
                    columnNamesBuilder.Append($"Id INT IDENTITY (1, 1) NOT NULL, ");
                    break;
                case 2:
                    columnNamesBuilder.Append($"Id UNIQUEIDENTIFIER DEFAULT NEWSEQUENTIALID() NOT NULL, ");
                    break;
            }

            foreach (var column in headers)
            {
                string columnName = trimWhiteSpaces ? column.Trim() : column;
                columnNamesBuilder.Append($"[{columnName}] varchar(max) NULL, ");
            }

            return columnNamesBuilder.ToString().TrimEnd(' ', ',');
        }
    }
}
