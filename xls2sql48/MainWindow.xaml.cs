using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace xls2sql48
{
    public partial class MainWindow : Window
    {
        private static readonly Regex _nonDigits = new Regex("[^0-9]+", RegexOptions.Compiled);
        private static readonly Regex _lineBreaks = new Regex(@"\t|\n|\r", RegexOptions.Compiled);
        private static readonly ExcelReaderConfiguration _readerConfig = new ExcelReaderConfiguration
        {
            FallbackEncoding = Encoding.GetEncoding(1250),
            AutodetectSeparators = new char[] { ',', ';', '\t' },
            LeaveOpen = false,
        };

        private DataSet _cachedDataSet;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            e.Handled = _nonDigits.IsMatch(e.Text);
        }

        private void LoadFile(string filePath)
        {
            var timer = Stopwatch.StartNew();
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                using (var excelReader = ExcelReaderFactory.CreateReader(stream, _readerConfig))
                {
                    _cachedDataSet = excelReader.AsDataSet();
                }

                int totalTables = _cachedDataSet.Tables.Count;
                string[] tableNames = new string[totalTables];
                for (int i = 0; i < totalTables; i++)
                    tableNames[i] = _cachedDataSet.Tables[i].TableName;

                cmbWorkbook.ItemsSource = tableNames;
                if (totalTables > 1)
                {
                    cmbWorkbook.Visibility = Visibility.Visible;
                    cmbWorkbook.SelectedIndex = -1;
                }
                else
                {
                    cmbWorkbook.Visibility = Visibility.Collapsed;
                    cmbWorkbook.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "xls2sql", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }

            txtStatus.Text = $"File Loading Time: {timer.ElapsedMilliseconds}ms";
        }

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xls *.xlsx *.csv)|*.xls; *.xlsx; *.csv",
                Title = "Please select an excel to convert"
            };

            if (dialog.ShowDialog() == true)
            {
                txtFilepath.Text = dialog.FileName;
                LoadFile(dialog.FileName);
            }
        }

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            string errors = ValidateInputs();
            if (!string.IsNullOrEmpty(errors))
            {
                MessageBox.Show(errors, "xls2sql", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            SetButtonsEnabled(false);
            var timer = Stopwatch.StartNew();

            SqlSettings settings = ReadSettings();
            string query = await Task.Run(() => GenerateSQLQuery(settings));

            txtEditor.Text = query;
            txtStatus.Text = $"Execution Time: {timer.ElapsedMilliseconds}ms";
            SetButtonsEnabled(true);
        }

        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            string errors = ValidateInputs();
            if (!string.IsNullOrEmpty(errors))
            {
                MessageBox.Show(errors, "xls2sql", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            SetButtonsEnabled(false);
            var timer = Stopwatch.StartNew();

            SqlSettings settings = ReadSettings();
            string query = await Task.Run(() => GenerateSQLQuery(settings));

            string fileName = string.IsNullOrEmpty(settings.TableName) ? "tablename.txt" : settings.TableName + ".txt";
            string filePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
            await Task.Run(() => File.WriteAllText(filePath, query));

            txtStatus.Text = $"Saved to {fileName} | Execution Time: {timer.ElapsedMilliseconds}ms";
            SetButtonsEnabled(true);
        }

        private void SetButtonsEnabled(bool enabled)
        {
            btnExecute.IsEnabled = enabled;
            btnSave.IsEnabled = enabled;
        }

        private SqlSettings ReadSettings() => new SqlSettings
        {
            DatabaseName = string.IsNullOrEmpty(txtDatabaseName.Text.Trim()) ? "databasename" : txtDatabaseName.Text.Trim(),
            TableName = string.IsNullOrEmpty(txtTableName.Text.Trim()) ? "tablename" : txtTableName.Text.Trim(),
            IsCreateTable = ckbCreateTable.IsChecked ?? false,
            Separator = int.TryParse(txtSeparator.Text, out int sep) && sep >= 1 ? sep : 1000,
            PreferNulls = ckbPrefferNulls.IsChecked ?? false,
            TrimWhiteSpaces = ckbTrimWhiteSpaces.IsChecked ?? false,
            FirstColumnId = cmbFirstColumn.SelectedIndex,
            DataTable = _cachedDataSet.Tables[cmbWorkbook.SelectedIndex],
        };

        private string ValidateInputs()
        {
            string errors = "";
            if (string.IsNullOrEmpty(txtFilepath.Text.Trim()))
                errors += "File is not selected.\n";
            if (cmbWorkbook.SelectedIndex == -1)
                errors += "Workbook is not selected.\n";
            return errors;
        }

        private static string GenerateSQLQuery(SqlSettings s)
        {
            var queryBuilder = new StringBuilder();
            queryBuilder.AppendLine($"USE {s.DatabaseName};");
            queryBuilder.AppendLine("SET ANSI_NULLS ON;");
            queryBuilder.AppendLine("SET QUOTED_IDENTIFIER ON;");
            queryBuilder.AppendLine();

            try
            {
                List<string> headers = GetColumnNames(s.DataTable, s.TrimWhiteSpaces);
                string columnNames = GenerateColumnNames(headers);
                int headerCount = headers.Count;

                if (s.IsCreateTable)
                {
                    queryBuilder.AppendLine($"CREATE TABLE {s.TableName} ({GenerateColumnNamesForCreateTable(headers, s.FirstColumnId)});");
                    queryBuilder.AppendLine();
                }

                var valuesBuilder = new StringBuilder();
                valuesBuilder.Append($"INSERT INTO {s.TableName} ({columnNames}) ");
                valuesBuilder.AppendLine();
                valuesBuilder.Append("VALUES ");

                var rowValues = new string[headerCount];
                bool firstInBatch = true;
                bool headerSkipped = false;
                int dataRowCount = 0;

                foreach (DataRow row in s.DataTable.Rows)
                {
                    if (!headerSkipped) { headerSkipped = true; continue; }

                    object[] cells = row.ItemArray;
                    int colCount = Math.Min(headerCount, cells.Length);

                    for (int j = 0; j < colCount; j++)
                    {
                        string value = cells[j].ToString().Replace("'", "''");
                        if (s.TrimWhiteSpaces) value = value.Trim();
                        if (value.Length == 0 && s.PreferNulls) value = "NULL";
                        rowValues[j] = value;
                    }

                    string rowString = $"('{string.Join("', N'", rowValues, 0, colCount)}')";
                    if (s.PreferNulls) rowString = rowString.Replace("N'NULL'", "NULL");

                    if (dataRowCount > 0 && dataRowCount % s.Separator == 0)
                    {
                        valuesBuilder.AppendLine(";");
                        valuesBuilder.AppendLine();
                        valuesBuilder.Append($"INSERT INTO {s.TableName} ({columnNames}) ");
                        valuesBuilder.AppendLine();
                        valuesBuilder.Append("VALUES ");
                        firstInBatch = true;
                    }

                    if (!firstInBatch) valuesBuilder.Append(", ");
                    valuesBuilder.Append(rowString);
                    firstInBatch = false;
                    dataRowCount++;
                }

                valuesBuilder.Append(";");
                return queryBuilder.Append(valuesBuilder).ToString();
            }
            catch
            {
                Application.Current.Dispatcher.Invoke(() =>
                    MessageBox.Show("Error Occurred", "Error", MessageBoxButton.OK, MessageBoxImage.Warning));
                return string.Empty;
            }
        }

        private static List<string> GetColumnNames(DataTable dataTable, bool trimWhiteSpaces)
        {
            var headers = new List<string>();
            if (dataTable.Rows.Count == 0) return headers;

            foreach (object column in dataTable.Rows[0].ItemArray)
            {
                string header = Convert.ToString(column);
                if (trimWhiteSpaces) header = header.Trim();
                header = _lineBreaks.Replace(header, "");
                if (!string.IsNullOrEmpty(header))
                    headers.Add(header);
            }
            return headers;
        }

        private static string GenerateColumnNames(List<string> headers)
        {
            if (headers == null || headers.Count == 0) return string.Empty;
            return $"[{string.Join("], [", headers)}]";
        }

        private static string GenerateColumnNamesForCreateTable(List<string> headers, int firstColumnId = 0)
        {
            if (headers == null || headers.Count == 0) return string.Empty;

            var sb = new StringBuilder();
            if (firstColumnId == 1)
                sb.Append("Id INT IDENTITY (1, 1) NOT NULL, ");
            else if (firstColumnId == 2)
                sb.Append("Id UNIQUEIDENTIFIER DEFAULT NEWSEQUENTIALID() NOT NULL, ");

            sb.Append($"[{string.Join("] varchar(max) NULL, [", headers)}] varchar(max) NULL");
            return sb.ToString();
        }

        private void TxtFilepath_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void TxtFilepath_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length != 1) return;

            string filePath = files[0];
            if (!filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) &&
                !filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                !filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Only Excel files (.xls, .xlsx, .csv) are supported.", "Invalid File", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            txtFilepath.Text = filePath;
            LoadFile(filePath);
        }
    }

    internal class SqlSettings
    {
        public string DatabaseName { get; set; }
        public string TableName { get; set; }
        public bool IsCreateTable { get; set; }
        public int Separator { get; set; }
        public bool PreferNulls { get; set; }
        public bool TrimWhiteSpaces { get; set; }
        public int FirstColumnId { get; set; }
        public DataTable DataTable { get; set; }
    }
}
