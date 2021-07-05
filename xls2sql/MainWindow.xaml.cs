using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelDataReader;
using System.Configuration;

namespace xls2sql
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

		private void btnOpenFile_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Excel files (*.xls *.xlsx *.csv)|*.xls; *.xlsx; *.csv";
            openFileDialog.Title = "Please select an excel to convert";

            if (openFileDialog.ShowDialog() == true)
            {
                //Get the path of specified file
                var filePath = openFileDialog.FileName;
                txtFilepath.Text = filePath;

                //cmbColors.ItemsSource = tablesNames;

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    var dataSet = excelReader.AsDataSet();
                    var totalTables = dataSet.Tables.Count;

                    string[] tablesNames = new string[totalTables];
                    for (var tableid = 0; tableid < totalTables; tableid++)
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
        }

        private void btnExecute_Click(object sender, RoutedEventArgs e)
        {
            string errors = string.Empty;
            string query = string.Empty;
            string columnNames = string.Empty;
            string columnNamesCreateTable = string.Empty;
            string values = string.Empty;
            
            string databaseName = txtDatabaseName.Text.Trim() != "" ? txtDatabaseName.Text.Trim() : "databasename";
            string tableName = txtTableName.Text.Trim() != "" ? txtTableName.Text.Trim() : "tablename";
            int workbook = cmbWorkbook.SelectedIndex;
            bool? isCreateTable = ckbCreateTable.IsChecked;
            var filePath = txtFilepath.Text;

            //validate inputs
            if (filePath == null || filePath == "" || filePath == string.Empty)
                errors += "File is not selected.\n";
            if (workbook == -1 && filePath != "")
                errors += "Workbook is not selected.\n";

            if (errors == "")
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    var dataSet = excelReader.AsDataSet();

                    var totalColumns = dataSet.Tables[workbook].Columns.Count;
                    var totalRows = dataSet.Tables[workbook].Rows.Count;

                    for (var i = 0; i < totalRows; i++)
                    {
                        if (i != 0)
                            values += "(";
                        for (var j = 0; j < totalColumns; j++)
                        {
                            var value = dataSet.Tables[workbook].Rows[i][j].ToString().Replace("'", "''");

                            //1st row contains column names so let's store it separatelly
                            //else generate data
                            if (i == 0)
                            {
                                //if we are not on last column, use separator
                                //else dont put , at the end
                                if (j != totalColumns - 1)
                                {
                                    columnNames += "[" + value + "], ";
                                    columnNamesCreateTable += "[" + value + "] varchar(max) NULL, ";
                                }
                                else
                                { 
                                    columnNames += "[" + value + "]";
                                    columnNamesCreateTable += "[" + value + "] varchar(max) NULL";
                                }
                            }
                            else
                            {
                                //if we are not on last column, use separator
                                //else dont put , at the end
                                if (j != totalColumns - 1)
                                    values += "N'" + value + "', ";
                                else
                                    values += "N'" + value + "'";
                            }
                        }
                        //if we are not on last column, use separator
                        //else dont put , at the end
                        if (i != totalRows - 1 && i != 0)
                            values += "), ";
                        else if (i == totalRows - 1 && i != 0)
                            values += ")";
                    }
                }
                query += $"USE {databaseName};\nSET ANSI_NULLS ON;\nSET QUOTED_IDENTIFIER ON;\n\n";
                if ((bool)isCreateTable)
                {
                    query += $"CREATE TABLE {tableName} ({columnNamesCreateTable});\n\n";
                }
                query += $"INSERT INTO {tableName} ({columnNames})\nVALUES {values};";
                txtEditor.Text = query;
            }
            else
                MessageBox.Show(errors, "xls2sql", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        }
    }
}
