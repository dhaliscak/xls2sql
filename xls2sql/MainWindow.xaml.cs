using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace xls2sql
{
    public partial class MainWindow : Window
    {
        private static readonly Regex _nonDigits = new Regex("[^0-9]+", RegexOptions.Compiled);

        private string _loadedFilePath;
        private string[] _sheetNames;
        private SqlOutput _output;

        // Parsed sheets of the loaded file, keyed by sheet name. Only touched on the UI thread.
        // Each sheet is read once in the background (on load / selection) and reused by every Generate/Save.
        private readonly Dictionary<string, Task<SheetData>> _sheetCache = new Dictionary<string, Task<SheetData>>();

        public MainWindow()
        {
            InitializeComponent();

            SourceInitialized += (s, e) => ThemeManager.ApplyTitleBar(this);
            ThemeManager.ThemeChanged += OnThemeChanged;
            Closed += (s, e) => ThemeManager.ThemeChanged -= OnThemeChanged;
            UpdateThemeButton();
        }

        private void BtnTheme_Click(object sender, RoutedEventArgs e) => ThemeManager.Cycle();

        private void OnThemeChanged(object sender, EventArgs e)
        {
            ThemeManager.ApplyTitleBar(this);
            UpdateThemeButton();
        }

        // Segoe MDL2 / Fluent glyphs: E7F4 monitor (System), E706 sun (Light), E708 moon (Dark).
        private void UpdateThemeButton()
        {
            switch (ThemeManager.Mode)
            {
                case ThemeMode.Light:
                    btnTheme.Content = "";
                    btnTheme.ToolTip = "Theme: Light (click for Dark)";
                    break;
                case ThemeMode.Dark:
                    btnTheme.Content = "";
                    btnTheme.ToolTip = "Theme: Dark (click to follow Windows)";
                    break;
                default:
                    btnTheme.Content = "";
                    btnTheme.ToolTip = $"Theme: Windows ({(ThemeManager.IsDark ? "Dark" : "Light")}) (click for Light)";
                    break;
            }
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
                string[] tableNames = SheetReader.ReadSheetNames(filePath);
                int totalTables = tableNames.Length;
                _loadedFilePath = filePath;
                _sheetNames = tableNames;
                _sheetCache.Clear();

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

        private void CmbWorkbook_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Start parsing the selected sheet right away so Generate/Save only has to format.
            if (_sheetNames != null && cmbWorkbook.SelectedIndex >= 0)
                GetSheetData(_sheetNames[cmbWorkbook.SelectedIndex]);
        }

        private Task<SheetData> GetSheetData(string sheetName)
        {
            if (!_sheetCache.TryGetValue(sheetName, out Task<SheetData> task) || task.IsFaulted || task.IsCanceled)
            {
                string filePath = _loadedFilePath;
                task = Task.Run(() => SheetReader.Load(filePath, sheetName));
                _sheetCache[sheetName] = task;
            }
            return task;
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

            // Drop the previous result first so old and new output are never alive together.
            _output = null;
            txtEditor.Clear();
            btnCopyAll.Visibility = Visibility.Collapsed;

            SqlSettings settings = ReadSettings();
            SqlOutput output = await GenerateSQLQueryAsync(settings);
            OutputIndex index = await Task.Run(() => OutputIndex.Build(output));

            _output = output;
            txtEditor.SetOutput(index);
            btnCopyAll.Visibility = output.Length > 0 ? Visibility.Visible : Visibility.Collapsed;

            txtStatus.Text = $"Execution Time: {timer.ElapsedMilliseconds}ms";
            SetButtonsEnabled(true);

            ReleaseFreedMemory();
        }

        // Generation leaves large temporary buffers on the large object heap, which the GC neither
        // compacts nor returns to Windows by default; compact once so Task Manager reflects real use.
        private static void ReleaseFreedMemory()
        {
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
#if NET7_0_OR_GREATER
            // Aggressive also decommits the freed regions (~0.5 GB less private memory on big files).
            GC.Collect(2, GCCollectionMode.Aggressive, blocking: true, compacting: true);
#else
            GC.Collect();
#endif
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
            SqlOutput output = await GenerateSQLQueryAsync(settings);

            string fileName = string.IsNullOrEmpty(settings.TableName) ? "tablename.txt" : settings.TableName + ".txt";
            string filePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
            // Same bytes as File.WriteAllText (UTF-8, no BOM), streamed chunk by chunk.
            await Task.Run(() =>
            {
                using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16))
                    output.WriteTo(stream);
            });

            txtStatus.Text = $"Saved to {fileName} | Execution Time: {timer.ElapsedMilliseconds}ms";
            SetButtonsEnabled(true);

            output = null;
            ReleaseFreedMemory();
        }

        private void BtnCopyAll_Click(object sender, RoutedEventArgs e)
        {
            if (_output == null) return;
            try
            {
                string text = _output.ToText();
                // copy: true keeps the text on the clipboard after the app exits.
                Clipboard.SetDataObject(text, true);
                txtStatus.Text = $"Copied {text.Length} characters";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not copy to clipboard: {ex.Message}", "xls2sql", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async Task<SqlOutput> GenerateSQLQueryAsync(SqlSettings settings)
        {
            try
            {
                SheetData data = await GetSheetData(settings.SheetName);
                return await Task.Run(() => SqlGenerator.Generate(settings, data));
            }
            catch
            {
                MessageBox.Show("Error Occurred", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return SqlOutput.Empty;
            }
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
            FilePath = _loadedFilePath,
            SheetName = _sheetNames[cmbWorkbook.SelectedIndex],
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
}
