using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using System.IO;
using System.Windows;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Controls;

namespace ExcelToHtmlConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            Logs = new ObservableCollection<Log>();
            FormulaCells = new ObservableCollection<Cell>();
            InitializeComponent();
            DataContext = this;
        }

        public ObservableCollection<Log> Logs { get; set; }
        public ObservableCollection<Cell> FormulaCells { get; set; }

        private async void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel file|*.xlsx",
                Multiselect = false
            };
            if (openFileDialog.ShowDialog() == true)
            {
                txtFilePath.Text = openFileDialog.FileName;
                btnGenerateHtml.IsEnabled = true;
            }
            LogInfo($"Selected excel file is {txtFilePath.Text}");
            await LoadFormulaCells(txtFilePath.Text);
        }

        private async Task LoadFormulaCells(string sourceFilePath)
        {
            FormulaCells.Clear();
            InitializeExcel();
            LogInfo($"Opening file {sourceFilePath}");
            var workbook = workbooks.Open(sourceFilePath);
            try
            {
                LogInfo($"Looping through all worksheets for formula cells");
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    LogInfo($"Searching for formula cells in Sheet '{sheet.Name}'");
                    foreach (Excel.Range cell in sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, 23))
                    {
                        var formulaCell = new Cell { Formula = cell.Formula.ToString(), Address = cell.Address };
                        FormulaCells.Add(formulaCell);
                        LogInfo($"Loaded formula cell {formulaCell}");
                    }
                    LogInfo($"Finished loading formula cells in Sheet '{sheet.Name}'");
                }
                LogInfo($"Successfully loaded all formula cells!!");
                tabFormulas.Header = $"Formulas ({FormulaCells.Count})";
            }
            catch (Exception ex)
            {
                LogError($"Error occured while loading excel file formula cells. {ex}");
            }
            finally
            {
                workbook.Close(false);
                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);
                await Task.Delay(1000);
                btnGenerateHtml.IsEnabled = true;
            }
        }

        private void SelectOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            var selectFolderDialog = new VistaFolderBrowserDialog
            {
                SelectedPath = Constants.DefaultOutputFolder
            };
            var result = selectFolderDialog.ShowDialog();
            if (result.HasValue && result.Value == true)
            {
                txtOutputPath.Text = selectFolderDialog.SelectedPath;
            }
            LogInfo($"Output folder is {txtOutputPath.Text}");
        }

        Excel.Application excel = null;
        Excel.Workbooks workbooks = null;

        private void InitializeExcel()
        {
            if (excel == null)
            {
                LogInfo($"Started Excel Process");
                excel = new Excel.Application();
                workbooks = excel.Workbooks;
            }
        }

        private void DisposeExcel()
        {
            LogInfo($"Quitting Excel Process");
            excel?.Quit();
            if (workbooks != null)
                Marshal.ReleaseComObject(workbooks);
            if (excel != null)
                Marshal.ReleaseComObject(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private async void GenerateHtmlClick(object sender, RoutedEventArgs e)
        {
            try
            {
                LogInfo($"Started generating html from the selected excel file.");
                btnGenerateHtml.IsEnabled = false;
                var sourceFilePath = txtFilePath.Text;
                LogInfo($"Opening excel file {sourceFilePath}");
                InitializeExcel();
                var workbook = workbooks.Open(sourceFilePath);
                try
                {
                    var outputFolder = Directory.CreateDirectory(txtOutputPath.Text).FullName;
                    var outputFileName = $"{Path.GetFileNameWithoutExtension(sourceFilePath)}.{DateTime.Now.Ticks}.html";
                    var outputPath = $"{outputFolder}\\{outputFileName}";
                    LogInfo($"Generating html files at {outputPath}");
                    await Task.Run(() => workbook.SaveAs(outputPath, Excel.XlFileFormat.xlHtml));
                    LogSuccess($"Success!!");
                }
                finally
                {
                    workbook.Close(false);
                    if (workbook != null)
                        Marshal.ReleaseComObject(workbook);
                    await Task.Delay(1000);
                    btnGenerateHtml.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                LogError($"Error occured while generating html file.{ex}");
            }
        }

        private void LogInfo(string data)
        {
            Logs.Add(new Log
            {
                Text = data,
                Type = LogType.Info
            });
        }

        private void LogError(string data)
        {
            Logs.Add(new Log
            {
                Text = data,
                Type = LogType.Error
            });
        }

        private void LogSuccess(string data)
        {
            Logs.Add(new Log
            {
                Text = data,
                Type = LogType.Success
            });
        }

        private void Window_SourceInitialized(object sender, EventArgs e)
        {
            this.Top = Properties.Settings.Default.Top;
            this.Left = Properties.Settings.Default.Left;
            this.Height = Properties.Settings.Default.Height;
            this.Width = Properties.Settings.Default.Width;
            // Very quick and dirty - but it does the job
            if (Properties.Settings.Default.Maximized)
            {
                WindowState = WindowState.Maximized;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                // Use the RestoreBounds as the current values will be 0, 0 and the size of the screen
                Properties.Settings.Default.Top = RestoreBounds.Top;
                Properties.Settings.Default.Left = RestoreBounds.Left;
                Properties.Settings.Default.Height = RestoreBounds.Height;
                Properties.Settings.Default.Width = RestoreBounds.Width;
                Properties.Settings.Default.Maximized = true;
            }
            else
            {
                Properties.Settings.Default.Top = this.Top;
                Properties.Settings.Default.Left = this.Left;
                Properties.Settings.Default.Height = this.Height;
                Properties.Settings.Default.Width = this.Width;
                Properties.Settings.Default.Maximized = false;
            }
            DisposeExcel();
            Properties.Settings.Default.Save();
        }

    }

    public class Cell
    {
        public string Address
        {
            get;
            set;
        }

        public string Formula
        {
            get;
            set;
        }
        public override string ToString()
        {
            return $"Address:{Address},Formula:{Formula}";
        }
    }

    public class Log
    {
        public Log()
        {
            Logged = DateTime.Now;
        }

        public DateTime Logged
        {
            get; set;
        }

        public string Text
        {
            get;
            set;
        }

        public LogType Type
        {
            get;
            set;
        }
    }

    public enum LogType
    {
        Info,
        Error,
        Success
    }

    public static class Constants
    {
        public const string DefaultTxtBoxOutputFolder = DefaultOutputFolder;
        public const string DefaultTxtBoxExcelFile = "Please select a excel file";
        public const string SelectExcelFileButtonText = "Select Excel File";
        public const string SelectOutputFolderButtonText = " Output Folder";
        public const string DefaultOutputFolder = @"C:\Temp\ExcelConvertor\Out";
        public const string LogsTabHeaderText = "Logs";
        public const string GenerateHtmlButtonText = "Generate HTML";
    }
}
