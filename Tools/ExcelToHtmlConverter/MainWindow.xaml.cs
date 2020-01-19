using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using System.IO;
using System.Windows;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter
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

        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel file|*.xlsx";
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == true)
            {
                txtFilePath.Text = openFileDialog.FileName;
                btnGenerateHtml.IsEnabled = true;
            }
        }

        private void SelectOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new VistaFolderBrowserDialog();
            openFileDialog.SelectedPath = Constants.DefaultOutputFolder;
            openFileDialog.ShowDialog();
        }

        private async void GenerateHtmlClick(object sender, RoutedEventArgs e)
        {
            var sourceFilePath = txtFilePath.Text;
            var excel = new Excel.Application();
            var workbooks = excel.Workbooks;
            var workbook = workbooks.Open(sourceFilePath);
            try
            {
                var outputFolder = Directory.CreateDirectory(txtOutputPath.Text).FullName;
                var outputFileName = $"{Path.GetFileNameWithoutExtension(sourceFilePath)}.{DateTime.Now.Ticks}.html";
                var outputPath = $"{outputFolder}\\{outputFileName}";
                await Task.Run(() => workbook.SaveAs(outputPath, Excel.XlFileFormat.xlHtml));
            }
            finally
            {
                excel.Quit();
                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);
                if (workbooks != null)
                    Marshal.ReleaseComObject(workbooks);
                if (excel != null)
                    Marshal.ReleaseComObject(excel);
            }
        }
    }

    public class Constants
    {
        public const string DefaultTxtBoxOutputFolder = DefaultOutputFolder;
        public const string DefaultTxtBoxExcelFile = "Please select a excel file";
        public const string SelectExcelFileButtonText = "Select Excel File";
        public const string SelectOutputFolderButtonText = " Output Folder";
        public const string DefaultOutputFolder = @"C:\Temp\ExcelConvertor\Out";
        public const string GenerateHtmlTabHeaderText = "Generate Html";
    }
}
