using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;

namespace ExcelTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<ExcelFile> ExcelFileList { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            ExcelFileList = new ObservableCollection<ExcelFile>();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var fbd = new FolderBrowserDialog();
            fbd.Description = "Browse to folder..."; //not mandatory

            if (fbd.ShowDialog().ToString().Equals("OK"))
            {
                txtFolderPath.Text = fbd.SelectedPath;
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            var template = new FileInfo("./Template.xlsx");

            if (!template.Exists)
            {
                var fd = new OpenFileDialog()
                {
                    Filter = "xlsx template files (*.xlsx)|*.xlsx",
                    Multiselect = false,
                };

                if (fd.ShowDialog().ToString().Equals("OK"))
                {
                    template = new FileInfo(fd.FileName);
                }
                else
                {
                    return;
                }
            }

            BeginExport(template);
        }
        
        void BeginExport(FileInfo template)
        {
            var resultData = new List<DataRow>();
            var selectedFile = ExcelFileList.Where(x => x.IsSelect == true);

            using (var masterPackage = new ExcelPackage(template))
            {
                foreach (var file in selectedFile)
                {
                    var pckg = new ExcelPackage(file.File);

                    foreach (var sheet in pckg.Workbook.Worksheets)
                    {
                        var startRow = 2;
                        var endRow = sheet.Dimension.End.Row;
                        for (int row = startRow; row <= endRow; row++)
                        {
                            var d = new DataRow()
                            {
                                Name = sheet.Cells[row, 1].Text,
                                Product = sheet.Cells[row, 2].Text,
                                Description = sheet.Cells[row, 3].Text,
                            };
                            resultData.Add(d);
                        }
                    }
                }

                var masterSheet = masterPackage.Workbook.Worksheets[1];

                for (int i = 0; i < resultData.Count; i++)
                {
                    masterSheet.SetValue(i + 2, 1, resultData[i].Name);
                    masterSheet.SetValue(i + 2, 2, resultData[i].Product);
                    masterSheet.SetValue(i + 2, 3, resultData[i].Description);
                }

                SaveFileDialog sfd = new SaveFileDialog()
                {
                    FileName = $"Result_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}",
                    DefaultExt = ".xlsx",
                    Filter = "Excel Sheet (.xlsx)|*.xlsx"
                };

                if (sfd.ShowDialog().ToString() == "OK")
                {
                    var file = new FileInfo(sfd.FileName);
                    masterPackage.SaveAs(file);
                    Process.Start(file.FullName);
                }
            }
        }

        private void txtFolderPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            getDataFormPath((sender as System.Windows.Controls.TextBox).Text);
        }

        private void getDataFormPath(string path)
        {
            ExcelFileList.Clear();

            try
            {
                var files = new DirectoryInfo(path).GetFiles("*.xlsx", SearchOption.AllDirectories);

                foreach (var file in files)
                {
                    ExcelFileList.Add(new ExcelFile
                    {
                        IsSelect = true,
                        File = file
                    });
                }
                grdExcelFiles.ItemsSource = ExcelFileList;
            }
            catch
            {

            }

            btnStart.IsEnabled = ExcelFileList.Count > 0 ? true : false;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtFolderPath.TextChanged += new TextChangedEventHandler(this.txtFolderPath_TextChanged);
        }
    }
}
