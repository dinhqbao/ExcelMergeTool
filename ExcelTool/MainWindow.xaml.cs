using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;

namespace ExcelTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public int StartRow { get; set; }
        public DateTime? filterDate { get; set; }
        public FileInfo template { get; set; }
        private ObservableCollection<ExcelFile> ExcelFileList { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            ExcelFileList = new ObservableCollection<ExcelFile>();

            LoadFormData();
        }

        private void LoadFormData()
        {
            int[] startRows = { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            cboStartRow.ItemsSource = startRows;
            cboStartRow.SelectedValue = 4;
            template = new FileInfo(@"./Template.xlsx");
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var fbd = new FolderBrowserDialog();
            fbd.Description = "Browse to folder...";

            if (fbd.ShowDialog().ToString().Equals("OK"))
            {
                txtFolderPath.Text = fbd.SelectedPath;
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
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
            using (var masterPackage = new ExcelPackage(template))
            {
                var selectedFile = ExcelFileList.Where(x => x.IsSelect == true);
                var startRow = 0;

                try
                {
                    var masterSheet = masterPackage.Workbook.Worksheets[txtSheetName.Text];
                    StartRow = int.Parse(cboStartRow.SelectedValue.ToString());

                    foreach (var file in selectedFile)
                    {
                        var pckg = new ExcelPackage(file.File);
                        var sheet = pckg.Workbook.Worksheets[txtSheetName.Text];

                        if (sheet == null)
                        {
                            MessageBox.Show("Không tìm thấy sheet tên " + txtSheetName.Text,
                                            "Confirmation",
                                            MessageBoxButton.OK,
                                            MessageBoxImage.Error);
                            return;
                        }

                        WriteData(masterSheet, sheet, ref startRow);
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
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                                    "Exception",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                }
            }
        }

        private void WriteData(ExcelWorksheet masterSheet, ExcelWorksheet sheet, ref int masterSheetIndex)
        {
            for (int row = StartRow; row <= sheet.Dimension.End.Row; row++)
            {
                if (sheet.Cells[row, 1].GetValue<int>() > 0)
                {
                    var requestDate = sheet.Cells[row, 2].GetValue<DateTime>();

                    if (filterDate == null || filterDate == requestDate)
                    {
                        masterSheet.SetValue(masterSheetIndex + StartRow, 1, masterSheetIndex + 1);
                        masterSheet.SetValue(masterSheetIndex + StartRow, 2, requestDate);
                        masterSheet.SetValue(masterSheetIndex + StartRow, 3, sheet.Cells[row, 3].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 4, sheet.Cells[row, 4].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 5, sheet.Cells[row, 5].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 6, sheet.Cells[row, 6].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 7, sheet.Cells[row, 7].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 8, sheet.Cells[row, 8].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 9, sheet.Cells[row, 9].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 10, sheet.Cells[row, 10].GetValue<int>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 11, sheet.Cells[row, 11].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 12, sheet.Cells[row, 12].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 13, sheet.Cells[row, 13].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 14, sheet.Cells[row, 14].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 15, sheet.Cells[row, 15].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 16, sheet.Cells[row, 16].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 17, sheet.Cells[row, 17].GetValue<string>());
                        masterSheet.SetValue(masterSheetIndex + StartRow, 18, sheet.Cells[row, 18].GetValue<string>());

                        masterSheetIndex++;
                    }
                }
                else
                {
                    break;
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

        private void dateFilter_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            filterDate = dateFilter.SelectedDate;
        }
    }
}
