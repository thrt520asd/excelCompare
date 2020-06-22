using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using System.IO;
using NPOI.SS.UserModel;

namespace excelCompare
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string excel_path_1 = "";
        private string excel_path_2 = "";
        public MainWindow()
        {
            InitializeComponent();
            btn1.Content = "未选中";
            btn2.Content = "未选中";

        }

        

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string filePath = SelectFileWpf();
            Button button = e.Source as Button;
            button.Content = System.IO.Path.GetFileName(filePath);
            excel_path_1 = filePath;
            MessageBox.Show(filePath);
        }

        

        public void Compare(string excel1,string excel2)
        {

        }


        public string SelectFileWpf()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "Excel (.xls)|*.xls|All files (*.*)|*.*"
            };
            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                return openFileDialog.FileName;
            }
            else
            {
                return null;
            }
        }

        private void compareBtn_Click(object sender, RoutedEventArgs e)
        {
            if(string.IsNullOrEmpty(excel_path_1) || string.IsNullOrEmpty(excel_path_2))
            {
                MessageBox.Show("重新选择文件");
                return;
            }
            if(excel_path_1 == excel_path_2)
            {
                MessageBox.Show("文件源相同，请重新选择文件");
                return;
            }
            FileStream fs1 = new FileStream(excel_path_1, FileMode.Open, FileAccess.ReadWrite);
            IWorkbook wb1 = WorkbookFactory.Create(fs1);
            FileStream fs2 = new FileStream(excel_path_1, FileMode.Open, FileAccess.ReadWrite);
            IWorkbook wb2 = WorkbookFactory.Create(fs2);
            string content1 = getWBContent(wb1);
            string content2 = getWBContent(wb2);
            Trace.WriteLine("content1"+content1);
            Trace.WriteLine("content2"+content2);
        }

        private string GetSheetRow(ISheet sheet)
        {
            StringBuilder sb = new StringBuilder();
            
            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                
            }
        }

        private string getWBContent(IWorkbook workbook) {
            int nSheets = workbook.NumberOfSheets;
            for (int i = 0; i < nSheets - 1; i++)
            {
                
                //获取表格名字
                string strSheetName = workbook.GetSheetName(i);
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    int nRowsCount = sheet.LastRowNum + 1;
                    StringBuilder sb = new StringBuilder();
                    for (int k = 0; k < nRowsCount - 1; k++)
                    {
                        IRow row = sheet.GetRow(k);
                        for (int j = 0; j < row.Cells.Count - 1; j++)
                        {
                            var cell = row.GetCell(j);
                            sb.Append(row.GetCell(j).ToString());
                            sb.Append("|");
                        }
                        sb.AppendLine();
                    }
                    return sb.ToString();
                }
            }
            return "";
        }

        

        private void btn2_Click(object sender, RoutedEventArgs e)
        {
            string filePath = SelectFileWpf();
            if (string.IsNullOrEmpty(filePath))
            {
                return;
            }
            btn2.Content = System.IO.Path.GetFileName(filePath);
            excel_path_2 = filePath;
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            string filePath = SelectFileWpf();
            if (string.IsNullOrEmpty(filePath))
            {
                return;
            }
            btn1.Content = System.IO.Path.GetFileName(filePath);
            excel_path_1 = filePath;
        }
    }
}
