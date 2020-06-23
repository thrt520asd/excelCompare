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

        public enum Operation { 
            Add = 1,
            Delete = 2,
            Move = 3,
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
            test();
            return;
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

        //private string GetSheetRow(ISheet sheet)
        //{
        //    StringBuilder sb = new StringBuilder();
            
        //    for (int i = 0; i < sheet.LastRowNum; i++)
        //    {
                
        //    }
        //}

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

        
        private void test()
        {
            List<string> list1 = new List<string>() { "a","b","c","d"};
            List<string> list2 = new List<string>() { "a", "b", "c", "d","e" };

            var res = shortestEditScript(list1, list2);
            int index1 = 0;
            int index2 = 0;

            for (int i = 0; i < res.Count; i++)
            {
                var oper = res[i];
                switch (oper) {
                    case Operation.Add:
                        Trace.WriteLine("+" + list2[index2]);
                        index2++;
                        break;
                    case Operation.Move:
                        Trace.WriteLine(" " + list1[index1]);
                        index2++;
                        index1++;
                        break;
                    case Operation.Delete:
                        Trace.WriteLine("-" + list1[index1]);
                        index1++;
                        break;
                }

            }
        }



        private List<Operation> shortestEditScript(List<string> src, List<string> dst)
        {
            int n = src.Count;
            int m = dst.Count;
            int max = n + m;
            List<Dictionary<int, int>> trace = new List<Dictionary<int, int>>();
            int x, y;
            for (int i = 0; i <= max; i++)
            {
                var v = new Dictionary<int, int>();
                trace.Add(v);
                if(i == 0)
                {
                    int t = 0;
                    while (n > t && m > t && src[t] == dst[t])
                    {
                        t++;
                    }
                    v[0] = t;
                    continue;
                }
                var lastV = trace[i - 1];
                for (int j = -i; j <= i; j+=2)
                {
                    if(j == -i || (j!= i && lastV[j-1] < lastV[j + 1])){
                        x = lastV[j + 1];
                    }
                    else
                    {
                        x = lastV[j - 1] + 1;
                    }
                    y = x - j;
                    while (x<n&& y<m && src[x] == dst[y])
                    {
                        x = x + 1;
                        y = y + 1;
                    }
                    v[j] = x;
                    if(x == n && y == m)
                    {
                        break;
                    }
                }
            }
            //1添加 2删除 3移动
            List<Operation> script = new List<Operation>();
            x = n;
            y = m;
            int k, prevK, prevX, prevY;
            for (int d = trace.Count; d > 0; d--)
            {
                k = x - y;
                var lastV = trace[d - 1];
                if(k == -d||k !=d && lastV[k - 1] < lastV[k + 1])
                {
                    prevK = k + 1;
                }
                else
                {
                    prevK = k - 1;
                }

                prevX = lastV[prevK];
                prevY = prevX - prevK;
                while(x>prevX&& y > prevY)
                {
                    script.Add(Operation.Move);
                    x -= 1;
                    y -= 1;
                }
                if(x == prevX)
                {
                    script.Append(Operation.Add);
                }
                else
                {
                    script.Append(Operation.Delete);
                }
                x = prevX;
                y = prevY;
            }
            if(trace[0][0] != 0)
            {
                for (int i = 0; i < trace[0][0]; i++)
                {
                    script.Add(Operation.Move);
                }
            }
             script.Reverse();
            return script;
            //return reverse(script);
        }

        
    }
}
