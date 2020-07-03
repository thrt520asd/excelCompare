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
//using NPOI.SS.UserModel;

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
            btn1.Content = "请选择第一个文件";
            btn2.Content = "请选择第二个文件";
            
            
        }

      

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string filePath = SelectFileWpf();
            Button button = e.Source as Button;
            
            button.Content = System.IO.Path.GetFileName(filePath);
            excel_path_1 = filePath;
            MessageBox.Show(filePath);
        }

        



        public string SelectFileWpf()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "Excel (.xlsx)|*.xlsx|All files (*.*)|*.*"
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
            if (string.IsNullOrEmpty(excel_path_1) || string.IsNullOrEmpty(excel_path_2))
            {
                MessageBox.Show("重新选择文件");
                return;
            }
            if(excel_path_1 == excel_path_2)
            {
                MessageBox.Show("文件源相同，请重新选择文件");
                return;
            }
            Window1 window1 = new Window1();
            window1.Show();
            window1.Init(excel_path_1, excel_path_2);
            
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

    public static class Logger {
        public static void Log(object obj)
        {
            Trace.WriteLine(obj.ToString());
        }

        public static void LogTree(List<Dictionary<int, int>> dicList)
        {
            for (int d = 0; d < dicList.Count; d++)
            {
                Log("d = " + d.ToString());
                var v = dicList[d];
                for (int k = -d; k <= d; k += 2)
                {
                    if (v.ContainsKey(k))
                    {
                        var x = v[k];
                        var y = x - k;
                        Log(string.Format("k = {0}:({1},{2})", k, x, y));
                    }
                }
            }
        }
    }

}
