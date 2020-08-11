using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace excelCompare
{
    /// <summary>
    /// Window1.xaml 的交互逻辑
    /// </summary>
    public partial class Window1 : Window
    {
        public ObservableCollection<MyersDiff.DiffData> resList = new ObservableCollection<MyersDiff.DiffData>();
        public ObservableCollection<MyersDiff.DiffData> resList2 = new ObservableCollection<MyersDiff.DiffData>();
        public const string NormalColor = "#FFFFFF";
        public const string AddColor = "#00FF00";
        public const string DeleteColor = "#FF0000";
        public const string ModifyColor = "#FFFF00";
        
        
        public string srcPath { get; set; }
        public string dstPath { get; set; }
        Dictionary<string, List<MyersDiff.DiffData>[]> wbDiffDic = new Dictionary<string, List<MyersDiff.DiffData>[]>();
        private ScrollViewer sv1, sv2;
        public Window1()
        {
            InitializeComponent();
            
            grid1.ItemsSource = resList;
            grid2.ItemsSource = resList2;
            this.DataContext = this;
            this.Loaded += Window1_Loaded;
            //分别获取两个DataGrid的ScrollViewer
            
        }

        public static void ShowCompareWin(string src, string dst)
        {
            if (string.IsNullOrEmpty(src) || string.IsNullOrEmpty(dst))
            {
                MessageBox.Show("重新选择文件");
                return;
            }
            if (src == dst)
            {
                MessageBox.Show("文件源相同，请重新选择文件");
                return;
            }
            Window1 window1 = new Window1();
            window1.Show();
            window1.Init(src, dst);
        }

        private void Window1_Loaded(object sender, RoutedEventArgs e)
        {
            sv1 = VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(this.grid1, 0), 0) as ScrollViewer;
            sv2 = VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(this.grid2, 0), 0) as ScrollViewer;

            //关联ScrollChanged事件
            sv1.ScrollChanged += new ScrollChangedEventHandler(sv1_ScrollChanged);
            sv2.ScrollChanged += new ScrollChangedEventHandler(sv2_ScrollChanged);
        }

        private void sv2_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            sv1.ScrollToHorizontalOffset(sv2.HorizontalOffset);
            sv1.ScrollToVerticalOffset(sv2.VerticalOffset);
        }

        private void sv1_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            sv2.ScrollToVerticalOffset(sv1.VerticalOffset);
            sv2.ScrollToHorizontalOffset(sv1.HorizontalOffset);
        }

        public void Init(string srcPath,string dstPath)
        {
            srcPathLabel.Content = srcPath;
            dstPathLabel.Content = dstPath;
            try
            {
                FileStream fs1 = new FileStream(srcPath, FileMode.Open, FileAccess.Read);
                IWorkbook srcWb = WorkbookFactory.Create(fs1);
                FileStream fs2 = new FileStream(dstPath, FileMode.Open, FileAccess.Read);
                IWorkbook dstWb = WorkbookFactory.Create(fs2);
                fs1.Close();
                fs2.Close();
                List<string> srcSheetNameList = GetSheetNames(srcWb);
                List<string> dstSheetNameList = GetSheetNames(dstWb);
                HashSet<string> nameHashSet = new HashSet<string>(srcSheetNameList);
                for (int i = 0; i < dstSheetNameList.Count; i++)
                {
                    nameHashSet.Add(dstSheetNameList[i]);
                }
                bool first = true;
                foreach (var name in nameHashSet)
                {
                    var srcSheet = srcWb.GetSheet(name);
                    var dstSheet = dstWb.GetSheet(name);
                    List<string> srcStrList = ConverSheetToString(srcSheet);
                    List<string> dstStrList = ConverSheetToString(dstSheet);
                    var diffRes = MyersDiff.Diff(srcStrList, dstStrList);
                    wbDiffDic[name] = diffRes;
                    var btn = createTabBtn(name, diffRes[0]);
                    TabStackPanel.Children.Add(btn);
                    if (first)
                    {
                        btn.RaiseEvent(new RoutedEventArgs(Button.ClickEvent, btn));
                        first = false;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show( e.Message + "\n"+e.StackTrace + "\n请关闭打开excel的应用");
            }
            
            
        }

        private Button createTabBtn(string name,List<MyersDiff.DiffData> diffRes)
        {
            Button btn = new Button
            {
                Name = name,
                Content = name,
                Height = 23,
                MinWidth = 50,
                MaxWidth = 200,
                Margin = new Thickness(10, 10, 0, 0), 
                Visibility = Visibility.Visible
            };
            SetBtnColor(btn, CalcateStateColor(diffRes));
            btn.Click += new RoutedEventHandler(btn_click);
            return btn;
        }

        private string CalcateStateColor(IEnumerable<MyersDiff.DiffData> diffRes)
        {
            int count = 0;
            int i = 0;
            bool same = true;
            foreach (var item in diffRes)
            {
                if(item.oper == MyersDiff.Add)
                {
                    i++;
                    same = false;
                }else if (item.oper == MyersDiff.Del)
                {
                    i--;
                    same = false;
                }
                count++;
            }
            if (same)
            {
                return NormalColor;
            }
            else
            {
                if(i == count)
                {
                    return AddColor;
                }else if(i == -count) {
                    return DeleteColor;
                }
                else
                {
                    return ModifyColor;
                }
            }
        }


        private Button lastBtn;
        private void btn_click(object sender, RoutedEventArgs e)
        {
            if(lastBtn != null)
            {
                lastBtn.Content = lastBtn.Name;
            }
            var button = e.Source as Button;
            lastBtn = button;
            button.Content = button.Name + "(选中)";
            updateDiffDataGrid(button.Name);
        }

        private void updateDiffDataGrid(string name)
        {
            var diffRes = wbDiffDic[name];
            resList.Clear();
            resList2.Clear();

            foreach (var item in diffRes[1])
            {
                resList.Add(item);
            }
            foreach (var item in diffRes[2])
            {
                resList2.Add(item);
            }
        }

        private void SetBtnColor(Button btn, string color)
        {
            BrushConverter conv = new BrushConverter();
            Brush bru = conv.ConvertFromInvariantString(color) as Brush;
            btn.Background = (System.Windows.Media.Brush)bru;
        }

        public List<string> ConverSheetToString(ISheet sheet)
        {
            List<string> stringList = new List<string>();
            if(sheet != null)
            {
                int nRowsCount = sheet.LastRowNum + 1;
                StringBuilder sb = new StringBuilder();
                int cellCount = 0;
                for (int k = 0; k < nRowsCount; k++)
                {
                    IRow row = sheet.GetRow(k);
                    if (row!=null)
                    {
                        if (k == 0)
                        {
                            cellCount = row.Cells.Count;
                        }
                        for (int j = 0; j < Math.Min(row.Cells.Count, cellCount); j++)
                        {
                            var cell = row.Cells[j];
                            if (cell != null)
                            {
                                sb.Append(cell.ToString());
                            }
                            sb.Append("|");
                        }
                    }
                    


                    
                    stringList.Add(sb.ToString());
                    sb.Clear();
                }

            }
            return stringList;
        }

        private List<string> GetSheetNames(IWorkbook wb)
        {
            List<string> nameList = new List<string>();
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                var sheet = wb.GetSheetAt(i);
                
                nameList.Add(wb.GetSheetName(i));
            }
            return nameList;
        }
    }


    public class MyersDiff {
        public enum Operation
        {
            Add = 1,
            Delete = 2,
            Move = 3,
            None = 4,
        }

        public const string Add = "+";
        public const string Del = "-";
        public const string Move = "";
        public const string None = "";

        public class DiffData {
            public string index { set; get; }
            public string oper { set; get; }
            public string content { set; get; }

            public static DiffData EmptyDiffData()
            {
                return new DiffData() {
                    content = "",
                    oper = None,
                    index = "",
                };

            }

            public DiffData()
            {

            }

            public DiffData(string content,string oper,string index)
            {
                this.index = index;
                this.oper = oper;
                this.content = content;
            }
        }


        public static List<DiffData>[] Diff(List<string> src,List<string> dst)
        {
            List<DiffData>[] diffResult = new List<DiffData>[3];
            List<DiffData> diffRes = new List<DiffData>();
            List<DiffData> diffSrc = new List<DiffData>();
            List<DiffData> diffDst = new List<DiffData>();
            var res = ShortestEditScript(src, dst);
            int index1 = 0;
            int index2 = 0;
            for (int i = 0; i < res.Count; i++)
            {
                var oper = res[i];
                switch (oper)
                {
                    case Operation.Add:
                        DiffData diff = new DiffData(dst[index2], Add, index2.ToString());
                        diffRes.Add(diff);
                        diffDst.Add(diff);
                        index2++;
                        break;
                    case Operation.Move:
                        DiffData diff2 = new DiffData(dst[index2], Move, index2.ToString());
                        diffRes.Add(diff2);
                        diffDst.Add(diff2);
                        diffSrc.Add(new DiffData(src[index1], Move, index1.ToString()));
                        index2++;
                        index1++;
                        break;
                    case Operation.Delete:
                        DiffData diff3 = new DiffData(src[index1], Del, index1.ToString());
                        diffRes.Add(diff3);
                        diffSrc.Add(diff3);
                        index1++;
                        break;
                }
            }
            diffResult[0] = diffRes;
            diffResult[1] = diffSrc;
            diffResult[2] = diffDst;
            return diffResult;
        }

        public static List<Operation> ShortestEditScript(List<string> src, List<string> dst)
        {
            int n = src.Count;
            int m = dst.Count;
            int max = n + m;
            List<Dictionary<int, int>> trace = new List<Dictionary<int, int>>();
            int x, y;
            for (int i = 0; i <= max; i++)
            {
                //Logger.Log(i);
                var v = new Dictionary<int, int>();
                trace.Add(v);
                if (i == 0)
                {
                    int t = 0;
                    while (n > t && m > t && src[t] == dst[t])
                    {
                        t++;
                    }
                    v[0] = t;
                    if(t == m && t == n)
                    {
                        goto quit;
                    }
                    continue;
                }
                var lastV = trace[i - 1];
                for (int j = -i; j <= i; j += 2)
                {
                    if (j == -i || (j != i && lastV[j - 1] < lastV[j + 1]))
                    {
                        x = lastV[j + 1];
                    }
                    else
                    {
                        x = lastV[j - 1] + 1;
                    }
                    y = x - j;
                    while (x < n && y < m && src[x] == dst[y])
                    {
                        x = x + 1;
                        y = y + 1;
                    }
                    v[j] = x;
                    if (x == n && y == m)
                    {
                        goto quit;
                    }
                }
            }

        //return;
        quit:
            //Log(trace.Count);
            //Logger.LogTree(trace);
            //1添加 2删除 3移动
            List<Operation> script = new List<Operation>();
            x = n;
            y = m;
            int k, prevK, prevX, prevY;
            for (int d = trace.Count - 1; d > 0; d--)
            {
                k = x - y;
                var lastV = trace[d - 1];
                if (k == -d || (k != d && lastV[k - 1] < lastV[k + 1]))
                {
                    prevK = k + 1;
                }
                else
                {
                    prevK = k - 1;
                }

                prevX = lastV[prevK];
                prevY = prevX - prevK;
                while (x > prevX && y > prevY)
                {
                    script.Add(Operation.Move);
                    x -= 1;
                    y -= 1;
                }
                if (x == prevX)
                {
                    script.Add(Operation.Add);
                }
                else
                {
                    script.Add(Operation.Delete);
                }
                x = prevX;
                y = prevY;
            }
            if (trace[0][0] != 0)
            {
                for (int i = 0; i < trace[0][0]; i++)
                {
                    script.Add(Operation.Move);
                }
            }
            script.Reverse();
            return script;
        }

    }

}
