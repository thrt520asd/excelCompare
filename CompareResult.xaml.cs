using NPOI.OpenXmlFormats.Dml.Diagram;
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

        public const string NormalColor = "#FFFFFF";
        public const string AddColor = "#00FF00";
        public const string DeleteColor = "#FF0000";
        public const string ModifyColor = "#FFFF00";
        private string _srcPath;
        private string _dstPath;
        public string srcPath { get => _srcPath; }
        public string dstPath { get => _dstPath; }
        Dictionary<string, List<MyersDiff.DiffData>> wbDiffDic = new Dictionary<string, List<MyersDiff.DiffData>>();
        public Window1()
        {
            InitializeComponent();
            
            grid1.ItemsSource = resList; 
            this.DataContext = this;
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
                    var btn = createTabBtn(name, diffRes);
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
                MessageBox.Show(e.Message+"\n请关闭打开excel的应用");
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
                if(item.oper == MyersDiff.Operation.Add)
                {
                    i++;
                    same = false;
                }else if (item.oper == MyersDiff.Operation.Delete)
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
            Logger.Log(diffRes.Count);
            foreach (var item in diffRes)
            {
                Logger.Log(item.oper.ToString() + item.content.ToString());
            }
            resList.Clear();
            foreach (var item in diffRes)
            {
                resList.Add(item);
            }
            //resList = new ObservableCollection<MyersDiff.DiffData>(diffRes);
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
                for (int k = 0; k < nRowsCount - 1; k++)
                {
                    IRow row = sheet.GetRow(k);
                    for (int j = 0; j < row.Cells.Count - 1; j++)
                    {
                        var cell = row.GetCell(j);
                        sb.Append(row.GetCell(j).ToString());
                        sb.Append("|");
                    }
                    stringList.Add(sb.ToString());
                    //Logger.Log(k + ":" + sb.ToString());
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
        }

        public class DiffData {
            public int index { set; get; }
            public Operation oper { set; get; }
            public string content { set; get; }

            public string operStr { set; get; }
        }


        public static List<DiffData> Diff(List<string> src,List<string> dst)
        {
            List<DiffData> diffRes = new List<DiffData>();
            var res = ShortestEditScript(src, dst);
            int index1 = 0;
            int index2 = 0;
            for (int i = 0; i < res.Count; i++)
            {
                var oper = res[i];
                switch (oper)
                {
                    case Operation.Add:
                        diffRes.Add(new DiffData() {
                            content = dst[index2], 
                            oper = Operation.Add,
                            operStr = "+",
                            index = index2 });
                        index2++;
                        break;
                    case Operation.Move:
                        diffRes.Add(new DiffData() { 
                            content = src[index1], 
                            oper = Operation.Move,
                            operStr = " ",
                            index = index1 });
                        index2++;
                        index1++;
                        break;
                    case Operation.Delete:
                        diffRes.Add(new DiffData()
                        {
                            content = src[index1],
                            oper = Operation.Delete,
                            operStr = "-",
                            index = index1
                        });
                        index1++;
                        break;
                }
            }
            foreach (var item in diffRes)
            {
                Logger.Log(item.oper.ToString() + item.content.ToString());
            }
            return diffRes;
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
            Logger.LogTree(trace);
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
