﻿using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace excelCompare
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            if(e.Args.Length > 2)
            {
                string localPath = e.Args[0] + "/" + e.Args[1];
                string remotePath = e.Args[2];
                Window1.ShowCompareWin(remotePath, localPath);
            }

        }
    }
}
