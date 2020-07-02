using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace excelCompare.ViewModel
{
    public class CompareResultViewModel:ViewModelBase
    {
        private ObservableCollection<CompareResult> compareResultList;

        public void Init(List<CompareResult> resList)
        {
            compareResultList = new ObservableCollection<CompareResult>(resList);
        }

        public void testInit()
        {
            this.Init(new List<CompareResult>() { 
                new CompareResult(){ index =1 ,oper="+",content="add"},
                new CompareResult(){ index =2 ,oper=" ",content="raw"},
                new CompareResult(){ index =3 ,oper="-",content="del"},
            });
        }

    }

    public class CompareResult
    {
        public int index;
        public string oper;
        public string content;
    }
}
