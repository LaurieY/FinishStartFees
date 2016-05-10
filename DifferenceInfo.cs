using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Syncfusion.XlsIO;
using Syncfusion.Windows.Shared;
namespace FinishStartFees
{
    public struct DateandValue
    {
        public string txnDate { get; set; }
        public decimal txnAmt { get; set; }

    }

    class DifferenceInfo
    {
        public Dictionary<string, int> fileCols { get; set; }
        public Dictionary<long, Dictionary<string, Object>> fileScan { get; set; }//= new Dictionary<string, Object>();

        public string fileName { get; set; }
        public static string fileName1 { get; set; }
        public static string fileName2 { get; set; }
        public DifferenceInfo()
        {
           
        }

 
     
       
    }
}