using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinishStartFees
{
    class FileScan2
    {
        public decimal feeTotal { get; set; }
        public int startRow { get; set; }
        public int endRow { get; set; }
        public string propName { get; set; }
        public decimal startBalance { get; set; }
        public decimal finishBalance { get; set; }
        public decimal payTotal { get; set; }
        public List<DateandValue> fees { get; set; }
        public List<DateandValue> payments { get; set; }
        public FileScan2()
        {
            propName = "";
        }
    }
}
