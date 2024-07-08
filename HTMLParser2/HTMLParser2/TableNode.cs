
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTMLParser2
{
    public class TableNode
    {
        public string TradeName { get; set; }
        public string TradeType { get; set; }
        public string ShortName { get; set; }
        public string FutureExpiry { get; set; }
        public string Strike { get; set; }
        public string CallPut { get; set; }
        public string Quantity { get; set; }
        public string Vol { get; set; }
        public string Premium { get; set; }
        public string FuturesPrice { get; set; }

    }
}