using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPMConnectAddin
{
    public class BOM
    {
        public int Id { get; set; }
        public int Qty { get; set; }
        public string ItemNo { get; set; }
        public string Description { get; set; }
        public string Path { get; set; }
    }
}
