using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    public class ItemDetails
    {
        public string Field { get; set; }
        public string Number { get; set; }
        public string ProjectID { get; set; }
        public string ProjectName { get; set; }
        public string ProjectLeader { get; set; }
        public string Unit { get; set; }
        // 小计
        public double Total { get; set; }
        // 委内
        public double Inner { get; set; }
        // 委外
        public double Outter { get; set; }

        public bool Checked { get; set; }
    }
}
