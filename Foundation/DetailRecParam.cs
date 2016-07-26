using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    class DetailRecParam
    {
        public string DepartmentName { get; set; }

        public int CulAcceptCount { get; set; }
        public int KeyAcceptCount { get; set; }
        public int YoungAcceptCount { get; set; }

        public double CulRate { get; set; }
        public double KeyRate { get; set; }
        public double YoungRate { get; set; }

        public double CulCaled { get; set; }
        public int CulRound { get; set; }
        public double KeyCaled { get; set; }
        public int KeyRound { get; set; }
        public double YoungCaled { get; set; }
        public int YoungRound { get; set; }

        public double Total { get; set; }

        public RecParameter RecParam { get; set; }
    }
}
