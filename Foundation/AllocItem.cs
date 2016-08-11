using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    public class AllocItem
    {
        public AllocItem()
        {
            arr = new ArrayList();
            total = inner = outter = 0.0;
        }

        public ArrayList arr;

        public double total;
        public double inner;
        public double outter;
    }
}
