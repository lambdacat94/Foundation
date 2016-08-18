using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    // 每一个 Item 其实是代表一整个重点或者培育项目组的，包含多个单独项目
    // 存在 arr 中，其 total 是
    public class AllocItem
    {
        public AllocItem()
        {
            arr = new ArrayList();
            total = inner = outter = 0.0;
        }

        public ArrayList arr;
        // 总金额
        public double total;
        // 委内
        public double inner;
        // 委外
        public double outter;
    }
}
