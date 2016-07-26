using System;

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    public sealed class InitParameters
    {
        
        private ArrayList departmentArray;
        public ArrayList GetDataBanding()
        {
            return departmentArray;
        }


        // 返回学部数量
        private Department departmentStatistic;
        public int DepartmentCount
        {
            get
            {
                if (departmentArray != null)
                    return departmentArray.Count;
                else
                    return 0;
            }
        }

        public InitParameters()
        {
            departmentStatistic = new Department("统计合计", 0, 0, 0);

            departmentArray = new ArrayList();
        }

        // 返回一个统计信息引用，未作出防止客户修改引用的防范
        public Department GetStatistic()
        {
            departmentStatistic.KeyAcceptCount = 0;
            departmentStatistic.CulAcceptCount = 0;
            departmentStatistic.YoungAcceptCount = 0;
            foreach (Department department in departmentArray)
            {
                departmentStatistic.KeyAcceptCount += department.KeyAcceptCount;
                departmentStatistic.CulAcceptCount += department.CulAcceptCount;
                departmentStatistic.YoungAcceptCount += department.YoungAcceptCount;
            }
            return departmentStatistic;
        }

        public void AddDepartment(Department newDepartment)
        {
            departmentArray.Add(newDepartment);
        }
    }
}
