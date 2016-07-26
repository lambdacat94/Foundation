using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    public sealed class Department
    {
        private string departmentName;
        public string DepartmentName
        {
            set
            {
                if (value is string)
                    departmentName = value;
            }
            get
            {
                return departmentName;
            }
        }

        private int keyAcceptCount;
        public int KeyAcceptCount
        {
            set
            {
                keyAcceptCount = value;
            }
            get
            {
                return keyAcceptCount;
            }
        }

        private int culAcceptCount;
        public int CulAcceptCount
        {
            set
            {
                culAcceptCount = value;
            }
            get
            {
                return culAcceptCount;
            }
        }

        private int youngAcceptCount;
        public int YoungAcceptCount
        {
            set
            {
                youngAcceptCount = value;
            }
            get
            {
                return youngAcceptCount;
            }
            
        }

        public Department()
        {

        }

        public Department(string newName, int keyCount, int culCount, int youngCount)
        {
            this.DepartmentName = newName;
            this.KeyAcceptCount = keyCount;
            this.CulAcceptCount = culCount;
            this.YoungAcceptCount = youngCount;
        }
    }
}
