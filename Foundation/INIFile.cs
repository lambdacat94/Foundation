using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace Foundation
{
    class INIFile
    {
        private string INIPath;

        [DllImport("Kernel32")]
        private static extern long WritePrivateProfileString(string Section, string Key, string Val, string FilePath);

        [DllImport("Kernel32")]
        private static extern long GetPrivateProfileString(string Section, string Key, string Def, StringBuilder RetVal, int Size, string FilePath);


        public INIFile(string FilePath)
        {
            INIPath = FilePath;
        }


        public string ReadIniData(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(1024);
            long a = GetPrivateProfileString(Section, Key, "", temp, 1024, INIPath);

            return temp.ToString();
        }

        public bool WriteIniData(string Section, string Key, string Val)
        {
            if (File.Exists(INIPath))
            {
                long OpStation = WritePrivateProfileString(Section, Key, Val, INIPath);
                if (OpStation == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return true;
            }
        }

        public bool ExistINIFile()
        {
            return File.Exists(INIPath);
        }
    }
}
