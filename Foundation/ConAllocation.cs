using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;


using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Foundation
{
    public class ConAllocation
    {

        // 带参数的构造函数，打开资源
        public ConAllocation(string fileName, int sheetID)
        {
            if (fileName == null) return;
            try
            {
                using (FileStream file = new FileStream(@fileName, FileMode.Open, FileAccess.Read))
                {

                    wbk = new XSSFWorkbook(file);
                }
                sh = wbk.GetSheetAt(sheetID);

                keyItem = new AllocItem();
                culItem = new AllocItem();
            }
            catch(Exception e)
            {
                
                MessageBox.Show(e.Message);
            }
        }


        public void DoChecked(int index)
        {
            ((ItemDetails)keyItem.arr[index]).Inner--;
        }

        public void DoUnchecked()
        {

        }


        // 委内中加和值和应有值的差
        private int keyInDiff = 0;
        private int culInDiff = 0;
        
        // 计算重点项目分配的加和与应有值的差
        public void ComputeDiffs()
        {
            if (keyItem != null)
            {
                keyInDiff = (int)(((ItemDetails)keyItem.arr[keyItem.arr.Count - 1]).Inner - keyItem.inner);
            }

            if (culItem != null)
            {
                culInDiff = (int)(((ItemDetails)culItem.arr[culItem.arr.Count - 1]).Inner - culItem.inner);
            }

            MessageBox.Show(keyInDiff.ToString() + "  " + culInDiff.ToString());
        }





        // 从界面中读取总额、委内、委外的参数
        public void ReadFromInterface(double totals, double inner, double outter)
        {
            this.AllocaInner = inner;
            this.AllocaOutter = outter;
            this.AllocaTotals = totals;
        }

        // 从已经打开的 Excel 文件中读取
        // excel 的读取应该放在单独的类里，重构
        public AllocItem ReadFromExcel()
        {
            // 获取起始和终点行
            GetStartAndEndRow();

            keyItem.arr.Clear();
            culItem.arr.Clear();

            if (wbk != null && sh != null)
            {
                IRow row = null;
                string field = null;
                string number = null;
                string projID = null;
                string projName = null;
                string projLeader = null;
                string unit = null;

                double total = 0.0;
                double inner = 0.0;
                double outter = 0.0;

                double innersum = 0.0;
                double outtersum = 0.0;
                double totsum = 0.0;

                for (int i = keyStartRow; i <= keyEndRow; ++i)
                {
                    row = sh.GetRow(i);
                    field = row.GetCell(0).ToString();
                    number = row.GetCell(1).ToString();
                    projID = row.GetCell(2).ToString();
                    projName = row.GetCell(3).ToString();
                    projLeader = row.GetCell(4).ToString();
                    unit = row.GetCell(5).ToString();

                    total = Math.Round(Convert.ToDouble(row.GetCell(6).ToString()));
                    inner = Math.Round(total * this.AllocaInner / this.AllocaTotals);
                    outter = Math.Round(total * this.AllocaOutter / this.AllocaTotals);

                    innersum += inner;
                    outtersum += outter;
                    totsum += total;

                    keyItem.arr.Add(new ItemDetails()
                    {
                        Field = field,
                        Inner = inner,
                        Number = number,
                        ProjectID = projID,
                        ProjectName = projName,
                        ProjectLeader = projLeader,
                        Unit = unit,
                        Outter = outter,
                        Total = total
                    });
                }

                row = sh.GetRow(keyEndRow + 1);
                projName = row.GetCell(1).ToString();
                total = totsum;
                inner = Math.Round(total * this.AllocaInner / this.AllocaTotals);
                outter = Math.Round(total * this.AllocaOutter / this.AllocaTotals);

                // 最后添加一个计算得到的最终理论值，但这个理论值由于四舍五入的存在导致
                // 最终结果不对，需要重新平衡
                keyItem.arr.Add(new ItemDetails()
                {
                    Field = "",
                    Number = "",
                    ProjectID = "",
                    ProjectName = projName,
                    ProjectLeader = "",
                    Inner = inner,
                    Unit = "",
                    Outter = outter,
                    Total = total
                });


                keyItem.total = totsum;
                keyItem.inner = innersum;
                keyItem.outter = outtersum;


                if (culStartRow == 0 || culEndRow == 0 || culStartRow > culEndRow) return keyItem;

                innersum = outtersum = totsum = 0.0;

                for (int i = culStartRow; i <= culEndRow; ++i)
                {
                    row = sh.GetRow(i);
                    field = row.GetCell(0).ToString();
                    number = row.GetCell(1).ToString();
                    projID = row.GetCell(2).ToString();
                    projName = row.GetCell(3).ToString();
                    projLeader = row.GetCell(4).ToString();
                    unit = row.GetCell(5).ToString();

                    total = Math.Round(Convert.ToDouble(row.GetCell(6).ToString()));
                    inner = Math.Round(total * this.AllocaInner / this.AllocaTotals);
                    outter = Math.Round(total * this.AllocaOutter / this.AllocaTotals);

                    innersum += inner;
                    outtersum += outter;
                    totsum += total;

                    culItem.arr.Add(new ItemDetails()
                    {
                        Field = field,
                        Inner = inner,
                        Number = number,
                        ProjectID = projID,
                        ProjectName = projName,
                        ProjectLeader = projLeader,
                        Unit = unit,
                        Outter = outter,
                        Total = total
                    });
                }

                row = sh.GetRow(culEndRow + 1);
                projName = row.GetCell(1).ToString();
                total = totsum;
                inner = Math.Round(total * this.AllocaInner / this.AllocaTotals);
                outter = Math.Round(total * this.AllocaOutter / this.AllocaTotals);

                culItem.arr.Add(new ItemDetails()
                {
                    Field = "",
                    Number = "",
                    ProjectID = "",
                    ProjectName = projName,
                    ProjectLeader = "",
                    Inner = inner,
                    Unit = "",
                    Outter = outter,
                    Total = total
                });

                culItem.total = totsum;
                culItem.inner = innersum;
                culItem.outter = outtersum;
            }

            AllocItem rstItem = new AllocItem();
            for (int i = 0; i < keyItem.arr.Count; ++i)
            {
                rstItem.arr.Add(keyItem.arr[i]);
            }
            for (int i = 0; i < culItem.arr.Count; ++i)
            {
                rstItem.arr.Add(culItem.arr[i]);
            }

            return rstItem;
        }


        // 获得关于 Excel 文件中重点、培育的起始行号
        private void GetStartAndEndRow()
        {
            if (sh != null && wbk != null)
            {
                IRow row = null;
                ICell cell = null;
                for (int i = 3; ; i++)
                {
                    row = sh.GetRow(i);
                    if (row != null)
                    {
                        cell = row.GetCell(1);
                        if (cell != null && IsNumber(cell.ToString()))
                        {
                            if (keyStartRow == 0) keyStartRow = i;
                            keyEndRow = i;
                        }
                        else break;
                    }
                }

                for (int i = keyEndRow + 2; ; i++)
                {
                    row = sh.GetRow(i);
                    if (row != null)
                    {
                        cell = row.GetCell(1);
                        if (cell != null &&!IsNumber(cell.ToString()) && culStartRow != 0) break;

                        if (cell != null && IsNumber(cell.ToString()))
                        { 
                            if (culStartRow == 0) culStartRow = i;
                            culEndRow = i;
                        }
                        
                    }
                    
                }
            }
            /* 
             MessageBox.Show(keyStartRow.ToString() + " " + keyEndRow.ToString() + " " +
                culStartRow.ToString() + " " + culEndRow.ToString());
            */
        }

        // 获取舍入与计算时  委内  的差值
        public int GetDiffs()
        {
            return 0;
        }

        // 结果写回到 Excel 文件中
        public void WriteToExcel(string fileName)
        {
            if (sh != null && wbk != null)
            {
                IRow row = null;
                ICell cell = null;
                // 将重点写回文件
                for (int i = keyStartRow; i <= keyEndRow; ++i)
                {
                    row = sh.GetRow(i);
                    cell = row.GetCell(8);
                    cell.SetCellValue(((ItemDetails)keyItem.arr[i - keyStartRow]).Outter.ToString());
                    // MessageBox.Show(((ItemDetails)keyItem.arr[i - keyStartRow]).Outter.ToString());
                    // MessageBox.Show(cell.ToString() + "  " + (cell == null).ToString());
                    
                }

                for (int i = keyStartRow; i <= keyEndRow; ++i)
                {
                    row = sh.GetRow(i);
                    cell = row.GetCell(7);
                    cell.SetCellValue(((ItemDetails)keyItem.arr[i - keyStartRow]).Inner.ToString());
                }

                row = sh.GetRow(keyEndRow + 1);
                cell = row.GetCell(8);
                cell.SetCellValue(keyItem.inner.ToString());

                cell = row.GetCell(7);
                cell.SetCellValue(keyItem.outter.ToString());

                // 将培育写回文件
                for (int i = culStartRow; i <= culEndRow; ++i)
                {
                    row = sh.GetRow(i);
                    cell = row.GetCell(8);
                    cell.SetCellValue(((ItemDetails)culItem.arr[i - culStartRow]).Outter.ToString());

                    
                }
                for (int i = culStartRow; i <= culEndRow; ++i)
                {
                    row = sh.GetRow(i);
                    cell = row.GetCell(7);
                    cell.SetCellValue(((ItemDetails)culItem.arr[i - culStartRow]).Inner.ToString());
                }


                row = sh.GetRow(culEndRow + 1);
                cell = row.GetCell(8);
                cell.SetCellValue(culItem.inner.ToString());

                cell = row.GetCell(7);
                cell.SetCellValue(culItem.outter.ToString());


                using (FileStream fs = File.Open(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    wbk.Write(fs);
                    fs.Close();
                }

            }
        }
        

        private AllocItem keyItem;
        private AllocItem culItem;

        // 重点、培育项目在 Excel 中的起始和结束行号
        private int keyStartRow = 0;
        private int culStartRow = 0;
        private int keyEndRow = 0;
        private int culEndRow = 0;

       
        private double AllocaInner;
        private double AllocaOutter;
        private double AllocaTotals;

        private IWorkbook wbk = null;
        private ISheet sh = null;

        // 判断是否数字的工具函数，不应放在此处，重构时移到工具类里
        private bool IsNumber(string str)
        {
            if (str == string.Empty || str.Length == 0)
                return false;
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                { return false; }
            }
            return true;
        }

    }
}
