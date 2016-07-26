using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;


namespace Foundation
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private InitParameters initParam;
        private Calculator calculator;

        public MainWindow()
        {
            InitializeComponent();

            initParam = new InitParameters();
            calculator = new Calculator();
            resultArray = new ArrayList();

            // 完成两个ListView的数据绑定
            //LstShowDeps.DataContext = initParam.GetDataBanding();
            LstShowDeps.ItemsSource = initParam.GetDataBanding();
            LstShowDetails.DataContext = calculator.GetUltimateArrayRef();

            isAlter = false;
        }


        private void DisplayDetailsLst(ArrayList detArray)
        {
            LstShowDetails.Items.Clear();
            foreach (DetailRecParam det in detArray)
            {
                LstShowDetails.Items.Add(det);
            }
        }

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


        // 确保输入的都是数字而非空或者其他字符
        private bool ConfirmFillIn()
        {
            if (!IsNumber(TbxTotalMoney.Text.ToString())) TbxTotalMoney.Text = "0";
            if (!IsNumber(TbxRemMoneyMin.Text.ToString())) TbxRemMoneyMin.Text = "0";
            if (!IsNumber(TbxRemMoneyMax.Text.ToString())) TbxRemMoneyMax.Text = "0";
            if (!IsNumber(TbxKeyCountMin.Text.ToString())) TbxKeyCountMin.Text = "0";
            if (!IsNumber(TbxKeyCountMax.Text.ToString())) TbxKeyCountMax.Text = "0";
            if (!IsNumber(TbxCulCountMin.Text.ToString())) TbxCulCountMin.Text = "0";
            if (!IsNumber(TbxCulCountMax.Text.ToString())) TbxCulCountMax.Text = "0";
            if (!IsNumber(TbxKeyIntenMin.Text.ToString())) TbxKeyIntenMin.Text = "0";
            if (!IsNumber(TbxKeyIntenMax.Text.ToString())) TbxKeyIntenMax.Text = "0";
            if (!IsNumber(TbxCulIntenMin.Text.ToString())) TbxCulIntenMin.Text = "0";
            if (!IsNumber(TbxCulIntenMax.Text.ToString())) TbxCulIntenMax.Text = "0";
            if (!IsNumber(TbxYoungCountMin.Text.ToString())) TbxYoungCountMin.Text = "0";
            if (!IsNumber(TbxYoungIntenMin.Text.ToString())) TbxYoungIntenMin.Text = "0";


            if (IsNumber(TbxTotalMoney.Text.ToString()) && IsNumber(TbxRemMoneyMin.Text.ToString()) &&
                IsNumber(TbxRemMoneyMax.Text.ToString()) && IsNumber(TbxKeyCountMin.Text.ToString()) &&
                IsNumber(TbxKeyCountMax.Text.ToString()) && IsNumber(TbxCulCountMin.Text.ToString()) &&
                IsNumber(TbxCulCountMax.Text.ToString()) && IsNumber(TbxKeyIntenMin.Text.ToString()) &&
                IsNumber(TbxKeyIntenMax.Text.ToString()) && IsNumber(TbxCulIntenMin.Text.ToString()) &&
                IsNumber(TbxCulIntenMax.Text.ToString())
                )
            {
                return true;
            }
            else
            {
                // 如何遍历控件
                return false;
            }  
        }


        // 在确保文本框里的数据格式正确后创建一个可变参数对象并且返回之
        private AdjParameter CollectAdjParams()
        {
            return (new AdjParameter() { 
                TotalMoney = Convert.ToDouble(TbxTotalMoney.Text),
                RemMoneyMin = Convert.ToDouble(TbxRemMoneyMin.Text),
                RemMoneyMax = Convert.ToDouble(TbxRemMoneyMax.Text),
                KeyCountMin = Convert.ToInt32(TbxKeyCountMin.Text),
                KeyCountMax = Convert.ToInt32(TbxKeyCountMax.Text),
                CulCountMin = Convert.ToInt32(TbxCulCountMin.Text),
                CulCountMax = Convert.ToInt32(TbxCulCountMax.Text),
                KeyIntenMin = Convert.ToDouble(TbxKeyIntenMin.Text),
                KeyIntenMax = Convert.ToDouble(TbxKeyIntenMax.Text),
                CulIntenMax = Convert.ToDouble(TbxCulIntenMax.Text),
                CulIntenMin = Convert.ToDouble(TbxCulIntenMin.Text),
                YoungCountMin = Convert.ToInt32(TbxYoungCountMin.Text),
                YoungIntenMin = Convert.ToDouble(TbxYoungIntenMin.Text)
            });
        }


        private void HideInitParamWindow()
        {
            inFrame.Visibility = Visibility.Collapsed;
            TbxCulCount.Visibility = Visibility.Collapsed;
            TbxKeyCount.Visibility = Visibility.Collapsed;
            TbxDepName.Visibility = Visibility.Collapsed;
            TbxYoungCount.Visibility = Visibility.Collapsed;
            BtnAddParam.Visibility = Visibility.Collapsed;
            //BtnCloseAdjWindow.Visibility = Visibility.Collapsed;

            LbDepName.Visibility = Visibility.Collapsed;
            LbKeyAc.Visibility = Visibility.Collapsed;
            LbCulAc.Visibility = Visibility.Collapsed;
            LbYoungAc.Visibility = Visibility.Collapsed;
        }


        private bool AddToInitParam(int idx)
        {
            string newName = TbxDepName.Text.ToString();

            int newKeyAccept, newCulAccept, newYoungAccept;

            if (!IsNumber(TbxKeyCount.Text.ToString()))
            {
                newKeyAccept = 0;
            }
            else
            {
                newKeyAccept = Convert.ToInt32(TbxKeyCount.Text);
            }

            if (!IsNumber(TbxCulCount.Text.Trim()))
            {
                newCulAccept = 0;
            }
            else
            {
                newCulAccept = Convert.ToInt32(TbxCulCount.Text);
            }

            if (!IsNumber(TbxYoungCount.Text.ToString()))
            {
                newYoungAccept = 0;
            }
            else
            {
                newYoungAccept = Convert.ToInt32(TbxYoungCount.Text);
            }

            if (string.IsNullOrEmpty(TbxDepName.Text.ToString()))
            {
                //MessageBox.Show("至少填入一个学部名字");
                LbPleaseInput.Visibility = Visibility.Visible;
                return false;
            }
            else
            {

                LbPleaseInput.Visibility = Visibility.Collapsed;
                // 添加一个新的学部实例
                Department newDep = new Department()
                {
                    DepartmentName = newName,
                    KeyAcceptCount = newKeyAccept,
                    CulAcceptCount = newCulAccept,
                    YoungAcceptCount = newYoungAccept
                };
                if (idx == -1)
                {
                    initParam.AddDepartment(newDep);
                }
                else
                {
                    initParam.GetDataBanding()[idx] = newDep;
                }

                LstShowDeps.Items.Refresh();
                //LstShowDeps.Items.Add(newDep);
                return true;
            }
        }

        // 
        private bool isAlter;
        private void BtnAddParam_Click(object sender, RoutedEventArgs e)
        {

            //HideInitParamWindow();
            if (isAlter == true)
            {
                bool rst = AddToInitParam(LstShowDeps.Items.IndexOf(LstShowDeps.SelectedItem));
                if (rst == true)
                {
                    //HideInitParamWindow();
                    ShowInitParamWindow();
                    isAlter = false;
                }
            }
            else
            {
                bool rst = AddToInitParam(-1);
                ShowInitParamWindow();
                //if (rst == true)
                    //HideInitParamWindow();
            }
            
        }

        private void ShowInitParamWindow()
        {
            inFrame.Visibility = Visibility.Visible;
            inFrame.Content = "输入学部信息";
            TbxCulCount.Visibility = Visibility.Visible;
            TbxCulCount.Text = "";
            TbxKeyCount.Visibility = Visibility.Visible;
            TbxKeyCount.Text = "";
            TbxDepName.Visibility = Visibility.Visible;
            TbxDepName.Text = "";
            TbxYoungCount.Visibility = Visibility.Visible;
            TbxYoungCount.Text = "";

            BtnAddParam.Visibility = Visibility.Visible;
            BtnAddParam.Content = "增加";
            //BtnCloseAdjWindow.Visibility = Visibility.Visible;

            LbDepName.Visibility = Visibility.Visible;
            LbKeyAc.Visibility = Visibility.Visible;
            LbCulAc.Visibility = Visibility.Visible;
            LbYoungAc.Visibility = Visibility.Visible;
        }

        // 显示出不同于新增窗口的更改参数窗口
        private void ShowAdjustWindow()
        {
            inFrame.Visibility = Visibility.Visible;
            inFrame.Content = "修改学部信息";
            TbxCulCount.Visibility = Visibility.Visible;
            TbxCulCount.Text = ((Department)LstShowDeps.SelectedItem).CulAcceptCount.ToString();
            TbxKeyCount.Visibility = Visibility.Visible;
            TbxKeyCount.Text = ((Department)LstShowDeps.SelectedItem).KeyAcceptCount.ToString();
            TbxDepName.Visibility = Visibility.Visible;
            TbxDepName.Text = ((Department)LstShowDeps.SelectedItem).DepartmentName.ToString();
            TbxYoungCount.Visibility = Visibility.Visible;
            TbxYoungCount.Text = ((Department)LstShowDeps.SelectedItem).YoungAcceptCount.ToString();

            BtnAddParam.Content = "修改";
            BtnAddParam.Visibility = Visibility.Visible;
            //BtnCloseAdjWindow.Visibility = Visibility.Visible;

            LbDepName.Visibility = Visibility.Visible;
            LbKeyAc.Visibility = Visibility.Visible;
            LbCulAc.Visibility = Visibility.Visible;
            LbYoungAc.Visibility = Visibility.Visible;
        }



        private void BtnAddItem_Click(object sender, RoutedEventArgs e)
        {
            ShowInitParamWindow();
        }

        
        // 排序规则对象 Start
        private class SortByKeyRate : System.Collections.IComparer
        {
            public int Compare(Object x, Object y)
            {
                // return Convert.ToInt32(((DetailRecParam)((ArrayList)x)[((ArrayList)x).Count - 1]).KeyRate - ((DetailRecParam)((ArrayList)y)[((ArrayList)y).Count - 1]).KeyRate);
                return Convert.ToInt32(((DetailRecParam)((ArrayList)y)[0]).KeyRound - ((DetailRecParam)((ArrayList)x)[0]).KeyRound);
            }
        }

        private class SortByKeyIntensity : System.Collections.IComparer
        {
            public int Compare(Object x, Object y)
            {
                ArrayList arrX = (ArrayList)x;
                ArrayList arrY = (ArrayList)y;

                return Convert.ToInt32(((DetailRecParam)((ArrayList)y)[0]).RecParam.KeyInten - ((DetailRecParam)((ArrayList)x)[0]).RecParam.KeyInten);
            }
        }

        private class SortByRemMoney : System.Collections.IComparer
        {
            public int Compare(Object x, Object y)
            {
                return Convert.ToInt32(((DetailRecParam)((ArrayList)y)[((ArrayList)y).Count - 1]).Total - ((DetailRecParam)((ArrayList)x)[((ArrayList)x).Count - 1]).Total);

            }
        }
        // 排序规则对象 End


        // 更新主界面上推荐数字的显示，此时如果传来一个空数组那么一定会引发问题
        private void UpdateMainRecParam(ArrayList detArr)
        { 
            if (detArr.Count != 0)
            {
                RecParameter rec = ((DetailRecParam)detArr[0]).RecParam;
                LbKeyCount.Content = rec.KeyCount.ToString();
                LbKeyInten.Content = rec.KeyInten.ToString();
                LbCulCount.Content = rec.CulCount.ToString();
                LbCulInten.Content = rec.CulInten.ToString();
                LbYoungCount.Content = rec.YoungCount.ToString();
                LbYoungInten.Content = rec.YoungInten.ToString();


                Department dep = initParam.GetStatistic();

                double keyAcRate = Convert.ToDouble(rec.KeyCount) / dep.KeyAcceptCount;
                double culAcRate = Convert.ToDouble(rec.CulCount) / dep.CulAcceptCount;

                double youngAcRate;
                if (TbxYoungCountMin.Text == "0")
                {
                    youngAcRate = 0.0;
                }
                else
                {
                    youngAcRate = Convert.ToDouble(TbxYoungCountMin.Text) / dep.YoungAcceptCount;
                }
                LbKeyAcRate.Content = string.Format("{0:#.00%}", keyAcRate);
                LbCulAcRate.Content = string.Format("{0:#.00%}", culAcRate);
                LbYoungAcRate.Content = string.Format("{0:#.00%}", youngAcRate);

                LbRemMoney.Content = rec.RemMoney.ToString();
                double total = rec.CulCount * rec.CulInten + rec.KeyCount * rec.KeyInten + rec.YoungCount * rec.YoungInten;
                LbRecTotalMoney.Content = total.ToString();
            } 
            else
            {
                LbNoResult.Visibility = Visibility.Visible;
                //MessageBox.Show("无合适的方案，请重新设定参数并计算");
            }
        }

        private void SetZeroMainRecParam()
        {
            LbKeyCount.Content = "";
            LbKeyInten.Content = "";
            LbCulCount.Content = "";
            LbCulInten.Content = "";
            LbRecTotalMoney.Content = "";

            LbKeyAcRate.Content = "";
            LbCulAcRate.Content = "";
            LbYoungAcRate.Content = "";

            LbRemMoney.Content = "";
        }



        
        private void BtnCalculate_Click(object sender, RoutedEventArgs e)
        {
            if (initParam.GetDataBanding().Count == 0)
            {
                SetZeroMainRecParam();
                LstShowDetails.Items.Clear();
                return;
            }
            calculator = new Calculator();
            LbNoResult.Visibility = Visibility.Collapsed;
            if (ConfirmFillIn())
            {
                AdjParameter adjParam = CollectAdjParams();
                calculator.InitCalculator(initParam, adjParam);
                calculator.RunCalculator();                
                // 以后补充上多态的做法，现在暂时判断

                Calculator.Idx = 0;
                //LbNoResult.Visibility = Visibility.Collapsed;
                
                if (CbRemMoney.IsChecked == true)
                    calculator.SortUltimateRecArray(new SortByRemMoney());
                else if (CbKeyInten.IsChecked == true)
                    calculator.SortUltimateRecArray(new SortByKeyIntensity());
                else if (CbKeyRate.IsChecked == true)
                    calculator.SortUltimateRecArray(new SortByKeyRate());

                if (((ArrayList)calculator.GetUltimateArrayRef()).Count != 0)
                {
                    //MessageBox.Show(calculator.GetUltimateArrayRef().Count.ToString());
                    DisplayDetailsLst((ArrayList)calculator.GetUltimateArrayRef()[Calculator.Idx]);
                    UpdateMainRecParam((ArrayList)calculator.GetUltimateArrayRef()[Calculator.Idx]);
                }
                else
                {
                    LstShowDetails.Items.Clear();
                    LbNoResult.Visibility = Visibility.Visible;
                    SetZeroMainRecParam();
                    //MessageBox.Show("无合适的方案，Click");
                }
            }
        }

        private void BtnNextOne_Click(object sender, RoutedEventArgs e)
        {
            Calculator.Idx++;
            if (Calculator.Idx >= calculator.GetUltimateArrayRef().Count)
                Calculator.Idx = calculator.GetUltimateArrayRef().Count - 1;
            else
            {
                DisplayDetailsLst((ArrayList)calculator.GetUltimateArrayRef()[Calculator.Idx]);
                UpdateMainRecParam((ArrayList)calculator.GetUltimateArrayRef()[Calculator.Idx]);
            }
        }

        private void BtnPreOne_Click(object sender, RoutedEventArgs e)
        {
            Calculator.Idx--;
            if (Calculator.Idx < 0)
                Calculator.Idx = 0;
            else
            {
                DisplayDetailsLst((ArrayList)calculator.GetUltimateArrayRef()[Calculator.Idx]);
                UpdateMainRecParam((ArrayList)calculator.GetUltimateArrayRef()[Calculator.Idx]);
            }
        }
        
        private void TbxKeyCountMin_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.Key == Key.Up)
            {
                if (Convert.ToInt32(TbxKeyCountMin.Text) < Convert.ToInt32(TbxKeyCountMax.Text))
                {
                    TbxKeyCountMin.Text = (Convert.ToInt32(TbxKeyCountMin.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxKeyCountMin.Text) >= 2)
                {
                    TbxKeyCountMin.Text = (Convert.ToInt32(TbxKeyCountMin.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
           
        }
        
        private void TbxKeyCountMax_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                if ((Convert.ToInt32(TbxKeyCountMax.Text) + 1) * Convert.ToInt32(TbxKeyIntenMin.Text) < Convert.ToInt32(TbxTotalMoney.Text))
                {
                    TbxKeyCountMax.Text = (Convert.ToInt32(TbxKeyCountMax.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxKeyCountMin.Text) < Convert.ToInt32(TbxKeyCountMax.Text))
                {
                    TbxKeyCountMax.Text = (Convert.ToInt32(TbxKeyCountMax.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
        }

        private void TbxCulCountMin_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                if (Convert.ToInt32(TbxCulCountMin.Text) < Convert.ToInt32(TbxCulCountMax.Text))
                {
                    TbxCulCountMin.Text = (Convert.ToInt32(TbxCulCountMin.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxCulCountMin.Text) >= 2)
                {
                    TbxCulCountMin.Text = (Convert.ToInt32(TbxCulCountMin.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            //BtnCalculate_Click(sender, e);
        }

        private void TbxCulCountMax_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                if ((Convert.ToInt32(TbxCulCountMax.Text) + 1) * Convert.ToInt32(TbxCulIntenMin.Text) < Convert.ToInt32(TbxTotalMoney.Text))
                {
                    TbxCulCountMax.Text = (Convert.ToInt32(TbxCulCountMax.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxCulCountMin.Text) < Convert.ToInt32(TbxCulCountMax.Text))
                {
                    TbxCulCountMax.Text = (Convert.ToInt32(TbxCulCountMax.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            //BtnCalculate_Click(sender, e);
        }

        private void TbxKeyIntenMin_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            //BtnCalculate_Click(sender, e);
            if (e.Key == Key.Up)
            {
                if (Convert.ToInt32(TbxKeyIntenMin.Text) < Convert.ToInt32(TbxKeyIntenMax.Text))
                {
                    TbxKeyIntenMin.Text = (Convert.ToInt32(TbxKeyIntenMin.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxKeyIntenMin.Text) > 1)
                {
                    TbxKeyIntenMin.Text = (Convert.ToInt32(TbxKeyIntenMin.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
        }

        private void TbxKeyIntenMax_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                if (Convert.ToInt32(TbxKeyIntenMax.Text) * Convert.ToInt32(TbxKeyCountMin.Text) < 
                    Convert.ToInt32(TbxTotalMoney.Text))
                {
                    TbxKeyIntenMax.Text = (Convert.ToInt32(TbxKeyIntenMax.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxKeyIntenMin.Text) < Convert.ToInt32(TbxKeyIntenMax.Text))
                {
                    TbxKeyIntenMax.Text = (Convert.ToInt32(TbxKeyIntenMax.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }

            //BtnCalculate_Click(sender, e);
        }

        private void TbxCulIntenMin_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                if (Convert.ToInt32(TbxCulIntenMin.Text) < Convert.ToInt32(TbxCulIntenMax.Text))
                {
                    TbxCulIntenMin.Text = (Convert.ToInt32(TbxCulIntenMin.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxCulIntenMin.Text) > 1)
                {
                    TbxCulIntenMin.Text = (Convert.ToInt32(TbxCulIntenMin.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            
            //BtnCalculate_Click(sender, e);
        }

        private void TbxCulIntenMax_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                if (Convert.ToInt32(TbxKeyCountMin.Text) * Convert.ToInt32(TbxKeyIntenMin.Text) +
                    (Convert.ToInt32(TbxCulIntenMax.Text) + 1) < Convert.ToInt32(TbxTotalMoney.Text))
                {
                    TbxCulIntenMax.Text = (Convert.ToInt32(TbxCulIntenMax.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxCulIntenMin.Text) < Convert.ToInt32(TbxCulIntenMax.Text))
                {
                    TbxCulIntenMax.Text = (Convert.ToInt32(TbxCulIntenMax.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            //BtnCalculate_Click(sender, e);
        }

        private void TbxRemMoneyMin_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                if (Convert.ToInt32(TbxRemMoneyMin.Text) < Convert.ToInt32(TbxRemMoneyMax.Text))
                {
                    TbxRemMoneyMin.Text = (Convert.ToInt32(TbxRemMoneyMin.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxRemMoneyMin.Text) > 0)
                {
                    TbxRemMoneyMin.Text = (Convert.ToInt32(TbxRemMoneyMin.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);

                }
            }
        }

        private void TbxRemMoneyMax_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                if (Convert.ToInt32(TbxRemMoneyMax.Text) < Convert.ToInt32(TbxTotalMoney.Text))
                {
                    TbxRemMoneyMax.Text = (Convert.ToInt32(TbxRemMoneyMax.Text) + 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
            else if (e.Key == Key.Down)
            {
                if (Convert.ToInt32(TbxRemMoneyMin.Text) < Convert.ToInt32(TbxRemMoneyMax.Text))
                {
                    TbxRemMoneyMax.Text = (Convert.ToInt32(TbxRemMoneyMax.Text) - 1).ToString();
                    BtnCalculate_Click(sender, e);
                }
            }
        }

        private ArrayList resultArray;
        private void BtnAddResult_Click(object sender, RoutedEventArgs e)
        {
            if (calculator.GetUltimateArrayRef().Count == 0) return;
            resultArray.Add((ArrayList)calculator.GetUltimateArrayRef()[Calculator.Idx]);
            MessageBox.Show("该方案已经添加到结果队列");
        }

        private void BtnWriteToFile_Click(object sender, RoutedEventArgs e)
        {
            if (resultArray.Count == 0) return;
            if (initParam.GetStatistic().YoungAcceptCount == 0)
            {
                WriteToFileWithNoYoung();
            }
            else
            {
                WriteToFileWithYoung();
            }
            resultArray.Clear();
        }

        private void WriteToFileWithYoung()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Result");
            /*
            for (int i = 1; i < 15; i++)
                sheet.SetColumnWidth(i, 25 * 256);
            */
            for (int i = 0; i < 3; i++)
                sheet.SetColumnWidth(i, 20 * 256);

            int rstCount = ((ArrayList)resultArray[0]).Count;
            for (int i = 3; i < rstCount + 3; i++)
                sheet.SetColumnWidth(i, 25 * 256);

            sheet.SetColumnWidth(3 + rstCount, 20 * 256);

            int offset = 30;

            for (int i = 0; i <= resultArray.Count * offset; i++)
            {
                sheet.CreateRow(i);
            }

            ICellStyle styleCul = workbook.CreateCellStyle();
            styleCul.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleCul.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;
            styleCul.FillPattern = FillPattern.SolidForeground;
            styleCul.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;
            styleCul.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            styleCul.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;

            ICellStyle styleKey = workbook.CreateCellStyle();
            styleKey.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleKey.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick; ;
            styleKey.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;




            ICellStyle styleHeader = workbook.CreateCellStyle();
            styleHeader.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleHeader.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            IFont fontBold = workbook.CreateFont();
            fontBold.Boldweight = (short)FontBoldWeight.Bold;
            styleHeader.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            styleHeader.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            styleHeader.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            styleHeader.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            styleHeader.WrapText = true;
            styleHeader.SetFont(fontBold);

            ICellStyle styleBold = workbook.CreateCellStyle();
            styleBold.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleBold.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            styleBold.SetFont(fontBold);


            ICellStyle styleRateCell = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            styleRateCell.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
            styleRateCell.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            styleRateCell.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            styleRateCell.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleRateCell.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            styleRateCell.SetFont(fontBold);
            styleRateCell.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;
            styleRateCell.FillPattern = FillPattern.SolidForeground;
            styleRateCell.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;


            ICellStyle styleRateCellRed = workbook.CreateCellStyle();
            styleRateCellRed.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
            styleRateCellRed.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleRateCellRed.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            styleRateCellRed.SetFont(fontBold);
            styleRateCellRed.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            styleRateCellRed.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            IFont fontBoldRed = workbook.CreateFont();
            fontBoldRed.Boldweight = (short)FontBoldWeight.Bold;
            fontBoldRed.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
            styleRateCellRed.SetFont(fontBoldRed);


            ICellStyle styleKeyRed = workbook.CreateCellStyle();
            styleKeyRed.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            styleKeyRed.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            styleKeyRed.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleKeyRed.SetFont(fontBoldRed);

            // 将所有方案写出到文件，i 代表第 i 个方案
            for (int i = 1; i <= resultArray.Count; i++)
            {

                IRow r = sheet.GetRow((i - 1) * offset + 2);
                ICell c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("申请科学部");

                r = sheet.GetRow((i - 1) * offset);
                c = r.CreateCell(1);
                c.CellStyle = styleBold;
                c.SetCellValue("2016年NSFC-XX联合基金建议资助方案");

                r = sheet.GetRow((i - 1) * offset + 1);
                c = r.CreateCell(1);
                c.CellStyle = styleBold;
                ArrayList one = (ArrayList)resultArray[i - 1];
                DetailRecParam dets = (DetailRecParam)one[one.Count - 1];
                DetailRecParam det1 = (DetailRecParam)one[0];
                c.SetCellValue("（2016年度计划安排直接费用" + (dets.Total + det1.RecParam.RemMoney).ToString() + "万元）");


                r = sheet.GetRow((i - 1) * offset + 3);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("受理申请");

                r = sheet.GetRow((i - 1) * offset + 13);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("办公室建议批准(项,万元)");

                r = sheet.GetRow((i - 1) * offset + 3);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("申请小计");

                r = sheet.GetRow((i - 1) * offset + 4);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育申请数");

                r = sheet.GetRow((i - 1) * offset + 5);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持申请数");

                r = sheet.GetRow((i - 1) * offset + 6);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育初审数");

                r = sheet.GetRow((i - 1) * offset + 7);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持初审数");



                r = sheet.GetRow((i - 1) * offset + 8);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育受理数");

                r = sheet.GetRow((i - 1) * offset + 9);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育所占百分比(%)");

                r = sheet.GetRow((i - 1) * offset + 10);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持受理数");

                r = sheet.GetRow((i - 1) * offset + 11);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点所占百分比(%)");

                r = sheet.GetRow((i - 1) * offset + 12);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("本地优青");


                r = sheet.GetRow((i - 1) * offset + 13);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育(计算)");

                r = sheet.GetRow((i - 1) * offset + 14);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育(取整)");

                r = sheet.GetRow((i - 1) * offset + 15);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持(计算)");

                r = sheet.GetRow((i - 1) * offset + 16);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持(取整)");

                r = sheet.GetRow((i - 1) * offset + 17);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("本地优青(计算)");

                r = sheet.GetRow((i - 1) * offset + 18);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("本地优青(取整)");


                r = sheet.GetRow((i - 1) * offset + 19);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("经费指标");


                r = sheet.GetRow((i - 1) * offset + 20);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("办公室建议会评项目(项)");

                r = sheet.GetRow((i - 1) * offset + 22);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("联席会建议会评项目(项)");

                r = sheet.GetRow((i - 1) * offset + 24);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("联席会建议批准(项,万元)");

                r = sheet.GetRow((i - 1) * offset + 20);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育");

                r = sheet.GetRow((i - 1) * offset + 21);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持");

                r = sheet.GetRow((i - 1) * offset + 22);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育");

                r = sheet.GetRow((i - 1) * offset + 23);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持");

                r = sheet.GetRow((i - 1) * offset + 24);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育");

                r = sheet.GetRow((i - 1) * offset + 25);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持");

                r = sheet.GetRow((i - 1) * offset + 26);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("经费指标");




                ArrayList oneResult = (ArrayList)resultArray[i - 1];
                for (int j = 1; j <= oneResult.Count; j++)
                {
                    DetailRecParam det = (DetailRecParam)oneResult[j - 1];
                    IRow row = sheet.GetRow((i - 1) * offset + 2);
                    ICell cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleHeader;
                    cell.SetCellValue(det.DepartmentName);

                    row = sheet.GetRow((i - 1) * offset + 3);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 4);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 5);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 6);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 7);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;



                    row = sheet.GetRow((i - 1) * offset + 8);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;
                    cell.SetCellValue(det.CulAcceptCount);

                    row = sheet.GetRow((i - 1) * offset + 9);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleRateCell;
                    cell.SetCellValue(det.CulRate);

                    row = sheet.GetRow((i - 1) * offset + 10);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;
                    cell.SetCellValue(det.KeyAcceptCount);

                    row = sheet.GetRow((i - 1) * offset + 11);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleRateCellRed;
                    cell.SetCellValue(det.KeyRate);

                    row = sheet.GetRow((i - 1) * offset + 12);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;
                    cell.SetCellValue(det.YoungAcceptCount);



                    row = sheet.GetRow((i - 1) * offset + 13);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;
                    cell.SetCellValue(det.CulCaled);

                    row = sheet.GetRow((i - 1) * offset + 14);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;
                    cell.SetCellValue(det.CulRound);

                    row = sheet.GetRow((i - 1) * offset + 15);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKeyRed;
                    cell.SetCellValue(det.KeyCaled);

                    row = sheet.GetRow((i - 1) * offset + 16);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;
                    cell.SetCellValue(det.KeyRound);

                    row = sheet.GetRow((i - 1) * offset + 17);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;
                    cell.SetCellValue(det.YoungCaled);

                    row = sheet.GetRow((i - 1) * offset + 18);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;
                    cell.SetCellValue(det.YoungRound);


                    row = sheet.GetRow((i - 1) * offset + 19);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleHeader;
                    cell.SetCellValue(det.Total);

                    row = sheet.GetRow((i - 1) * offset + 20);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 21);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 22);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 23);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 24);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 25);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 26);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleHeader;
                }

                for (int j = (i - 1) * offset + 4; j <= (i - 1) * offset + 12; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }
                for (int j = (i - 1) * offset + 14; j <= (i - 1) * offset + 19; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }
                for (int j = (i - 1) * offset + 21; j <= (i - 1) * offset + 21; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }
                for (int j = (i - 1) * offset + 23; j <= (i - 1) * offset + 23; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }
                for (int j = (i - 1) * offset + 25; j <= (i - 1) * offset + 26; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }

                sheet.GetRow((i - 1) * offset + 2).CreateCell(2).CellStyle = styleHeader;



                IRow ir = sheet.GetRow((i - 1) * offset + 27);
                ICell ic = ir.CreateCell(1);
                ic.CellStyle = styleBold;
                ic.SetCellValue("注：1、指南：重点支持项目平均资助强度为XXX万元/项；培育项目平均资助强度为XX万元/项。");

                ir = sheet.GetRow((i - 1) * offset + 28);
                ic = ir.CreateCell(1);
                ic.CellStyle = styleKeyRed;
                ic.SetCellValue("2、考虑：培育：" + ((DetailRecParam)oneResult[0]).RecParam.CulInten.ToString() +
                    "万元/项；重点支持：" + ((DetailRecParam)oneResult[0]).RecParam.KeyInten.ToString() + "万元/项；剩余：" +
                     ((DetailRecParam)oneResult[0]).RecParam.RemMoney.ToString() + "万元");

                // 下方说明1
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 27, (i - 1) * offset + 27, 1, 4));
                // 下方说明2
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 28, (i - 1) * offset + 28, 1, 4));

                // 受理申请
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 3, (i - 1) * offset + 12, 1, 1));
                // 办公室
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 13, (i - 1) * offset + 19, 1, 1));
                // 申请科学部
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 2, (i - 1) * offset + 2, 1, 2));
                // 抬头
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset, (i - 1) * offset, 1, 5));
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 1, (i - 1) * offset + 1, 1, 5));


                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 20, (i - 1) * offset + 21, 1, 1));
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 22, (i - 1) * offset + 23, 1, 1));
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 24, (i - 1) * offset + 26, 1, 1));

            }

            Microsoft.Win32.SaveFileDialog FileDialog = new Microsoft.Win32.SaveFileDialog();
            FileDialog.Filter = "Excel(.xls)|*.xls";
            FileDialog.ShowDialog();
            if (!FileDialog.FileName.Equals(""))
            {


                FileStream sw = File.Create(FileDialog.FileName);
                workbook.Write(sw);
                sw.Close();

                //MessageBox.Show("OK");
            }
        }

        private void WriteToFileWithNoYoung()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Result");
            /*
            for (int i = 1; i < 15; i++)
                sheet.SetColumnWidth(i, 25 * 256);
            */
            for (int i = 0; i < 3; i++)
                sheet.SetColumnWidth(i, 20 * 256);

            int rstCount = ((ArrayList)resultArray[0]).Count;
            for (int i = 3; i < rstCount + 3; i++)
                sheet.SetColumnWidth(i, 25 * 256);

            sheet.SetColumnWidth(3 + rstCount, 20 * 256);

            int offset = 30;

            for (int i = 0; i <= resultArray.Count * offset; i++)
            {
                sheet.CreateRow(i);
            }

            ICellStyle styleCul = workbook.CreateCellStyle();
            styleCul.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleCul.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;
            styleCul.FillPattern = FillPattern.SolidForeground;
            styleCul.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;
            styleCul.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            styleCul.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;

            ICellStyle styleKey = workbook.CreateCellStyle();
            styleKey.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleKey.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick; ;
            styleKey.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;

            


            ICellStyle styleHeader = workbook.CreateCellStyle();
            styleHeader.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleHeader.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            IFont fontBold = workbook.CreateFont();
            fontBold.Boldweight = (short)FontBoldWeight.Bold;
            styleHeader.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            styleHeader.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            styleHeader.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            styleHeader.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            styleHeader.WrapText = true;
            styleHeader.SetFont(fontBold);

            ICellStyle styleBold = workbook.CreateCellStyle();
            styleBold.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleBold.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            styleBold.SetFont(fontBold);


            ICellStyle styleRateCell = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            styleRateCell.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
            styleRateCell.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            styleRateCell.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            styleRateCell.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleRateCell.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            styleRateCell.SetFont(fontBold);
            styleRateCell.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;
            styleRateCell.FillPattern = FillPattern.SolidForeground;
            styleRateCell.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;


            ICellStyle styleRateCellRed = workbook.CreateCellStyle();
            styleRateCellRed.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
            styleRateCellRed.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleRateCellRed.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            styleRateCellRed.SetFont(fontBold);
            styleRateCellRed.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            styleRateCellRed.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            IFont fontBoldRed = workbook.CreateFont();
            fontBoldRed.Boldweight = (short)FontBoldWeight.Bold;
            fontBoldRed.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
            styleRateCellRed.SetFont(fontBoldRed);


            ICellStyle styleKeyRed = workbook.CreateCellStyle();
            styleKeyRed.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            styleKeyRed.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            styleKeyRed.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleKeyRed.SetFont(fontBoldRed);

            // 将所有方案写出到文件，i 代表第 i 个方案
            for (int i = 1; i <= resultArray.Count; i++)
            {

                IRow r = sheet.GetRow((i - 1) * offset + 2);
                ICell c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("申请科学部");

                r = sheet.GetRow((i - 1) * offset);
                c = r.CreateCell(1);
                c.CellStyle = styleBold;
                c.SetCellValue("2016年NSFC-XX联合基金建议资助方案");

                r = sheet.GetRow((i - 1) * offset + 1);
                c = r.CreateCell(1);
                c.CellStyle = styleBold;
                ArrayList one = (ArrayList)resultArray[i - 1];
                DetailRecParam dets = (DetailRecParam)one[one.Count - 1];
                DetailRecParam det1 = (DetailRecParam)one[0];
                c.SetCellValue("（2016年度计划安排直接费用" + (dets.Total + det1.RecParam.RemMoney).ToString() + "万元）");



                r = sheet.GetRow((i - 1) * offset + 3);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("受理申请");

                r = sheet.GetRow((i - 1) * offset + 12);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("办公室建议批准(项,万元)");

                r = sheet.GetRow((i - 1) * offset + 3);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("申请小计");

                r = sheet.GetRow((i - 1) * offset + 4);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育申请数");

                r = sheet.GetRow((i - 1) * offset + 5);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持申请数");

                r = sheet.GetRow((i - 1) * offset + 6);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育初审数");

                r = sheet.GetRow((i - 1) * offset + 7);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持初审数");



                r = sheet.GetRow((i - 1) * offset + 8);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育受理数");

                r = sheet.GetRow((i - 1) * offset + 9);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育所占百分比(%)");

                r = sheet.GetRow((i - 1) * offset + 10);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持受理数");

                r = sheet.GetRow((i - 1) * offset + 11);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点所占百分比(%)");


                r = sheet.GetRow((i - 1) * offset + 12);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育(计算)");

                r = sheet.GetRow((i - 1) * offset + 13);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育(取整)");

                r = sheet.GetRow((i - 1) * offset + 14);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持(计算)");

                r = sheet.GetRow((i - 1) * offset + 15);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持(取整)");

                r = sheet.GetRow((i - 1) * offset + 16);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("经费指标");


                r = sheet.GetRow((i - 1) * offset + 17);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("办公室建议会评项目(项)");

                r = sheet.GetRow((i - 1) * offset + 19);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("联席会建议会评项目(项)");

                r = sheet.GetRow((i - 1) * offset + 21);
                c = r.CreateCell(1);
                c.CellStyle = styleHeader;
                c.SetCellValue("联席会建议批准(项,万元)");

                r = sheet.GetRow((i - 1) * offset + 17);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育");

                r = sheet.GetRow((i - 1) * offset + 18);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持");

                r = sheet.GetRow((i - 1) * offset + 19);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育");

                r = sheet.GetRow((i - 1) * offset + 20);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持");

                r = sheet.GetRow((i - 1) * offset + 21);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("培育");

                r = sheet.GetRow((i - 1) * offset + 22);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("重点支持");

                r = sheet.GetRow((i - 1) * offset + 23);
                c = r.CreateCell(2);
                c.CellStyle = styleHeader;
                c.SetCellValue("经费指标");




                ArrayList oneResult = (ArrayList)resultArray[i - 1];
                for (int j = 1; j <= oneResult.Count; j++)
                {
                    DetailRecParam det = (DetailRecParam)oneResult[j - 1];
                    IRow row = sheet.GetRow((i - 1) * offset + 2);
                    ICell cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleHeader;
                    cell.SetCellValue(det.DepartmentName);

                    row = sheet.GetRow((i - 1) * offset + 3);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 4);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 5);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 6);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 7);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;



                    row = sheet.GetRow((i - 1) * offset + 8);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;
                    cell.SetCellValue(det.CulAcceptCount);

                    row = sheet.GetRow((i - 1) * offset + 9);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleRateCell;
                    cell.SetCellValue(det.CulRate);

                    row = sheet.GetRow((i - 1) * offset + 10);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;
                    cell.SetCellValue(det.KeyAcceptCount);

                    row = sheet.GetRow((i - 1) * offset + 11);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleRateCellRed;
                    cell.SetCellValue(det.KeyRate);

                    row = sheet.GetRow((i - 1) * offset + 12);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;
                    cell.SetCellValue(det.CulCaled);

                    row = sheet.GetRow((i - 1) * offset + 13);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;
                    cell.SetCellValue(det.CulRound);

                    row = sheet.GetRow((i - 1) * offset + 14);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKeyRed;
                    cell.SetCellValue(det.KeyCaled);

                    row = sheet.GetRow((i - 1) * offset + 15);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;
                    cell.SetCellValue(det.KeyRound);

                    row = sheet.GetRow((i - 1) * offset + 16);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleHeader;
                    cell.SetCellValue(det.Total);

                    row = sheet.GetRow((i - 1) * offset + 17);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 18);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 19);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 20);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 21);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleCul;

                    row = sheet.GetRow((i - 1) * offset + 22);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleKey;

                    row = sheet.GetRow((i - 1) * offset + 23);
                    cell = row.CreateCell(j + 2);
                    cell.CellStyle = styleHeader;



                }

                for (int j = (i - 1) * offset + 4; j <= (i - 1) * offset + 11; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }
                for (int j = (i - 1) * offset + 13; j <= (i - 1) * offset + 16; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }
                for (int j = (i - 1) * offset + 18; j <= (i - 1) * offset + 18; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }
                for (int j = (i - 1) * offset + 20; j <= (i - 1) * offset + 20; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }
                for (int j = (i - 1) * offset + 22; j <= (i - 1) * offset + 23; j++)
                {
                    sheet.GetRow(j).CreateCell(1).CellStyle = styleHeader;
                }

                sheet.GetRow((i - 1) * offset + 2).CreateCell(2).CellStyle = styleHeader;



                IRow ir = sheet.GetRow((i - 1) * offset + 24);
                ICell ic = ir.CreateCell(1);
                ic.CellStyle = styleBold;
                ic.SetCellValue("注：1、指南：重点支持项目平均资助强度为XXX万元/项；培育项目平均资助强度为XX万元/项。");

                ir = sheet.GetRow((i - 1) * offset + 25);
                ic = ir.CreateCell(1);
                ic.CellStyle = styleKeyRed;
                ic.SetCellValue("2、考虑：培育：" + ((DetailRecParam)oneResult[0]).RecParam.CulInten.ToString() +
                    "万元/项；重点支持：" + ((DetailRecParam)oneResult[0]).RecParam.KeyInten.ToString() + "万元/项；剩余：" +
                     ((DetailRecParam)oneResult[0]).RecParam.RemMoney.ToString() + "万元");

                // 下方说明1
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 24, (i - 1) * offset + 24, 1, 4));
                // 下方说明2
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 25, (i - 1) * offset + 25, 1, 4));

                // 受理申请
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 3, (i - 1) * offset + 11, 1, 1));
                // 办公室
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 12, (i - 1) * offset + 16, 1, 1));
                // 申请科学部
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 2, (i - 1) * offset + 2, 1, 2));
                // 抬头
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset, (i - 1) * offset, 1, 5));
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 1, (i - 1) * offset + 1, 1, 5));


                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 17, (i - 1) * offset + 18, 1, 1));
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 19, (i - 1) * offset + 20, 1, 1));
                sheet.AddMergedRegion(new CellRangeAddress((i - 1) * offset + 21, (i - 1) * offset + 23, 1, 1));

            }

            Microsoft.Win32.SaveFileDialog FileDialog = new Microsoft.Win32.SaveFileDialog();
            FileDialog.Filter = "Excel(.xls)|*.xls";
            FileDialog.ShowDialog();
            if (!FileDialog.FileName.Equals(""))
            {


                FileStream sw = File.Create(FileDialog.FileName);
                workbook.Write(sw);
                sw.Close();

                //MessageBox.Show("OK");
            }
        }

        private void BtnDeleteItem_Click(object sender, RoutedEventArgs e)
        {
            if (initParam.GetDataBanding().Count != 0 && LstShowDeps.Items.IndexOf(LstShowDeps.SelectedItem) != -1)
            {
                //MessageBox.Show(LstShowDeps.Items.IndexOf(LstShowDeps.SelectedItem).ToString());
                //LstShowDeps.Items.RemoveAt(LstShowDeps.Items.IndexOf(LstShowDeps.SelectedItem));
                initParam.GetDataBanding().RemoveAt(LstShowDeps.Items.IndexOf(LstShowDeps.SelectedItem));
                LstShowDeps.Items.Refresh();
                //MessageBox.Show(initParam.GetDataBanding().Count.ToString());
            }
            else
            {
                MessageBox.Show("请先选中其中一个学部，再按删除");
            }
        }

        private void LstShowDeps_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (LstShowDeps.SelectedItems.Count > 0)
            {
                ShowAdjustWindow();
                isAlter = true;
            }
        }

        private void BtnCloseAdjWindow_Click(object sender, RoutedEventArgs e)
        {
            HideInitParamWindow();
        }

        // 清空所有
        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {
            //LstShowDeps.Items.Clear();
            initParam.GetDataBanding().Clear();
            LstShowDeps.Items.Refresh();
        }


        private void CbRemMoney_Click(object sender, RoutedEventArgs e)
        {
            BtnCalculate_Click(sender, e);
        }

        private void CbKeyInten_Click(object sender, RoutedEventArgs e)
        {
            BtnCalculate_Click(sender, e);
        }

        private void CbKeyRate_Click(object sender, RoutedEventArgs e)
        {
            BtnCalculate_Click(sender, e);
        }

        private void LoadParams()
        {
            string currentPath = Directory.GetCurrentDirectory();

            string file = currentPath + "\\params.ini";
            INIFile readINI = new INIFile(file);

            if (!File.Exists(file))
            {
                MessageBox.Show("上次未设定参数");
                return;
            }

            int depCount = Convert.ToInt32(readINI.ReadIniData("DepartmentCount", "Count"));

            initParam.GetDataBanding().Clear();
            for (int i = 1; i <= depCount; i++)
            {
                initParam.GetDataBanding().Add(new Department() {
                    DepartmentName = readINI.ReadIniData("InitParams", "Department" + i.ToString()),
                    CulAcceptCount = Convert.ToInt32(readINI.ReadIniData("InitParams", "CulAcCount" + i.ToString())),
                    KeyAcceptCount = Convert.ToInt32(readINI.ReadIniData("InitParams", "KeyAcCount" + i.ToString())),
                    YoungAcceptCount = Convert.ToInt32(readINI.ReadIniData("InitParams", "YoungAcCount" + i.ToString()))
                });
            }
            LstShowDeps.Items.Refresh();

            TbxTotalMoney.Text = readINI.ReadIniData("AdjParams", "Total");
            TbxRemMoneyMin.Text = readINI.ReadIniData("AdjParams", "RemMoneyMin");
            TbxRemMoneyMax.Text = readINI.ReadIniData("AdjParams", "RemMoneyMax");
            TbxKeyCountMin.Text = readINI.ReadIniData("AdjParams", "KeyCountMin");
            TbxKeyCountMax.Text = readINI.ReadIniData("AdjParams", "KeyCountMax");
            TbxCulCountMin.Text = readINI.ReadIniData("AdjParams", "CulCountMin");
            TbxCulCountMax.Text = readINI.ReadIniData("AdjParams", "CulCountMax");
            TbxKeyIntenMin.Text = readINI.ReadIniData("AdjParams", "KeyIntenMin");
            TbxKeyIntenMax.Text = readINI.ReadIniData("AdjParams", "KeyIntenMax");
            TbxCulIntenMin.Text = readINI.ReadIniData("AdjParams", "CulIntenMin");
            TbxCulIntenMax.Text = readINI.ReadIniData("AdjParams", "CulIntenMax");
            TbxYoungCountMin.Text = readINI.ReadIniData("AdjParams", "YoungCountMin");
            TbxYoungIntenMin.Text = readINI.ReadIniData("AdjParams", "YoungIntenMin");

        }

        private void SaveParams()
        {
            string currentPath = Directory.GetCurrentDirectory();

            string file = currentPath + "\\params.ini";
            INIFile saveINI = new INIFile(file);
            
            if (!File.Exists(file))
            {
                File.Create(file);
            }

            int depCount = initParam.GetDataBanding().Count;
            ArrayList depArray = initParam.GetDataBanding();
            saveINI.WriteIniData("DepartmentCount", "Count", depCount.ToString());

            for (int i = 1; i <= depCount; i++)
            {
                Department dep = (Department)depArray[i - 1];
                saveINI.WriteIniData("InitParams", "Department" + i.ToString(), dep.DepartmentName);
                saveINI.WriteIniData("InitParams", "CulAcCount" + i.ToString(), dep.CulAcceptCount.ToString());
                saveINI.WriteIniData("InitParams", "KeyAcCount" + i.ToString(), dep.KeyAcceptCount.ToString());
                saveINI.WriteIniData("InitParams", "YoungAcCount" + i.ToString(), dep.YoungAcceptCount.ToString());
            }
            AdjParameter adjParams = CollectAdjParams();
            saveINI.WriteIniData("AdjParams", "Total", adjParams.TotalMoney.ToString());
            saveINI.WriteIniData("AdjParams", "RemMoneyMin", adjParams.RemMoneyMin.ToString());
            saveINI.WriteIniData("AdjParams", "RemMoneyMax", adjParams.RemMoneyMax.ToString());
            saveINI.WriteIniData("AdjParams", "KeyCountMin", adjParams.KeyCountMin.ToString());
            saveINI.WriteIniData("AdjParams", "KeyCountMax", adjParams.KeyCountMax.ToString());
            saveINI.WriteIniData("AdjParams", "CulCountMin", adjParams.CulCountMin.ToString());
            saveINI.WriteIniData("AdjParams", "CulCountMax", adjParams.CulCountMax.ToString());
            saveINI.WriteIniData("AdjParams", "KeyIntenMin", adjParams.KeyIntenMin.ToString());
            saveINI.WriteIniData("AdjParams", "KeyIntenMax", adjParams.KeyIntenMax.ToString());
            saveINI.WriteIniData("AdjParams", "CulIntenMin", adjParams.CulIntenMin.ToString());
            saveINI.WriteIniData("AdjParams", "CulIntenMax", adjParams.CulIntenMax.ToString());
            saveINI.WriteIniData("AdjParams", "YoungCountMin", adjParams.YoungCountMin.ToString());
            saveINI.WriteIniData("AdjParams", "YoungIntenMin", adjParams.YoungIntenMin.ToString());

            MessageBox.Show("参数保存成功");
        }

        private bool ConfirmAllRightToSaveParam()
        {
            /*
            if (TbxDepName.Text == "") return false;
            if (!IsNumber(TbxKeyCount.Text)) return false;
            if (!IsNumber(TbxCulCount.Text)) return false;
            if (!IsNumber(TbxYoungCount.Text)) return false;
            */

            if (IsNumber(TbxTotalMoney.Text.ToString()) && IsNumber(TbxRemMoneyMin.Text.ToString()) &&
                IsNumber(TbxRemMoneyMax.Text.ToString()) && IsNumber(TbxKeyCountMin.Text.ToString()) &&
                IsNumber(TbxKeyCountMax.Text.ToString()) && IsNumber(TbxCulCountMin.Text.ToString()) &&
                IsNumber(TbxCulCountMax.Text.ToString()) && IsNumber(TbxKeyIntenMin.Text.ToString()) &&
                IsNumber(TbxKeyIntenMax.Text.ToString()) && IsNumber(TbxCulIntenMin.Text.ToString()) &&
                IsNumber(TbxCulIntenMax.Text.ToString())
                )
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        private void BtnSaveParams_Click(object sender, RoutedEventArgs e)
        {
            if (ConfirmAllRightToSaveParam())
                SaveParams();
            else
                MessageBox.Show("需要所有参数不为空才可保存");
        }

        private void BtnLoadParams_Click(object sender, RoutedEventArgs e)
        {
            LoadParams();
        }

        private void BtnAdjItem_Click(object sender, RoutedEventArgs e)
        {
            if (LstShowDeps.SelectedItems.Count > 0)
            {
                ShowAdjustWindow();
                isAlter = true;
            }
        }


    }
}
