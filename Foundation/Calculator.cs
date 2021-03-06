﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Diagnostics;

using System.Windows;

namespace Foundation
{
    public sealed class Calculator
    {
        private ArrayList ultimateRecArray;

        public static int Idx = 0;

        private InitParameters initParameters;
        private AdjParameter adjParameters;
        private RecParamArray recParamArray;



        public Calculator()
        {
            recParamArray = new RecParamArray();
            ultimateRecArray = new ArrayList();
            arrLevel2Aux = new ArrayList();
            arrLevel3Aux = new ArrayList();
            arrLevel4Aux = new ArrayList();
        }

        public ArrayList GetUltimateArrayRef()
        {
            return ultimateRecArray;
        }

        public void SortUltimateRecArray(IComparer comparer)
        {
            ultimateRecArray.Sort(comparer);
        }


        public void InitCalculator(InitParameters newInitParam, AdjParameter newAdjParam)
        {
            
            initParameters = newInitParam;
            adjParameters = newAdjParam;
            recParamArray.ClearArray();
            ultimateRecArray.Clear();
        }

        private void RefreshUltimateArray()
        {

            for (int i = 0; i < recParamArray.GetRecParamsCount(); i++ )
            {
                ArrayList newList = new ArrayList();

                double totalCulAccept, totalKeyAccept, totalYoungAccept, culRate, keyRate, youngRate, culCaled, keyCaled, youngCaled;
                double culFundingRate, keyFundingRate, youngFundingRate;
                int culAc, keyAc, culRound, keyRound, youngAc, youngRound;

                for (int j = 0; j < initParameters.DepartmentCount; j++)
                {
                    totalCulAccept = initParameters.GetStatistic().CulAcceptCount;
                    totalKeyAccept = initParameters.GetStatistic().KeyAcceptCount;
                    totalYoungAccept = initParameters.GetStatistic().YoungAcceptCount;

                    culAc = ((Department)(initParameters.GetDataBanding()[j])).CulAcceptCount;
                    keyAc = ((Department)(initParameters.GetDataBanding()[j])).KeyAcceptCount;
                    youngAc = ((Department)(initParameters.GetDataBanding()[j])).YoungAcceptCount;

                    

                    if (totalCulAccept == 0.0)
                    {
                        culRate = 0.0;
                    }
                    else
                    {
                        culRate = culAc * 1.0 / totalCulAccept;
                    }
                    // ===========================================
                    if (totalKeyAccept == 0.0)
                    {
                        keyRate = 0.0;
                    }
                    else
                    {
                        keyRate = keyAc * 1.0 / totalKeyAccept;
                    }
                    // ===========================================
                    if (totalYoungAccept == 0.0)
                    {
                        youngRate = 0.0;
                    }
                    else
                    {
                        youngRate = youngAc * 1.0 / totalYoungAccept;
                    }


                    culCaled = ((RecParameter)recParamArray.GetRecArrayRef()[i]).CulCount * 1.0 * culRate;
                    keyCaled = ((RecParameter)recParamArray.GetRecArrayRef()[i]).KeyCount * 1.0 * keyRate;
                    youngCaled = ((RecParameter)recParamArray.GetRecArrayRef()[i]).YoungCount * 1.0 * youngRate;


                    youngRound = (int)Math.Round(youngCaled);
                    culRound = (int)Math.Round(culCaled);
                    keyRound = (int)Math.Round(keyCaled);

                    if (totalCulAccept != 0.0) culFundingRate = culRound * 1.0 / totalCulAccept; else culFundingRate = 0.0;
                    if (totalKeyAccept != 0.0) keyFundingRate = keyRound * 1.0 / totalKeyAccept; else keyFundingRate = 0.0;
                    if (totalYoungAccept != 0.0) youngFundingRate = youngRound * 1.0 / totalYoungAccept; else youngFundingRate = 0.0;


                    DetailRecParam detailParam = new DetailRecParam() { 
                        DepartmentName = ((Department)(initParameters.GetDataBanding()[j])).DepartmentName,
                        CulAcceptCount = culAc, KeyAcceptCount = keyAc, YoungAcceptCount = youngAc,
                        CulRate = culRate, KeyRate = keyRate, YoungRate = youngRate,
                         
                        CulCaled = culCaled, KeyCaled = keyCaled, YoungCaled = youngCaled,
                        CulRound = culRound, KeyRound = keyRound, YoungRound = youngRound,
                        Total = (((RecParameter)recParamArray.GetRecArrayRef()[i]).CulInten * culRound +
                                    ((RecParameter)recParamArray.GetRecArrayRef()[i]).KeyInten * keyRound) +
                                    ((RecParameter)recParamArray.GetRecArrayRef()[i]).YoungInten * youngRound,
                        // ====================== Extended ==========================
                        CulFundingRate = culFundingRate,
                        KeyFundingRate = keyFundingRate,
                        YoungFundingRate = youngFundingRate,
                        // ====================== Extended ==========================
                        RecParam = (RecParameter)recParamArray.GetRecArrayRef()[i]
                    };
                    // MessageBox.Show(detailParam.RecParam.CulInten.ToString());
                    newList.Add(detailParam);
                }
                culAc = keyAc = youngAc = culRound = keyRound = youngRound = 0;
                totalCulAccept = totalKeyAccept = totalYoungAccept = culRate = keyRate = youngRate = culCaled = keyCaled = youngCaled = 0.0;
                culFundingRate = keyFundingRate = youngFundingRate = 0.0;


                double total = 0;
                foreach (DetailRecParam det in newList)
                {
                    culAc += det.CulAcceptCount;
                    keyAc += det.KeyAcceptCount;
                    youngAc += det.YoungAcceptCount;

                    culRate += det.CulRate;
                    keyRate += det.KeyRate;
                    youngRate += det.YoungRate;

                    culCaled += det.CulCaled;
                    keyCaled += det.KeyCaled;
                    youngCaled += det.YoungCaled;
                    
                    culRound += det.CulRound;
                    keyRound += det.KeyRound;
                    youngRound += det.YoungRound;

                    culFundingRate += det.CulFundingRate;
                    keyFundingRate += det.KeyFundingRate;
                    youngFundingRate += det.YoungFundingRate;

                    total += det.Total;
                }

                newList.Add(new DetailRecParam() { 
                    DepartmentName = "合计", CulAcceptCount = culAc, KeyAcceptCount = keyAc, YoungAcceptCount = youngAc,
                    CulRate = culRate, KeyRate = keyRate, YoungRate = youngRate, CulCaled = culCaled, KeyCaled = keyCaled, YoungCaled = youngCaled,
                    CulRound = culRound, KeyRound = keyRound, YoungRound = youngRound, Total = total,
                    CulFundingRate = culFundingRate, KeyFundingRate = keyFundingRate, YoungFundingRate = youngFundingRate
                });
                ultimateRecArray.Add(newList);
            }
            // MessageBox.Show(ultimateRecArray.Count.ToString());
        }


        ArrayList arrLevel1Aux = new ArrayList();
        // 四舍五入后是否
        private void ReduceLevel1()
        {
            int idx = 0, up = ultimateRecArray.Count;
            while (idx < up)
            {
                ArrayList uArr = (ArrayList)ultimateRecArray[idx];
                if ((((DetailRecParam)uArr[uArr.Count - 1]).KeyRound != ((DetailRecParam)uArr[uArr.Count - 1]).KeyCaled) ||
                    (((DetailRecParam)uArr[uArr.Count - 1]).CulRound != ((DetailRecParam)uArr[uArr.Count - 1]).CulCaled))
                {
                    arrLevel1Aux.Add(ultimateRecArray[idx]);
                    ultimateRecArray.RemoveAt(idx);

                    up--;
                }
                else
                {
                    idx++;
                }
            }
        }

        private ArrayList arrLevel2Aux;
        private ArrayList arrLevel3Aux;
        private ArrayList arrLevel4Aux;
        // 
        private ArrayList rst;



        private class SortBySum : System.Collections.IComparer
        {
            public int Compare(Object x, Object y)
            {
                return Convert.ToInt32(((Level2Aux)y).Sum - ((Level2Aux)x).Sum);
            }
        }

        // 第二层筛选函数，返回一个筛选后解集的大小
        private int ReduceLevel2()
        {
            rst = new ArrayList();
            for (int i = 0; i < ultimateRecArray.Count; i++)
            {
                rst.Add(new Level2Aux() { Index = i, Sum = 0.0 });
                
                double sum = 0.0;
                ArrayList oneResult = (ArrayList)ultimateRecArray[i];
                for (int j = 0; j < oneResult.Count; j++)
                {
                    sum += Math.Pow(((DetailRecParam)oneResult[j]).CulCaled - ((DetailRecParam)oneResult[j]).CulRound, 2);
                }
                ((Level2Aux)rst[i]).Sum = sum;
            }
            SortBySum sbs = new SortBySum();
            rst.Sort(sbs);

            //MessageBox.Show(rst.Count.ToString());
            int up = ultimateRecArray.Count - 1;
            int start = Convert.ToInt32(ultimateRecArray.Count * 0.2);

            ArrayList ar = new ArrayList();
            for (int i = up; i > start; --i)
            {
                arrLevel2Aux.Add(ultimateRecArray[((Level2Aux)rst[i]).Index]);
                ar.Add(((Level2Aux)rst[i]).Index);
            }
            ar.Sort();

            for (int i = ar.Count - 1; i >= 0; --i)
            {
                arrLevel2Aux.Add(ultimateRecArray[(int)ar[i]]);
                ultimateRecArray.RemoveAt((int)ar[i]);
            }

            /*
            for (int i = 0; i < arrLevel1Aux.Count; ++i)
            {
                if (ultimateRecArray.Count < 10)
                {
                    ultimateRecArray.Add(arrLevel1Aux[i]);
                }
            }
            */
            return ultimateRecArray.Count;

            //MessageBox.Show(ultimateRecArray.Count.ToString());
        }

        private class SortByRemMoney : System.Collections.IComparer
        {
            public int Compare(Object x, Object y)
            {
                return Convert.ToInt32(((DetailRecParam)((ArrayList)y)[((ArrayList)y).Count - 1]).Total - ((DetailRecParam)((ArrayList)x)[((ArrayList)x).Count - 1]).Total);

            }
        }



        private int ReduceLevel3()
        {
            arrLevel3Aux.Clear();
            SortByRemMoney sbrm = new SortByRemMoney();
            ultimateRecArray.Sort(sbrm);
            int up = ultimateRecArray.Count - 1;
            int start = Convert.ToInt32(ultimateRecArray.Count * 0.2);

            for (int i = up; i >= start; --i)
            {
                arrLevel3Aux.Add(ultimateRecArray[i]);
                ultimateRecArray.RemoveAt(i);
            }

            return ultimateRecArray.Count;
        }


        private class SortByKeyCul : System.Collections.IComparer
        {
            public int Compare(Object x, Object y)
            {
                double rate1 = ((DetailRecParam)((ArrayList)y)[((ArrayList)y).Count - 1]).RecParam.KeyRatio - ((DetailRecParam)((ArrayList)y)[((ArrayList)y).Count - 1]).RecParam.CulRatio;
                double rate2 = ((DetailRecParam)((ArrayList)x)[((ArrayList)x).Count - 1]).RecParam.KeyRatio - ((DetailRecParam)((ArrayList)x)[((ArrayList)x).Count - 1]).RecParam.CulRatio;
                return Convert.ToInt32(rate1 - rate2);
            }
        }

        // 重点最小舍入优先排序规则对象
        // 把方案 x 和方案 y 的每一个重点舍入绝对值相加，总数最小的优先
        private class SortByMinRound : System.Collections.IComparer
        {
            public int Compare(Object x, Object y)
            {
                // round1 为 x 中每个重点舍入绝对值之和，round2 为 y 中每个重点舍入绝对值之和
                double round1 = 0.0, round2 = 0.0;
                ArrayList obx = (ArrayList)x, oby = (ArrayList)y;
                DetailRecParam det1, det2;
                for (int i = 0; i < obx.Count; i++)
                {
                    // 获取 DetailRecParam 对象
                    det1 = (DetailRecParam)obx[i]; det2 = (DetailRecParam)oby[i];
                    round1 += Math.Abs(det1.KeyRound - det1.KeyCaled);
                    round2 += Math.Abs(det2.KeyRound - det2.KeyCaled);
                }

                return (int)((round1 - round2) * 10);
            }
        }



        private int  ReduceLevel4()
        {
            SortByKeyCul sbkc = new SortByKeyCul();
            ultimateRecArray.Sort(sbkc);

            int up = ultimateRecArray.Count - 1;
            int start = Convert.ToInt32(ultimateRecArray.Count * 0.2);
            
            for (int i = up; i >= start; --i)
            {
                arrLevel4Aux.Add(ultimateRecArray[i]);
                ultimateRecArray.RemoveAt(i);
            }

            return ultimateRecArray.Count;
        }


        public void RunCalculator()
        {
            
            recParamArray.ClearArray();

            int keyCount = 0, culCount = 0, youngCount = 0;
            double keyInten = 0.0, culInten = 0.0, youngInten = 0.0;

            double testTotalMoney = 0.0;

            youngCount = adjParameters.YoungCountMin;
            youngInten = adjParameters.YoungIntenMin;

            for (keyCount = adjParameters.KeyCountMin; keyCount <= adjParameters.KeyCountMax; keyCount++)
            {
                for (keyInten = adjParameters.KeyIntenMin; keyInten <= adjParameters.KeyIntenMax; keyInten++)
                {
                    for (culCount = adjParameters.CulCountMin; culCount <= adjParameters.CulCountMax; culCount++)
                    {
                        for (culInten = adjParameters.CulIntenMin; culInten <= adjParameters.CulIntenMax; culInten++)
                        {
                            testTotalMoney = keyCount * keyInten + culCount * culInten + youngCount * youngInten;
                            //MessageBox.Show(culInten.ToString());
                            /*MessageBox.Show(testTotalMoney.ToString() + " = " + keyCount.ToString() + "*" + keyInten.ToString() + "+" + culCount.ToString() +
                                "*" + culInten.ToString() + "+" + adjParameters.YoungCountMin.ToString() + "*" + adjParameters.YoungIntenMin.ToString());*/
                            if ((testTotalMoney >= adjParameters.TotalMoney - adjParameters.RemMoneyMax) && (testTotalMoney <= adjParameters.TotalMoney - adjParameters.RemMoneyMin))
                            {

                                Department dep = initParameters.GetStatistic();
                                recParamArray.AddRecParam(new RecParameter() { 
                                    RemMoney = adjParameters.TotalMoney - testTotalMoney,
                                    KeyCount = keyCount, CulCount = culCount, YoungCount = adjParameters.YoungCountMin,
                                    KeyInten = keyInten, CulInten = culInten, YoungInten = adjParameters.YoungIntenMin,
                                    KeyRatio = keyCount * 1.0 / dep.KeyAcceptCount,
                                    CulRatio = culCount * 1.0 / dep.CulAcceptCount,
                                    YoungRatio = youngCount * 1.0 / dep.YoungAcceptCount
                                });
                                // MessageBox.Show(culInten.ToString());
                            }
                        }
                    }
                }
            }
            
            // MessageBox.Show(recParamArray.GetRecParamsCount().ToString());
            RefreshUltimateArray();
            // MessageBox.Show("Least " + ultimateRecArray.Count.ToString() + " fangan");
            ReduceLevel1();
            // MessageBox.Show(ultimateRecArray.Count.ToString() + " fangans after level1");
            int rd2 = ReduceLevel2();
            // MessageBox.Show(ultimateRecArray.Count.ToString() + " fangans after level2");
            int rd3 = 0;
            if (rd2 > 10) rd3 = ReduceLevel3();
            // MessageBox.Show(ultimateRecArray.Count.ToString() + " fangans after level3");
            //int rd4 = ReduceLevel4();

            // MessageBox.Show(rd2.ToString());
            // MessageBox.Show(rd3.ToString());

            arrLevel1Aux.Sort(new SortByMinRound());
            arrLevel2Aux.Sort(new SortByMinRound());
            arrLevel3Aux.Sort(new SortByMinRound());

            // MessageBox.Show(rd2.ToString() + "    " + rd3.ToString());

            // 此处的两个判断均是防止结果太少而准备的向结果中添加已经删除的解以适当增加解的个数
            if (rd2 == 0)
            {
                for (int i = 0; i < arrLevel2Aux.Count; i++)
                {
                    if (ultimateRecArray.Count > 14) break;
                    ultimateRecArray.Add(arrLevel2Aux[i]);
                }
                for (int i = 0; i < arrLevel1Aux.Count; ++i)
                {
                    if (ultimateRecArray.Count < 10)
                    {
                        ultimateRecArray.Add(arrLevel1Aux[i]);
                    }
                    else
                        break;
                }

            }
            else if (rd3 == 0)
            {
                for (int i = 0; i < arrLevel3Aux.Count; i++)
                {
                    ultimateRecArray.Add(arrLevel3Aux[i]);
                }

                for (int j = 0; j < arrLevel2Aux.Count; ++j)
                {
                    ultimateRecArray.Add(arrLevel2Aux[j]);
                    if (ultimateRecArray.Count > 10) break;
                }
                   
                for (int j = 0; j < arrLevel1Aux.Count; ++j)
                {
                    if (ultimateRecArray.Count < 10)
                    {
                        ultimateRecArray.Add(arrLevel1Aux[j]);
                    }
                    else
                        break;
                }
            }
            else
            {
                for (int i = 0; i < 3 && i < arrLevel3Aux.Count && i < arrLevel2Aux.Count; ++i)
                {
                    ultimateRecArray.Add(arrLevel2Aux[i]);
                    ultimateRecArray.Add(arrLevel3Aux[i]);
                }
            }
            //else if (rd4 == 0)
                //ultimateRecArray = arrLevel4Aux;

            // MessageBox.Show(recParamArray.GetRecParamsCount().ToString());
        }

        // public 


    }
}
