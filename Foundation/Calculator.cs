using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
                    /*
                    MessageBox.Show("TotalCulAccept " + totalCulAccept.ToString() + " TotalKeyAccept " + totalKeyAccept.ToString() +
                        " CulAcc " + culAc.ToString() + " KeyAcc " + keyAc.ToString());
                    */
                    // ===========================================
                    if (totalCulAccept == 0.0)
                        culRate = 0.0;
                    else
                        culRate = culAc * 1.0 / totalCulAccept;
                    // ===========================================
                    if (totalKeyAccept == 0.0)
                        keyRate = 0.0;
                    else
                        keyRate = keyAc * 1.0 / totalKeyAccept;
                    // ===========================================
                    if (totalYoungAccept == 0.0)
                        youngRate = 0.0;
                    else
                        youngRate = youngAc * 1.0 / totalYoungAccept;


                    // MessageBox.Show("CulRate " + culRate.ToString() + " KeyRate " + keyRate.ToString());



                    culCaled = ((RecParameter)recParamArray.GetRecArrayRef()[i]).CulCount * 1.0 * culRate;
                    keyCaled = ((RecParameter)recParamArray.GetRecArrayRef()[i]).KeyCount * 1.0 * keyRate;
                    youngCaled = ((RecParameter)recParamArray.GetRecArrayRef()[i]).YoungCount * 1.0 * youngRate;
                    //======= = ========================================================================================
                    //MessageBox.Show(youngCaled.ToString() + "    " + youngRate.ToString() + "    " + totalYoungAccept.ToString());

                    culFundingRate = ((RecParameter)recParamArray.GetRecArrayRef()[i]).CulCount * 1.0 / totalCulAccept;
                    keyFundingRate = ((RecParameter)recParamArray.GetRecArrayRef()[i]).KeyCount * 1.0 / totalKeyAccept;
                    youngFundingRate = ((RecParameter)recParamArray.GetRecArrayRef()[i]).YoungCount * 1.0 / totalYoungAccept;

                    youngRound = (int)Math.Round(youngCaled);

                    culRound = (int)Math.Round(culCaled);
                    keyRound = (int)Math.Round(keyCaled);

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
                ultimateRecArray.RemoveAt((int)ar[i]);
            }

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
            ReduceLevel1();
            int rd2 = ReduceLevel2();
            int rd3 = ReduceLevel3();
            //int rd4 = ReduceLevel4();

            if (rd2 == 0)
            {
                for (int i = 0; i < arrLevel2Aux.Count; i++)
                {
                    ultimateRecArray.Add(arrLevel2Aux[i]);
                }
            }
            else if (rd3 == 0)
            {
                for (int i = 0; i < arrLevel3Aux.Count; i++)
                {
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
