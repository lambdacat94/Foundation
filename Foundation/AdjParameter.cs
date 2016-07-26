using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    public sealed class AdjParameter
    {
        private double totalMoney;
        public double TotalMoney
        {
            get { return totalMoney; }
            set { if (value is double) totalMoney = value; }
        }


        // 剩余金额最大最小值
        private double remMoneyMin;
        private double remMoneyMax;
        public double RemMoneyMin
        {
            get { return remMoneyMin; }
            set { if (value is double) remMoneyMin = value; }
        }
        public double RemMoneyMax
        {
            get { return remMoneyMax; }
            set { if (value is double) remMoneyMax = value; }
        }



        // 重点类支持数最大最小值
        private int keyCountMin;
        private int keyCountMax;
        public int KeyCountMin
        { 
            get { return keyCountMin; }
            set { if (value is int) keyCountMin = value; }
        }
        public int KeyCountMax
        {
            get { return keyCountMax; }
            set { if (value is int) keyCountMax = value; }
        }
        
        // 培育类支持数最大最小值
        private int culCountMin;
        private int culCountMax;
        public int CulCountMin
        {
            get { return culCountMin; }
            set { if (value is int) culCountMin = value; }
        }
        public int CulCountMax
        {
            get { return culCountMax; }
            set { if (value is int) culCountMax = value; }
        }



        // 青年人才类最大最小值
        private int youngCountMin;
        // private int yougCountMax;
        public int YoungCountMin
        {
            get { return youngCountMin; }
            set { if (value is int) youngCountMin = value; }
        }



        // 重点类强度最大最小值
        private double keyIntenMin;
        private double keyIntenMax;
        public double KeyIntenMin
        {
            get { return keyIntenMin; }
            set { if (value is double) keyIntenMin = value; }
        }
        public double KeyIntenMax
        {
            get { return keyIntenMax; }
            set { if (value is double) keyIntenMax = value; }
        }



        // 培育类强度最大最小值
        private double culIntenMin;
        private double culIntenMax;
        public double CulIntenMin
        {
            get { return culIntenMin; }
            set { if (value is double) culIntenMin = value; }
        }
        public double CulIntenMax
        {
            get { return culIntenMax; }
            set { if (value is double) culIntenMax = value; }
        }


        // 青年人才强度最大最小值
        private double youngIntenMin;
        // private double youngIntenMax;
        public double YoungIntenMin
        {
            get { return youngIntenMin; }
            set { if (value is double) youngIntenMin = value; }
        }

        
        public AdjParameter()
        {

        }
        



    }
}
