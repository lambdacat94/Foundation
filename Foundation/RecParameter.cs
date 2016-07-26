using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    public sealed class RecParameter
    {
        // 推荐重点类支持数
        private int keyCount;
        public int KeyCount
        {
            get { return keyCount; }
            set { keyCount = value; }
        }
        // 推荐培育类支持数
        private int culCount;
        public int CulCount
        {
            get { return culCount; }
            set { culCount = value; }
        }
        // 推荐人才支持数
        private int youngCount;
        public int YoungCount
        { 
            get { return youngCount; }
            set { youngCount = value; }
        }


        // 推荐重点类强度
        private double keyInten;
        public double KeyInten
        {
            get { return keyInten; }
            set { keyInten = value; }
        }
        // 推荐培育类强度
        private double culInten;
        public double CulInten
        {
            get { return culInten; }
            set { culInten = value; }
        }
        // 推荐人才类强度
        private double youngInten;
        public double YoungInten
        {
            get { return youngInten; }
            set { youngInten = value; }
        }


        // 推荐重点类资助率
        private double keyRatio;
        public double KeyRatio
        {
            get { return keyRatio; }
            set { keyRatio = value; }
        }
        // 推荐培育类资助率
        private double culRatio;
        public double CulRatio
        {
            get { return culRatio; }
            set { culRatio = value; }
        }

        // ==============================================
        // 推荐本地优青资助率
        private double youngRatio;
        public double YoungRatio
        {
            get { return youngRatio; }
            set { youngRatio = value; }
        }


        // ==============================================



        // 推荐剩余金额
        private double remMoney;
        public double RemMoney
        {
            get { return remMoney; }
            set { remMoney = value; }
        }

        // 设置各种参数，供构造函数或者单独使用，public，但不抛出异常
        public void SetParameter()
        {

        }

        // 无参构造函数
        public RecParameter()
        {

        }

    }
}
