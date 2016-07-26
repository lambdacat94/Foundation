using System;

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Foundation
{
    public sealed class RecParamArray
    {

        private ArrayList recParamArray;
        public ArrayList GetRecArrayRef()
        {
            return recParamArray;
        }


        public void SortAll(IComparer comparer)
        {
            recParamArray.Sort(comparer);
        }


        public RecParamArray()
        {
            recParamArray = new ArrayList();

        }

        public int GetRecParamsCount()
        {
            return recParamArray.Count;
        }

        public void AddRecParam(RecParameter newRecParam)
        {
            recParamArray.Add(newRecParam);
        }

        public void ClearArray()
        {
            recParamArray.Clear();
        }

    }
}
