
// This meant to be my lybrary of coommon static functions
// It will eventually go into separate shared/common class library project

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WashU.BatemanLab.Common
{
    public static class ConvertUtil
    {
        public static double doubleTryParse(string s)
        {
            double result;
            double.TryParse(s, out result);
            return result;
        }

        public static double DictionaryTryGetValue()
        {
            return 0;
        }
    }
}
