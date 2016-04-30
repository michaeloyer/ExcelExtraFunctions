using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelExtraFunctions
{
    internal static class OptionalParameterObjectExtensions
    {
        internal static T IfMissing<T>(object parameter, T valueIfMissing) where T : struct
        {
            if (parameter is ExcelMissing)
                return valueIfMissing;
            else
                return (T)parameter;
        }

        internal static string IfMissing<T>(object parameter, string valueIfMissing)
        {
            if (parameter is ExcelMissing)
                return valueIfMissing;
            else
                return (string)parameter;
        }
    }
}
