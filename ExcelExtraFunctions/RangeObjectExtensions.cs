using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtraFunctions
{
    internal static class RangeObjectExtensions
    {
        internal static int RowCount(this object[,] range)
        {
            return range.GetLength(0);
        }

        internal static int ColumnCount(this object[,] range)
        {
            return range.GetLength(1);
        }
    }
}
