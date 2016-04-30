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

        /// <summary>
        /// Converts a range to an IEnumerable by row and then by column.
        /// </summary>
        /// <param name="range">Intended for parameters that take in a
        /// range and it is converted to a two dimensional object array by ExcelDNA</param>
        /// <returns></returns>
        internal static IEnumerable<object> SelectByRow(this object[,] range)
        {
            for (int x = 0; x < range.RowCount(); x++)
                for (int y = 0; y < range.ColumnCount(); y++)
                    yield return range[x, y];
        }

        /// <summary>
        /// Converts a range to an IEnumerable by column and then by row.
        /// </summary>
        /// <param name="range">Intended for parameters that take in a
        /// range and it is converted to a two dimensional object array by ExcelDNA</param>
        /// <returns></returns>
        internal static IEnumerable<object> SelectByColumn(this object[,] range)
        {
            for (int y = 0; y < range.ColumnCount(); y++)
                for (int x = 0; x < range.RowCount(); x++)
                    yield return range[x, y];
        }

    }
}
