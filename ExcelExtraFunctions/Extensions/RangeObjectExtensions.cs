using System.Collections.Generic;

namespace ExcelExtraFunctions.Extensions
{
    internal static class RangeObjectExtensions
    {
        /// <summary>
        /// Converts a range to an IEnumerable by row and then by column.
        /// </summary>
        /// <param name="range">Intended for parameters that take in a
        /// range and it is converted to a two dimensional object array by ExcelDNA</param>
        /// <returns></returns>
        internal static IEnumerable<object> SelectByRow(this object[,] range)
        {
            for (int x = 0; x < range.GetLength(0); x++)
                for (int y = 0; y < range.GetLength(1); y++)
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
            for (int y = 0; y < range.GetLength(1); y++)
                for (int x = 0; x < range.GetLength(0); x++)
                    yield return range[x, y];
        }

    }
}
