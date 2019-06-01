using ExcelDna.Integration;
using ExcelExtraFunctions.Extensions;
using System;
using static ExcelDna.Integration.ExcelError;

namespace ExcelExtraFunctions.Functions
{
    public static class Strings
    {
        [ExcelFunction(Category = "EXF Strings", Name = "STR.BETWEEN",
            Description = "Returns the string between two other strings")]
        public static object Between(string Source, string LeftDelimiter, string RightDelimiter, object LeftCount, object RightCount)
        {
            var leftIndex = Source.IndexOfN(LeftDelimiter, LeftCount.IfMissing(1), StringComparison.CurrentCultureIgnoreCase);
            if (leftIndex == -1)
                return ExcelErrorValue;

            Source = Source.Substring(leftIndex + 1);

            var rightIndex = Source.IndexOfN(RightDelimiter, RightCount.IfMissing(1), StringComparison.CurrentCultureIgnoreCase);
            if (rightIndex == -1)
                return ExcelErrorValue;

            Source = Source.Remove(rightIndex);

            return Source;
        }
    }
}
