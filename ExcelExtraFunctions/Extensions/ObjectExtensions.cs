using ExcelDna.Integration;
using System;

namespace ExcelExtraFunctions.Extensions
{
    internal static class ObjectExtensions
    {
        internal static int IfMissing(this object parameter, int defaultValue)
        {
            if (parameter is ExcelMissing)
                return defaultValue;
            else
                return int.Parse(parameter.ToString());
        }

        internal static bool IfMissing(this object parameter, bool defaultValue)
        {
            if (parameter is ExcelMissing)
                return defaultValue;
            else
                return bool.Parse(parameter.ToString());
        }

        internal static string IfMissing(this object parameter, string defaultValue)
        {
            if (parameter is ExcelMissing)
                return defaultValue;
            else
                return parameter.ToString();
        }

        internal static TEnum IfMissingEnum<TEnum>(this object Parameter, TEnum DefaultValue) where TEnum : struct, IComparable, IConvertible, IFormattable
        {
            TEnum ReturnValue;

            if (Enum.TryParse(Parameter.ToString(), out ReturnValue))
                return ReturnValue;
            else
                return DefaultValue;
        }
    }
}
