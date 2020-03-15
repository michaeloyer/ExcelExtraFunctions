using ExcelDna.Integration;
using System;

namespace ExcelExtraFunctions.Extensions
{
    internal static class ObjectExtensions
    {
        internal static int IfMissing(this object parameter, int defaultValue) =>
            parameter is ExcelMissing
                ? defaultValue
            : parameter is int i
                ? i
            :
                int.Parse(parameter.ToString());

        internal static bool IfMissing(this object parameter, bool defaultValue) =>
            parameter is ExcelMissing
                ? defaultValue
            : parameter is bool b
                ? b
            :
                bool.Parse(parameter.ToString());

        internal static string IfMissing(this object parameter, string defaultValue) =>
            parameter is ExcelMissing
                ? defaultValue
            : parameter is string s
                ? s
            :
                parameter.ToString();

        internal static TEnum IfMissingEnum<TEnum>(this object Parameter, TEnum DefaultValue) where TEnum : struct, IComparable, IConvertible, IFormattable
        {
            if (Enum.TryParse(Parameter.ToString(), out TEnum ReturnValue))
                return ReturnValue;
            else
                return DefaultValue;
        }
    }
}
