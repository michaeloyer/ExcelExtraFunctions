using System.Text.RegularExpressions;

namespace ExcelExtraFunctions
{
    public static class RegularExpressions
    {
        public static bool IsMatch(string input, string pattern) => Regex.IsMatch(input, pattern);
        public static string Escape(string input) => Regex.Escape(input);
        public static string Replace(string input, string pattern, string replacement) => Regex.Replace(input, pattern, replacement);
        public static int Count(string input, string pattern) => Regex.Matches(input, pattern).Count;
    }
}
