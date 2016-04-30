using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace ExcelExtraFunctions
{
    public static class RegularExpressions
    {
        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.ISMATCH",
            Description = "Returns TRUE if a single pattern match is found in the input, otherwise FALSE")]
        public static bool IsMatch(string input, string pattern) => Regex.IsMatch(input, pattern);

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.ESCAPE",
            Description = "Puts a '\' (backslash) character in front of all regex modifier characters")]
        public static string Escape(string pattern) => Regex.Escape(pattern);

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.REPLACE",
            Description = "Replaces all pattern matches in the input with the replacement string.")]
        public static string Replace(string input, string pattern, 
            [ExcelArgument("If a capture group is used it can be reference with $1, or if is explicitly referenced you can use $Name")] string replacement
            ) => Regex.Replace(input, pattern, replacement);

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.COUNT", 
            Description = "Counts the number of pattern matches in the input.")]
        public static int Count(string input, string pattern) => Regex.Matches(input, pattern).Count;

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.SPLIT",
            Description = "Splits the input string at each matched pattern and returns the array.")]
        public static object Split(string input, string pattern) => Regex.Split(input, pattern);

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.MATCH",
            Description = "Returns the first matched pattern in the input string.")]
        public static object Match(string input, string pattern) => Regex.Match(input, pattern);
    }
}
