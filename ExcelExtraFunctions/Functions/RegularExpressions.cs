using ExcelDna.Integration;
using System.Linq;
using System.Text.RegularExpressions;
using static ExcelDna.Integration.ExcelError;

namespace ExcelExtraFunctions.Functions
{
    public static class RegularExpressions
    {
        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.ISMATCH",
            Description = "Returns TRUE if a single pattern match is found in the input, otherwise FALSE")]
        public static object IsMatch(string input, string pattern) =>
            string.IsNullOrEmpty(pattern)
                ? (object)ExcelErrorValue
                : Regex.IsMatch(input, pattern);

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.ESCAPE",
            Description = "Puts a '\' (backslash) character in front of all regex modifier characters")]
        public static string Escape(string pattern) => Regex.Escape(pattern);

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.REPLACE",
            Description = "Replaces all pattern matches in the input with the replacement string.")]
        public static object Replace(string input, string pattern,
            [ExcelArgument("If a capture group is used it can be reference with $1, or if is explicitly referenced you can use $Name")]
            string replacement
        ) => string.IsNullOrEmpty(pattern)
                ? (object)ExcelErrorValue
                : Regex.Replace(input, pattern, replacement);

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.COUNT",
            Description = "Counts the number of pattern matches in the input.")]
        public static object Count(string input, string pattern) =>
            string.IsNullOrEmpty(pattern)
                ? (object)ExcelErrorValue
                : Regex.Matches(input, pattern).Count;

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.SPLIT",
            Description = "Splits the input string at each matched pattern and returns the array.")]
        public static object Split(string input, string pattern) =>
            string.IsNullOrEmpty(pattern)
                ? (object)ExcelErrorValue
                : Regex.Split(input, pattern);

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.MATCH",
            Description = "Returns the first matched pattern in the input string.")]
        public static object Match(string input, string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
                return ExcelErrorValue;

            Match match = Regex.Match(input, pattern);
            return match.Success
                ? match.Value
                : (object)ExcelErrorNA;
        }

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.MATCHES",
            Description = "Returns array of matched patterns in the input string.")]
        public static object Matches(string input, string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
                return ExcelErrorValue;

            MatchCollection matches = Regex.Matches(input, pattern);
            if (matches.Count == 0)
                return ExcelErrorNA;

            return matches.Cast<Match>()
                    .Select(m => m.Value)
                    .ToArray();
        }

        [ExcelFunction(Category = "EXF Regular Expression", Name = "RE.GROUPS",
            Description = "Returns 2D array of the groups of each match. Can return full match")]
        public static object Groups(string input, string pattern,
            [ExcelArgument("Includes the full match along with the group of matches at the front of the array")]
                bool includeFullMatch = false)
        {
            if (string.IsNullOrEmpty(pattern))
                return ExcelErrorValue;

            MatchCollection matches = Regex.Matches(input, pattern);
            if (matches.Count == 0 || matches[0].Groups.Count == 0)
                return ExcelErrorNA;

            int includeInt = includeFullMatch ? 0 : 1;
            var array = new string[matches.Count,
                matches[0].Groups.Count - includeInt];

            foreach (var (match, x) in matches.Cast<Match>().Select((m, i) => (m, i)))
                foreach (var (group, y) in match.Groups.Cast<Group>().Skip(includeInt).Select((g, i) => (g: g, i: i)))
                    array[x, y] = group.Value;

            return array;
        }
    }
}
