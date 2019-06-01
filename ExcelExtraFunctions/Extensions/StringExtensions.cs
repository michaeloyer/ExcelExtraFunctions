using System;

namespace ExcelExtraFunctions.Extensions
{
    public static class StringExtensions
    {
        /// <summary>
        /// String Extension Method that reports the zero-based index of the nth occurrence of the specified string
        /// </summary>
        /// <param name="searchString">The string to search through</param>
        /// <param name="value">The string to search for in searchString</param>
        /// <param name="occurance">The nth occurrence of value in searchString. 
        /// A negative number will find the nth occurrence going backwards from the end of the string. </param>
        /// <param name="comparisonType"></param>
        /// <returns>Zero-based index of value. -1 if the nth occurrence doesn't exist.</returns>
        public static int IndexOfN(this string searchString, string value, int occurance, StringComparison comparisonType = StringComparison.CurrentCulture)
        {
            if (occurance > 0)
            {
                var index = -1;
                do
                {
                    index = searchString.IndexOf(value, index + 1, comparisonType);
                    occurance--;
                } while (occurance > 0 && index != -1);

                return index;
            }
            else if (occurance < 0)
            {
                int index = searchString.Length;
                do
                {
                    occurance++;
                    index = searchString.LastIndexOf(value, index - 1, comparisonType);
                } while (occurance < 0 && index != -1);

                return index;
            }
            else
                return -1;
        }
    }
}
