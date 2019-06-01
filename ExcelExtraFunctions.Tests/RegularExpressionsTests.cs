using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static ExcelExtraFunctions.Functions.RegularExpressions;

namespace ExcelExtraFunctions.Tests
{
    [TestClass]
    public class RegularExpressionsTests
    {
        [DataTestMethod]
        [DataRow("abc123", "[a-c]{3}[1-3]{3}", true)]
        [DataRow("abc123", "[1-3]{3}[a-c]{3}", false)]
        public void IsMatch_Tests(string source, string pattern, bool result)
        {
            IsMatch(source, pattern).Should().Equals(result);
        }
    }
}
