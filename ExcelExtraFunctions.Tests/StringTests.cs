using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static ExcelDna.Integration.ExcelError;
using static ExcelExtraFunctions.Functions.Strings;

namespace ExcelExtraFunctions.Tests
{
    [TestClass]
    public class StringTests
    {
        [DataTestMethod]
        [DataRow("1abc2abc3abc4abc5abc", "1", "2", 1, 1, "abc")]
        [DataRow("1abc2abc3abc4abc5abc", "a", "a", 1, 2, "bc2abc3")]
        [DataRow("1abc2abc3abc4abc5abc", "a", "a", 2, 2, "bc3abc4")]
        [DataRow("1abc2abc3abc4abc5abc", "a", "a", -2, 1, "bc5")]
        [DataRow("1abc2abc3abc4abc5abc", "a", "a", 1, -1, "bc2abc3abc4abc5")]
        [DataRow("1abc2abc3abc4abc5abc", "a", "c", 1, -3, "bc2abc3ab")]
        [DataRow("abc", "d", "d", 500, 500, ExcelErrorValue)]
        [DataRow("abc", "d", "d", 1, 1, ExcelErrorValue)]
        [DataRow("abc", "a", "d", 1, 1, ExcelErrorValue)]
        [DataRow("abc", "d", "c", 1, 1, ExcelErrorValue)]
        public void Between_Tests(string Source, string LeftDelimiter, string RightDelimiter, double LeftCount, double RightCount, object result) =>
            Between(Source, LeftDelimiter, RightDelimiter, LeftCount, RightCount)
                .Should().Be(result);
    }
}
