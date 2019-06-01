using ExcelDna.Integration;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static ExcelDna.Integration.ExcelError;
using static ExcelExtraFunctions.Functions.RegularExpressions;

namespace ExcelExtraFunctions.Tests
{
    [TestClass]
    public class RegularExpressionsTests
    {
        [DataTestMethod]
        [DataRow("abc123", "[a-c]{3}[1-3]{3}", true)]
        [DataRow("abc123", "[x-z]{3}[4-9]{3}", false)]
        public void IsMatch_Tests(string source, string pattern, bool result)
        {
            IsMatch(source, pattern).Should().Be(result);
        }

        [DataTestMethod]
        [DataRow("abc123", "[a-c][1-3]", "c1")]
        [DataRow("a1b2c3", "[b-c][1-3]", "b2")]
        [DataRow("abc123", "[x-z]{3}[4-9]{3}", ExcelErrorNA)]
        [DataRow("abc123", "", ExcelErrorValue)]
        public void Match_Tests(string source, string pattern, object result)
        {
            Match(source, pattern).Should().Be(result);
        }

        [TestMethod]
        public void Matches_OnSuccessReturnArray()
        {
            Matches("a1b2c3", "[a-z][0-9]")
                .Should().BeEquivalentTo(new[] { "a1", "b2", "c3" });
        }

        [DataTestMethod]
        [DataRow("abc123", "[x-z]{3}[4-9]{3}", ExcelErrorNA)]
        [DataRow("abc123", "", ExcelErrorValue)]
        public void Matches_Errors(string source, string pattern, ExcelError result)
        {
            Matches(source, pattern).Should().Be(result);
        }

        [TestMethod]
        public void Groups_OnSuccessReturnArray()
        {
            Groups("a1b2c3", "[a-z]([0-9])")
                .Should().BeEquivalentTo(new[,] {
                        { "1" },
                        { "2" },
                        { "3" },
                    }
                );
        }

        [TestMethod]
        public void Groups_OnSuccessReturnArrayWithFullMatch()
        {
            Groups("a1b2c3", "[a-z]([0-9])", includeFullMatch: true)
                .Should().BeEquivalentTo(new[,] {
                        { "a1", "1" },
                        { "b2", "2" },
                        { "c3", "3" },
                    }
                );
        }

        [DataTestMethod]
        [DataRow("a1b2c3", "[x-z]{3}[0-9]{3}", ExcelErrorNA)]
        [DataRow("a1b2c3", "[x-z]{3}([4-9]{3})", ExcelErrorNA)]
        [DataRow("a1b2c3", "", ExcelErrorValue)]
        public void Groups_Errors(string source, string pattern, ExcelError result)
        {
            Groups(source, pattern).Should().Be(result);
        }
    }
}
