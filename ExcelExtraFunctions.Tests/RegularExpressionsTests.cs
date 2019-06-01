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
        [DataRow("abc123", "", ExcelErrorValue)]
        public void IsMatch_Tests(string source, string pattern, object result)
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

        [DataTestMethod]
        [DataRow("", "")]
        [DataRow("abc", "abc")]
        [DataRow("(abc)", @"\(abc\)")]
        public void Escape_Tests(string pattern, string result)
        {
            Escape(pattern).Should().Be(result);
        }

        [DataTestMethod]
        [DataRow("a1b2c3", "[a-z]", "", "123")]
        [DataRow("a1b2c3", @"\d", "x", "axbxcx")]
        [DataRow("abc123", "[x-z]{3}[4-9]{3}", "", "abc123")]
        [DataRow("abc123", "", "abc", ExcelErrorValue)]
        public void Replace_Tests(string source, string pattern, string replacement, object result)
        {
            Replace(source, pattern, replacement).Should().Be(result);
        }

        [DataTestMethod]
        [DataRow("a1b2c3", "[a-c][1-3]", 3)]
        [DataRow("abc123", "[a-c][1-3]", 1)]
        [DataRow("a1b2c3", "[b-c][1-3]", 2)]
        [DataRow("abc123", "[x-z]{3}[4-9]{3}", 0)]
        [DataRow("abc123", "", ExcelErrorValue)]
        public void Count_Tests(string source, string pattern, object result)
        {
            Count(source, pattern).Should().Be(result);
        }

        [TestMethod]
        public void Split_OnSuccessReturnArray()
        {
            Split("a1b2c3", @"\d")
                .Should().BeEquivalentTo(new[] { "a", "b", "c", "" });
        }

        [TestMethod]
        public void Split_OnNoMatchReturnsArrayOfFullString()
        {
            Split("a1b2c3", "z")
                .Should().BeEquivalentTo(new[] { "a1b2c3" });
        }

        [TestMethod]
        public void Split_EmptyPatternReturnsError()
        {
            Split("abc123", "").Should().Be(ExcelErrorValue);
        }
    }
}
