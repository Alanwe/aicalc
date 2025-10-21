using AiCalc.Services;
// using AiCalc.ViewModels; // not required for these tests
using Xunit;

namespace AiCalc.Tests;

public class FormulaParsingTests
{
    [Fact]
    public void SplitArguments_SimpleArgs()
    {
        var input = "A1, B2, 3";
    var parts = FormulaParser.SplitArguments(input).ToArray();
        Assert.Equal(3, parts.Length);
        Assert.Equal("A1", parts[0].Trim());
        Assert.Equal("B2", parts[1].Trim());
        Assert.Equal("3", parts[2].Trim());
    }

    [Fact]
    public void SplitArguments_QuotedStrings()
    {
        var input = "\"hello, world\", 123";
    var parts = FormulaParser.SplitArguments(input).ToArray();
        Assert.Equal(2, parts.Length);
        Assert.Equal("\"hello, world\"", parts[0].Trim());
        Assert.Equal("123", parts[1].Trim());
    }

    [Fact]
    public void SplitArguments_NestedFunctions()
    {
        var input = "SUM(A1, A2), 5";
    var parts = FormulaParser.SplitArguments(input).ToArray();
        Assert.Equal(2, parts.Length);
        Assert.Equal("SUM(A1, A2)", parts[0].Trim());
        Assert.Equal("5", parts[1].Trim());
    }

    [Fact]
    public void SplitArguments_EscapedQuotes()
    {
        var input = "\"He said \\\"hi\\\"\",A1";
    var parts = FormulaParser.SplitArguments(input).ToArray();
        Assert.Equal(2, parts.Length);
        Assert.Equal("\"He said \\\"hi\\\"\"", parts[0].Trim());
        Assert.Equal("A1", parts[1].Trim());
    }
}
