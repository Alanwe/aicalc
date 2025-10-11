using Xunit;
using AiCalc.Models;

namespace AiCalc.Tests;

public class CellAddressTests
{
    [Theory]
    [InlineData("A1", "Sheet1", 0, 0)]
    [InlineData("B2", "Sheet1", 1, 1)]
    [InlineData("Z26", "Sheet1", 25, 25)]
    [InlineData("AA1", "Sheet1", 0, 26)]
    [InlineData("AB10", "Sheet1", 9, 27)]
    [InlineData("Sheet2!C5", "Sheet1", 4, 2)] // Sheet2 should be parsed
    public void TryParse_ValidInput_ReturnsTrue(string input, string defaultSheet, int expectedRow, int expectedCol)
    {
        // Act
        var result = CellAddress.TryParse(input, defaultSheet, out var address);

        // Assert
        Assert.True(result);
        Assert.Equal(expectedRow, address.Row);
        Assert.Equal(expectedCol, address.Column);
    }

    [Theory]
    [InlineData("Sheet2!C5", "Sheet2")] // Sheet name should be parsed correctly
    public void TryParse_WithSheetName_ParsesSheetCorrectly(string input, string expectedSheet)
    {
        // Act
        var result = CellAddress.TryParse(input, "Sheet1", out var address);

        // Assert
        Assert.True(result);
        Assert.Equal(expectedSheet, address.SheetName);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("A0")]
    [InlineData("@1")]
    [InlineData("A-1")]
    public void TryParse_InvalidInput_ReturnsFalse(string input)
    {
        // Act
        var result = CellAddress.TryParse(input, "Sheet1", out _);

        // Assert
        Assert.False(result);
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(1, "B")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(27, "AB")]
    [InlineData(51, "AZ")]
    [InlineData(52, "BA")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    public void ColumnIndexToName_VariousIndices_ReturnsCorrectName(int index, string expected)
    {
        // Act
        var result = CellAddress.ColumnIndexToName(index);

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ToString_FormatsCorrectly()
    {
        // Arrange
        var address = new CellAddress("Sheet1", 0, 0);

        // Act
        var result = address.ToString();

        // Assert
        Assert.Equal("Sheet1!A1", result);
    }

    [Fact]
    public void ToString_WithLargeIndices_FormatsCorrectly()
    {
        // Arrange
        var address = new CellAddress("Sheet2", 99, 26); // Row 100, Column AA

        // Act
        var result = address.ToString();

        // Assert
        Assert.Equal("Sheet2!AA100", result);
    }
}
