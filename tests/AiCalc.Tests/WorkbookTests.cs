using Xunit;
using AiCalc.Models;

namespace AiCalc.Tests;

public class WorkbookDefinitionTests
{
    [Fact]
    public void Constructor_InitializesEmptySheets()
    {
        // Arrange & Act
        var workbook = new WorkbookDefinition();

        // Assert
        Assert.NotNull(workbook.Sheets);
        Assert.Empty(workbook.Sheets);
    }

    [Fact]
    public void AddSheet_IncreasesSheetCount()
    {
        // Arrange
        var workbook = new WorkbookDefinition();
        var sheet = new SheetDefinition { Name = "Sheet1" };

        // Act
        workbook.Sheets.Add(sheet);

        // Assert
        Assert.Single(workbook.Sheets);
        Assert.Equal("Sheet1", workbook.Sheets[0].Name);
    }

    [Fact]
    public void Settings_IsInitialized()
    {
        // Arrange & Act
        var workbook = new WorkbookDefinition();

        // Assert
        Assert.NotNull(workbook.Settings);
    }

    [Fact]
    public void MultipleSheets_CanBeAdded()
    {
        // Arrange
        var workbook = new WorkbookDefinition();

        // Act
        workbook.Sheets.Add(new SheetDefinition { Name = "Sheet1" });
        workbook.Sheets.Add(new SheetDefinition { Name = "Sheet2" });
        workbook.Sheets.Add(new SheetDefinition { Name = "Sheet3" });

        // Assert
        Assert.Equal(3, workbook.Sheets.Count);
    }
}

public class SheetDefinitionTests
{
    [Fact]
    public void Constructor_InitializesEmptyCells()
    {
        // Arrange & Act
        var sheet = new SheetDefinition();

        // Assert
        Assert.NotNull(sheet.Cells);
        Assert.Empty(sheet.Cells);
    }

    [Fact]
    public void Name_CanBeSet()
    {
        // Arrange
        var sheet = new SheetDefinition();

        // Act
        sheet.Name = "TestSheet";

        // Assert
        Assert.Equal("TestSheet", sheet.Name);
    }

    [Fact]
    public void Cells_CanBeAdded()
    {
        // Arrange
        var sheet = new SheetDefinition();
        var cell = new CellDefinition
        {
            Address = "A1",
            Value = new CellValue(CellObjectType.Text, "Hello", "Hello")
        };

        // Act
        sheet.Cells.Add(cell);

        // Assert
        Assert.Single(sheet.Cells);
        Assert.Equal("A1", sheet.Cells[0].Address);
    }

    [Fact]
    public void MultipleCells_CanBeAddedToSameSheet()
    {
        // Arrange
        var sheet = new SheetDefinition { Name = "Sheet1" };

        // Act
        for (int i = 0; i < 100; i++)
        {
            sheet.Cells.Add(new CellDefinition
            {
                Address = $"A{i + 1}",
                Value = new CellValue(CellObjectType.Text, $"Value{i}", $"Value{i}")
            });
        }

        // Assert
        Assert.Equal(100, sheet.Cells.Count);
    }
}
