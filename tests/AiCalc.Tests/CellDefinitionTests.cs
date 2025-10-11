using Xunit;
using AiCalc.Models;

namespace AiCalc.Tests;

public class CellDefinitionTests
{
    [Fact]
    public void Constructor_InitializesWithDefaults()
    {
        // Arrange & Act
        var cell = new CellDefinition();

        // Assert
        Assert.NotNull(cell.Value);
        Assert.Equal(CellObjectType.Empty, cell.Value.ObjectType);
        Assert.Null(cell.Formula);
        Assert.Equal(CellAutomationMode.Manual, cell.AutomationMode);
    }

    [Fact]
    public void Value_CanBeSet()
    {
        // Arrange
        var cell = new CellDefinition();

        // Act
        cell.Value = new CellValue(CellObjectType.Number, "42", "42");

        // Assert
        Assert.Equal(CellObjectType.Number, cell.Value.ObjectType);
        Assert.Equal("42", cell.Value.SerializedValue);
    }

    [Fact]
    public void AutomationMode_CanBeChanged()
    {
        // Arrange
        var cell = new CellDefinition
        {
            AutomationMode = CellAutomationMode.Manual
        };

        // Act
        cell.AutomationMode = CellAutomationMode.AutoOnOpen;

        // Assert
        Assert.Equal(CellAutomationMode.AutoOnOpen, cell.AutomationMode);
    }

    [Fact]
    public void Formula_CanBeSetAndRetrieved()
    {
        // Arrange
        var cell = new CellDefinition();
        var formula = "=SUM(A1:A10)";

        // Act
        cell.Formula = formula;

        // Assert
        Assert.Equal(formula, cell.Formula);
    }

    [Fact]
    public void Notes_CanContainMarkdown()
    {
        // Arrange
        var cell = new CellDefinition();
        var notes = "# Important Cell\nThis cell contains **critical** data.";

        // Act
        cell.Notes = notes;

        // Assert
        Assert.Equal(notes, cell.Notes);
    }
}
