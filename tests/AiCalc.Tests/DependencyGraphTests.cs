using Xunit;
using AiCalc.Models;
using AiCalc.Services;

namespace AiCalc.Tests;

public class DependencyGraphTests
{
    [Fact]
    public void UpdateCellDependencies_SimpleFormula_ExtractsDependencies()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);
        var cellB1 = new CellAddress("Sheet1", 0, 1);

        // Act
        graph.UpdateCellDependencies(cellA1, "=B1+5");

        // Assert
        var deps = graph.GetDirectDependencies(cellA1);
        Assert.Single(deps);
        Assert.Contains(cellB1, deps);
    }

    [Fact]
    public void UpdateCellDependencies_MultipleReferences_ExtractsAll()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);
        var cellB1 = new CellAddress("Sheet1", 0, 1);
        var cellC1 = new CellAddress("Sheet1", 0, 2);

        // Act
        graph.UpdateCellDependencies(cellA1, "=B1+C1");

        // Assert
        var deps = graph.GetDirectDependencies(cellA1).ToList();
        Assert.Equal(2, deps.Count);
        Assert.Contains(cellB1, deps);
        Assert.Contains(cellC1, deps);
    }

    [Fact]
    public void UpdateCellDependencies_RangeReference_ExtractsRange()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);

        // Act
        graph.UpdateCellDependencies(cellA1, "=SUM(B1:B3)");

        // Assert
        var deps = graph.GetDirectDependencies(cellA1).ToList();
        Assert.Equal(3, deps.Count);
        Assert.Contains(new CellAddress("Sheet1", 0, 1), deps); // B1
        Assert.Contains(new CellAddress("Sheet1", 1, 1), deps); // B2
        Assert.Contains(new CellAddress("Sheet1", 2, 1), deps); // B3
    }

    [Fact]
    public void DetectCircularReference_DirectCircle_ReturnsTrue()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);
        var cellB1 = new CellAddress("Sheet1", 0, 1);

        // Act
        graph.UpdateCellDependencies(cellA1, "=B1");
        graph.UpdateCellDependencies(cellB1, "=A1");

        // Assert
        Assert.NotNull(graph.DetectCircularReference(cellA1));
        Assert.NotNull(graph.DetectCircularReference(cellB1));
    }

    [Fact]
    public void DetectCircularReference_IndirectCircle_ReturnsTrue()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);
        var cellB1 = new CellAddress("Sheet1", 0, 1);
        var cellC1 = new CellAddress("Sheet1", 0, 2);

        // Act
        graph.UpdateCellDependencies(cellA1, "=B1");
        graph.UpdateCellDependencies(cellB1, "=C1");
        graph.UpdateCellDependencies(cellC1, "=A1");

        // Assert
        Assert.NotNull(graph.DetectCircularReference(cellA1));
        Assert.NotNull(graph.DetectCircularReference(cellB1));
        Assert.NotNull(graph.DetectCircularReference(cellC1));
    }

    [Fact]
    public void DetectCircularReference_NoCircle_ReturnsFalse()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);
        var cellB1 = new CellAddress("Sheet1", 0, 1);
        var cellC1 = new CellAddress("Sheet1", 0, 2);

        // Act
        graph.UpdateCellDependencies(cellA1, "=B1");
        graph.UpdateCellDependencies(cellB1, "=C1");

        // Assert
        Assert.Null(graph.DetectCircularReference(cellA1));
        Assert.Null(graph.DetectCircularReference(cellB1));
        Assert.Null(graph.DetectCircularReference(cellC1));
    }

    [Fact]
    public void GetEvaluationOrder_SimpleChain_ReturnsCorrectOrder()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);
        var cellB1 = new CellAddress("Sheet1", 0, 1);
        var cellC1 = new CellAddress("Sheet1", 0, 2);

        // Act
        graph.UpdateCellDependencies(cellA1, "=B1");
        graph.UpdateCellDependencies(cellB1, "=C1");
        var batches = graph.GetEvaluationOrder();

        // Assert
        Assert.NotEmpty(batches);
        
        // Find which batch each cell is in
        int batchC1 = -1, batchB1 = -1, batchA1 = -1;
        for (int i = 0; i < batches.Count; i++)
        {
            if (batches[i].Contains(cellC1)) batchC1 = i;
            if (batches[i].Contains(cellB1)) batchB1 = i;
            if (batches[i].Contains(cellA1)) batchA1 = i;
        }

        // C1 should be in an earlier batch than B1, and B1 earlier than A1
        Assert.True(batchC1 < batchB1);
        Assert.True(batchB1 < batchA1);
    }

    [Fact]
    public void UpdateCellDependencies_ClearOldDependencies_RemovesOldLinks()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);
        var cellB1 = new CellAddress("Sheet1", 0, 1);
        var cellC1 = new CellAddress("Sheet1", 0, 2);

        // Act
        graph.UpdateCellDependencies(cellA1, "=B1");
        graph.UpdateCellDependencies(cellA1, "=C1");

        // Assert
        var deps = graph.GetDirectDependencies(cellA1).ToList();
        Assert.Single(deps);
        Assert.Contains(cellC1, deps);
        Assert.DoesNotContain(cellB1, deps);
    }

    [Fact]
    public void UpdateCellDependencies_NoFormula_ClearsDependencies()
    {
        // Arrange
        var graph = new DependencyGraph();
        var cellA1 = new CellAddress("Sheet1", 0, 0);

        // Act
        graph.UpdateCellDependencies(cellA1, "=B1");
        graph.UpdateCellDependencies(cellA1, "100");

        // Assert
        var deps = graph.GetDirectDependencies(cellA1).ToList();
        Assert.Empty(deps);
    }
}
