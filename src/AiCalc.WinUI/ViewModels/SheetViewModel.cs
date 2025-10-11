using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using AiCalc.Models;

namespace AiCalc.ViewModels;

public class SheetViewModel
{
    private readonly WorkbookViewModel _workbook;

    public SheetViewModel(WorkbookViewModel workbook, string name, int rows, int columns)
    {
        _workbook = workbook;
        Workbook = workbook;
        Name = name;
        Rows = new ObservableCollection<RowViewModel>(Enumerable.Range(0, rows).Select(r => CreateRow(r, columns)));
    }

    public WorkbookViewModel Workbook { get; }

    public string Name { get; }

    public ObservableCollection<RowViewModel> Rows { get; }

    public int ColumnCount => Rows.FirstOrDefault()?.Cells.Count ?? 0;

    public IReadOnlyList<string> ColumnHeaders => Enumerable.Range(0, ColumnCount).Select(CellAddress.ColumnIndexToName).ToList();

    public CellViewModel? GetCell(int row, int column)
    {
        if (row < 0 || row >= Rows.Count)
        {
            return null;
        }

        var rowVm = Rows[row];
        if (column < 0 || column >= rowVm.Cells.Count)
        {
            return null;
        }

        return rowVm.Cells[column];
    }

    public IEnumerable<CellViewModel> Cells => Rows.SelectMany(r => r.Cells);

    /// <summary>
    /// Evaluate all cells in this sheet using multi-threaded evaluation engine
    /// Skips cells with Manual automation mode per Task 10 requirement
    /// </summary>
    public async Task EvaluateAllAsync()
    {
        var cellsToEvaluate = Cells
            .Where(c => c.HasFormula && c.AutomationMode != CellAutomationMode.Manual)
            .ToDictionary(c => c.Address, c => c);

        if (cellsToEvaluate.Count == 0)
        {
            return;
        }

        await _workbook.EvaluationEngine.EvaluateAllAsync(cellsToEvaluate);
    }

    /// <summary>
    /// Update dependency graph when a cell's formula changes
    /// </summary>
    public void UpdateCellDependencies(CellViewModel cell)
    {
        _workbook.DependencyGraph.UpdateCellDependencies(cell.Address, cell.Formula);
    }

    private RowViewModel CreateRow(int rowIndex, int columnCount)
    {
        var row = new RowViewModel(rowIndex);
        for (var col = 0; col < columnCount; col++)
        {
            row.Cells.Add(new CellViewModel(_workbook, this, rowIndex, col));
        }

        return row;
    }
    
    // ==================== Row/Column Operations (Phase 5 Task 16) ====================
    
    public void InsertRow(int index)
    {
        if (index < 0 || index > Rows.Count) return;
        
        var newRow = CreateRow(index, ColumnCount);
        Rows.Insert(index, newRow);
        
        // Update row indices for all rows after the inserted row
        for (int r = index + 1; r < Rows.Count; r++)
        {
            Rows[r] = RecreateRowWithNewIndex(Rows[r], r);
        }
    }
    
    public void DeleteRow(int index)
    {
        if (index < 0 || index >= Rows.Count || Rows.Count <= 1) return;
        
        Rows.RemoveAt(index);
        
        // Update row indices for all rows after the deleted row
        for (int r = index; r < Rows.Count; r++)
        {
            Rows[r] = RecreateRowWithNewIndex(Rows[r], r);
        }
    }
    
    public void InsertColumn(int index)
    {
        if (index < 0 || index > ColumnCount) return;
        
        foreach (var row in Rows)
        {
            var newCell = new CellViewModel(_workbook, this, row.Index, index);
            row.Cells.Insert(index, newCell);
            
            // Update column indices for all cells after the inserted column
            for (int c = index + 1; c < row.Cells.Count; c++)
            {
                row.Cells[c] = RecreateCellWithNewIndex(row.Cells[c], row.Index, c);
            }
        }
    }
    
    public void DeleteColumn(int index)
    {
        if (index < 0 || index >= ColumnCount || ColumnCount <= 1) return;
        
        foreach (var row in Rows)
        {
            row.Cells.RemoveAt(index);
            
            // Update column indices for all cells after the deleted column
            for (int c = index; c < row.Cells.Count; c++)
            {
                row.Cells[c] = RecreateCellWithNewIndex(row.Cells[c], row.Index, c);
            }
        }
    }
    
    private RowViewModel RecreateRowWithNewIndex(RowViewModel oldRow, int newIndex)
    {
        var newRow = new RowViewModel(newIndex);
        for (int c = 0; c < oldRow.Cells.Count; c++)
        {
            var oldCell = oldRow.Cells[c];
            var newCell = new CellViewModel(_workbook, this, newIndex, c)
            {
                RawValue = oldCell.RawValue,
                Formula = oldCell.Formula,
                Notes = oldCell.Notes,
                AutomationMode = oldCell.AutomationMode
            };
            newRow.Cells.Add(newCell);
        }
        return newRow;
    }
    
    private CellViewModel RecreateCellWithNewIndex(CellViewModel oldCell, int newRow, int newCol)
    {
        return new CellViewModel(_workbook, this, newRow, newCol)
        {
            RawValue = oldCell.RawValue,
            Formula = oldCell.Formula,
            Notes = oldCell.Notes,
            AutomationMode = oldCell.AutomationMode
        };
    }
}
