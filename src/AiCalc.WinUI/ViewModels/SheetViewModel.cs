using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using AiCalc.Models;

namespace AiCalc.ViewModels;

public class SheetViewModel
{
    private readonly WorkbookViewModel _workbook;
    public const double DefaultColumnWidth = 120d;
    
    private readonly ObservableCollection<double> _columnWidths;
    private readonly ObservableCollection<bool> _columnVisibility;
    private readonly ObservableCollection<bool> _rowVisibility;
    
    // Freeze panes support (Phase 8)
    private int _frozenColumnCount = 0;
    private int _frozenRowCount = 0;

    public SheetViewModel(WorkbookViewModel workbook, string name, int rows, int columns)
    {
        _workbook = workbook;
        Workbook = workbook;
        Name = name;
        Rows = new ObservableCollection<RowViewModel>(Enumerable.Range(0, rows).Select(r => CreateRow(r, columns)));
        _columnWidths = new ObservableCollection<double>(Enumerable.Repeat(DefaultColumnWidth, columns));
        _columnVisibility = new ObservableCollection<bool>(Enumerable.Repeat(true, columns));
        _rowVisibility = new ObservableCollection<bool>(Enumerable.Repeat(true, rows));
    }

    public WorkbookViewModel Workbook { get; }

    public string Name { get; }

    public ObservableCollection<RowViewModel> Rows { get; }

    public int ColumnCount => Rows.FirstOrDefault()?.Cells.Count ?? 0;

    public ObservableCollection<double> ColumnWidths => _columnWidths;

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

    // ==================== Visibility Support (Phase 8) ====================
    
    public IReadOnlyList<int> GetVisibleColumnIndices() => _columnVisibility
        .Select((visible, index) => (visible, index))
        .Where(pair => pair.visible)
        .Select(pair => pair.index)
        .ToList();

    public IReadOnlyList<int> GetVisibleRowIndices() => _rowVisibility
        .Select((visible, index) => (visible, index))
        .Where(pair => pair.visible)
        .Select(pair => pair.index)
        .ToList();

    public bool IsColumnVisible(int index) => index >= 0 && index < _columnVisibility.Count && _columnVisibility[index];

    public bool IsRowVisible(int index) => index >= 0 && index < _rowVisibility.Count && _rowVisibility[index];

    public void SetColumnVisibility(int index, bool isVisible)
    {
        if (index < 0 || index >= _columnVisibility.Count)
        {
            return;
        }
        _columnVisibility[index] = isVisible;
    }

    public void SetRowVisibility(int index, bool isVisible)
    {
        if (index < 0 || index >= _rowVisibility.Count)
        {
            return;
        }
        _rowVisibility[index] = isVisible;
    }

    public void ShowAllColumns()
    {
        for (int i = 0; i < _columnVisibility.Count; i++)
        {
            _columnVisibility[i] = true;
        }
    }

    public void ShowAllRows()
    {
        for (int i = 0; i < _rowVisibility.Count; i++)
        {
            _rowVisibility[i] = true;
        }
    }

    public int VisibleColumnCount => _columnVisibility.Count(v => v);

    public int VisibleRowCount => _rowVisibility.Count(v => v);
    
    // Freeze panes properties
    public int FrozenColumnCount
    {
        get => _frozenColumnCount;
        set => _frozenColumnCount = Math.Max(0, Math.Min(value, ColumnCount));
    }
    
    public int FrozenRowCount
    {
        get => _frozenRowCount;
        set => _frozenRowCount = Math.Max(0, Math.Min(value, Rows.Count));
    }
    
    public bool HasFrozenPanes => _frozenColumnCount > 0 || _frozenRowCount > 0;

    public void EnsureCapacity(int minimumRows, int minimumColumns)
    {
        EnsureRowCapacity(minimumRows);
        EnsureColumnCapacity(minimumColumns);
    }

    public IReadOnlyList<CellViewModel> GetCellsInRange(CellViewModel origin, int height, int width)
    {
        var cells = new List<CellViewModel>();
        for (int r = 0; r < height; r++)
        {
            var rowIndex = origin.Row + r;
            if (rowIndex >= Rows.Count)
            {
                break;
            }

            var row = Rows[rowIndex];
            for (int c = 0; c < width; c++)
            {
                var columnIndex = origin.Column + c;
                if (columnIndex >= row.Cells.Count)
                {
                    break;
                }

                cells.Add(row.Cells[columnIndex]);
            }
        }

        return cells;
    }

    public List<CellViewModel> ApplySpill(CellViewModel origin, CellValue[,] values)
    {
        var affected = new List<CellViewModel>();
        var requiredRows = origin.Row + values.GetLength(0);
        var requiredColumns = origin.Column + values.GetLength(1);

        EnsureCapacity(requiredRows, requiredColumns);

        for (int r = 0; r < values.GetLength(0); r++)
        {
            for (int c = 0; c < values.GetLength(1); c++)
            {
                var rowIndex = origin.Row + r;
                var columnIndex = origin.Column + c;
                var cell = Rows[rowIndex].Cells[columnIndex];

                if (cell == origin)
                {
                    continue;
                }

                using (cell.SuppressHistory())
                {
                    cell.Value = values[r, c];
                    cell.Formula = null;
                    cell.AutomationMode = CellAutomationMode.Manual;
                }

                cell.MarkAsUpdated();
                affected.Add(cell);
            }
        }

        return affected;
    }

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

    private void EnsureRowCapacity(int minimumRows)
    {
        while (Rows.Count < minimumRows)
        {
            var newRow = CreateRow(Rows.Count, ColumnCount);
            Rows.Add(newRow);
            _rowVisibility.Add(true);
        }
    }

    private void EnsureColumnCapacity(int minimumColumns)
    {
        while (ColumnCount < minimumColumns)
        {
            foreach (var row in Rows)
            {
                var newCell = new CellViewModel(_workbook, this, row.Index, row.Cells.Count);
                row.Cells.Add(newCell);
            }
            _columnWidths.Add(DefaultColumnWidth);
            _columnVisibility.Add(true);
        }
    }
    
    // ==================== Row/Column Operations (Phase 5/8) ====================
    
    public void InsertRow(int index)
    {
        if (index < 0 || index > Rows.Count) return;
        
        var newRow = CreateRow(index, ColumnCount);
        Rows.Insert(index, newRow);
        _rowVisibility.Insert(index, true);
        
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
        if (index < _rowVisibility.Count)
        {
            _rowVisibility.RemoveAt(index);
        }
        
        // Update row indices for all rows after the deleted row
        for (int r = index; r < Rows.Count; r++)
        {
            Rows[r] = RecreateRowWithNewIndex(Rows[r], r);
        }
    }
    
    public void InsertColumn(int index)
    {
        if (index < 0 || index > ColumnCount) return;
        
        _columnWidths.Insert(index, DefaultColumnWidth);
        _columnVisibility.Insert(index, true);
        
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
        
        if (index < _columnWidths.Count)
        {
            _columnWidths.RemoveAt(index);
        }
        if (index < _columnVisibility.Count)
        {
            _columnVisibility.RemoveAt(index);
        }
        
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
            var newCell = new CellViewModel(_workbook, this, newIndex, c);
            newCell.CopyFrom(oldCell);
            newRow.Cells.Add(newCell);
        }
        return newRow;
    }
    
    private CellViewModel RecreateCellWithNewIndex(CellViewModel oldCell, int newRow, int newCol)
    {
        var newCell = new CellViewModel(_workbook, this, newRow, newCol);
        newCell.CopyFrom(oldCell);
        return newCell;
    }
}
