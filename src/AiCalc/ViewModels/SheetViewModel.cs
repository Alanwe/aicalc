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

    public async Task EvaluateAllAsync()
    {
        foreach (var cell in Cells.Where(c => c.HasFormula))
        {
            if (cell.AutomationMode != CellAutomationMode.Manual)
            {
                await cell.EvaluateAsync();
            }
        }
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
}
