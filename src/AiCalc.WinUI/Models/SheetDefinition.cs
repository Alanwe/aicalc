using System.Collections.ObjectModel;

namespace AiCalc.Models;

public class SheetDefinition
{
    public string Name { get; set; } = "Sheet1";

    public ObservableCollection<CellDefinition> Cells { get; set; } = new();

    public int ColumnCount { get; set; } = 8;

    public int RowCount { get; set; } = 12;
}
