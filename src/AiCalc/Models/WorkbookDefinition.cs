using System.Collections.ObjectModel;

namespace AiCalc.Models;

public class WorkbookDefinition
{
    public string Title { get; set; } = "AiCalc Workbook";

    public ObservableCollection<SheetDefinition> Sheets { get; set; } = new();

    public WorkbookSettings Settings { get; set; } = new();
}
