using System.Collections.ObjectModel;

namespace AiCalc.ViewModels;

public class RowViewModel
{
    public RowViewModel(int rowIndex)
    {
        Index = rowIndex;
    }

    public int Index { get; }

    public string Label => (Index + 1).ToString();

    public ObservableCollection<CellViewModel> Cells { get; } = new();
}
