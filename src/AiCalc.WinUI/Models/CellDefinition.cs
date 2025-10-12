using System.Text.Json.Serialization;

namespace AiCalc.Models;

public class CellDefinition
{
    public string Address { get; set; } = string.Empty;

    public string? Formula { get; set; }

    public CellValue Value { get; set; } = CellValue.Empty;

    public CellFormat Format { get; set; } = CellFormat.Default;

    public System.Collections.ObjectModel.ObservableCollection<CellHistoryEntry> History { get; set; } = new();

    public CellAutomationMode AutomationMode { get; set; } = CellAutomationMode.Manual;

    public string? Notes { get; set; }

    public string? SourcePath { get; set; }
}
