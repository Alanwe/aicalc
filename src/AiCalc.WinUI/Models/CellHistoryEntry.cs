using System;

namespace AiCalc.Models;

/// <summary>
/// Represents a single entry in the cell change history (Phase 5 Task 16).
/// </summary>
public class CellHistoryEntry
{
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;

    public CellValue OldValue { get; set; } = CellValue.Empty;

    public CellValue NewValue { get; set; } = CellValue.Empty;

    public string? OldFormula { get; set; }

    public string? NewFormula { get; set; }

    public string? Notes { get; set; }

    public string Author { get; set; } = Environment.UserName;

    public string Summary => $"{Timestamp:G} • {Author} • {Describe()}";

    public CellHistoryEntry Clone()
    {
        return new CellHistoryEntry
        {
            Timestamp = Timestamp,
            OldValue = OldValue,
            NewValue = NewValue,
            OldFormula = OldFormula,
            NewFormula = NewFormula,
            Notes = Notes,
            Author = Author
        };
    }

    private string Describe()
    {
        if (!string.Equals(OldFormula, NewFormula, StringComparison.Ordinal))
        {
            return $"Formula updated to '{NewFormula}'";
        }

        if (!string.Equals(OldValue.SerializedValue, NewValue.SerializedValue, StringComparison.Ordinal))
        {
            return $"Value changed from '{OldValue.DisplayValue ?? OldValue.SerializedValue}' to '{NewValue.DisplayValue ?? NewValue.SerializedValue}'";
        }

        return "Cell updated";
    }
}
