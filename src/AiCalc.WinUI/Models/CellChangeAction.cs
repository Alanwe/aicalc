using System;

namespace AiCalc.Models;

/// <summary>
/// Represents a reversible cell change action for undo/redo (Phase 5)
/// </summary>
public class CellChangeAction
{
    public required CellAddress Address { get; init; }
    public string? OldValue { get; init; }
    public string? NewValue { get; init; }
    public string? OldFormula { get; init; }
    public string? NewFormula { get; init; }
    public CellFormat? OldFormat { get; init; }
    public CellFormat? NewFormat { get; init; }
    public CellAutomationMode OldAutomationMode { get; init; }
    public CellAutomationMode NewAutomationMode { get; init; }
    public DateTime Timestamp { get; init; } = DateTime.Now;
    public string Description { get; init; } = "Cell Change";

    /// <summary>
    /// Create action for value/formula change
    /// </summary>
    public static CellChangeAction ForValueChange(
        CellAddress address,
        string? oldValue,
        string? newValue,
        string? oldFormula,
        string? newFormula,
        CellAutomationMode oldMode,
        CellAutomationMode newMode,
        string description = "Edit Cell")
    {
        return new CellChangeAction
        {
            Address = address,
            OldValue = oldValue,
            NewValue = newValue,
            OldFormula = oldFormula,
            NewFormula = newFormula,
            OldAutomationMode = oldMode,
            NewAutomationMode = newMode,
            Description = description
        };
    }

    /// <summary>
    /// Create action for format change
    /// </summary>
    public static CellChangeAction ForFormatChange(
        CellAddress address,
        CellFormat? oldFormat,
        CellFormat? newFormat,
        string description = "Format Cell")
    {
        return new CellChangeAction
        {
            Address = address,
            OldFormat = oldFormat,
            NewFormat = newFormat,
            Description = description
        };
    }

    /// <summary>
    /// Create action for automation mode change
    /// </summary>
    public static CellChangeAction ForModeChange(
        CellAddress address,
        CellAutomationMode oldMode,
        CellAutomationMode newMode,
        string description = "Change Mode")
    {
        return new CellChangeAction
        {
            Address = address,
            OldAutomationMode = oldMode,
            NewAutomationMode = newMode,
            Description = description
        };
    }
}
