using System;
using System.Collections.ObjectModel;

namespace AiCalc.Models;

/// <summary>
/// Application theme options (Task 10)
/// </summary>
public enum AppTheme
{
    System,
    Light,
    Dark
}

/// <summary>
/// Cell visual state theme options (Task 10)
/// </summary>
public enum CellVisualTheme
{
    Light,
    Dark,
    HighContrast,
    Custom
}

public class WorkbookSettings
{
    public ObservableCollection<WorkspaceConnection> Connections { get; set; } = new();

    public string DefaultImageModel { get; set; } = "stable-diffusion";

    public string DefaultTextModel { get; set; } = "gpt-4";

    public string DefaultVideoModel { get; set; } = "gen-2";
    
    public string? WorkspacePath { get; set; }
    
    public bool AutoSave { get; set; } = false;

    // Evaluation Settings (Task 9)
    public int MaxEvaluationThreads { get; set; } = Environment.ProcessorCount;
    
    public int DefaultEvaluationTimeoutSeconds { get; set; } = 100;

    // Appearance Settings (Task 10)
    public AppTheme ApplicationTheme { get; set; } = AppTheme.System;
    
    public CellVisualTheme SelectedTheme { get; set; } = CellVisualTheme.Light;

    // History / UI (Phase 5 enhancements)
    public int MaxHistoryEntries { get; set; } = 100;

    public double FunctionsPanelWidth { get; set; } = 280;

    public double InspectorPanelWidth { get; set; } = 320;

    public bool EnterMovesDown { get; set; } = true;
}
