using System;

namespace AiCalc.Models;

/// <summary>
/// User preferences for UI state and settings (Phase 5)
/// </summary>
public class UserPreferences
{
    /// <summary>
    /// Application theme: Light or Dark
    /// </summary>
    public string Theme { get; set; } = "Dark";

    /// <summary>
    /// Width of the Functions panel (0 = collapsed)
    /// </summary>
    public double FunctionsPanelWidth { get; set; } = 280;

    /// <summary>
    /// Width of the Inspector panel (0 = collapsed)
    /// </summary>
    public double InspectorPanelWidth { get; set; } = 320;

    /// <summary>
    /// Window width (saved on close)
    /// </summary>
    public double WindowWidth { get; set; } = 1400;

    /// <summary>
    /// Window height (saved on close)
    /// </summary>
    public double WindowHeight { get; set; } = 900;

    /// <summary>
    /// Whether Functions panel is visible
    /// </summary>
    public bool FunctionsPanelVisible { get; set; } = true;

    /// <summary>
    /// Whether Inspector panel is visible
    /// </summary>
    public bool InspectorPanelVisible { get; set; } = true;

    /// <summary>
    /// Enter key behavior: MoveDown, MoveRight, or Stay
    /// </summary>
    public string EnterKeyBehavior { get; set; } = "MoveDown";

    /// <summary>
    /// Enable formula autocomplete
    /// </summary>
    public bool EnableFormulaAutocomplete { get; set; } = true;

    /// <summary>
    /// Enable formula syntax highlighting
    /// </summary>
    public bool EnableFormulaSyntaxHighlighting { get; set; } = true;

    /// <summary>
    /// Recent workbook paths (up to 10)
    /// </summary>
    public string[] RecentWorkbooks { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Last opened workbook path
    /// </summary>
    public string? LastWorkbookPath { get; set; }

    /// <summary>
    /// Enable automatic saving (Phase 6)
    /// </summary>
    public bool AutoSaveEnabled { get; set; } = true;

    /// <summary>
    /// AutoSave interval in minutes (Phase 6)
    /// </summary>
    public int AutoSaveIntervalMinutes { get; set; } = 5;

    /// <summary>
    /// Selected Python environment path (Phase 7)
    /// </summary>
    public string? PythonEnvironmentPath { get; set; }

    /// <summary>
    /// Enable Python SDK bridge (Phase 7)
    /// </summary>
    public bool PythonBridgeEnabled { get; set; } = true;

    /// <summary>
    /// Python functions directory path (Phase 7 - Task 21)
    /// </summary>
    public string? PythonFunctionsDirectory { get; set; }

    /// <summary>
    /// Enable hot reload for Python functions (Phase 7 - Task 21)
    /// </summary>
    public bool PythonHotReloadEnabled { get; set; } = true;
}
