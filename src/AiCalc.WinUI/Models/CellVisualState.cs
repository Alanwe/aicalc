namespace AiCalc.Models;

/// <summary>
/// Visual state of a cell for UI feedback
/// </summary>
public enum CellVisualState
{
    /// <summary>
    /// Normal state - no special visual indication
    /// </summary>
    Normal,
    
    /// <summary>
    /// Cell value just changed - flash green border for 2 seconds
    /// </summary>
    JustUpdated,
    
    /// <summary>
    /// Cell is currently being calculated - show spinner
    /// </summary>
    Calculating,
    
    /// <summary>
    /// Cell needs recalculation (dependency changed) - show blue border
    /// </summary>
    Stale,
    
    /// <summary>
    /// Manual update mode - show orange indicator
    /// </summary>
    ManualUpdate,
    
    /// <summary>
    /// Cell has an error - show red border
    /// </summary>
    Error,
    
    /// <summary>
    /// Cell is in a dependency chain of selected cell - highlight
    /// </summary>
    InDependencyChain
}
