namespace AiCalc.Models;

/// <summary>
/// Represents a function suggestion for intellisense
/// </summary>
public class FunctionSuggestion
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Signature { get; set; } = string.Empty;
}
