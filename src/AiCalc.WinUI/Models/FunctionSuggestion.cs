using System;
using AiCalc.Services;

namespace AiCalc.Models;

/// <summary>
/// Represents a function suggestion for intellisense with additional metadata.
/// </summary>
public class FunctionSuggestion
{
    public string Name { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    public string Signature { get; set; } = string.Empty;

    public FunctionCategory Category { get; set; } = FunctionCategory.Math;

    public string CategoryGlyph { get; set; } = string.Empty;

    public string CategoryLabel { get; set; } = string.Empty;

    public string TypeHint { get; set; } = string.Empty;

    public string ProviderHint { get; set; } = string.Empty;

    public bool HasTypeHint => !string.IsNullOrWhiteSpace(TypeHint);

    public bool HasProviderHint => !string.IsNullOrWhiteSpace(ProviderHint);
}
