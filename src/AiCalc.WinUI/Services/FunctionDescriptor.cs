using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.Services.AI;
using AiCalc.ViewModels;

namespace AiCalc.Services;

/// <summary>
/// Categories for organizing functions
/// </summary>
public enum FunctionCategory
{
    Math,
    Text,
    DateTime,
    File,
    Directory,
    Table,
    Image,
    Video,
    Pdf,
    Data,
    AI,
    Contrib
}

public class FunctionDescriptor
{
    public FunctionDescriptor(
        string name, 
        string description, 
        Func<FunctionEvaluationContext, Task<FunctionExecutionResult>> handler, 
        FunctionCategory category = FunctionCategory.Math,
        params FunctionParameter[] parameters)
    {
        Name = name;
        Description = description;
        Handler = handler;
        Category = category;
        Parameters = parameters;
        
        // Extract applicable cell types from parameters
        var types = new HashSet<CellObjectType>();
        foreach (var param in parameters)
        {
            foreach (var type in param.AcceptableTypes)
            {
                types.Add(type);
            }
        }
        ApplicableTypes = types.ToArray();
    }

    public string Name { get; }

    public string Description { get; }

    public FunctionCategory Category { get; }

    public IReadOnlyList<FunctionParameter> Parameters { get; }
    
    /// <summary>
    /// Cell types that this function can accept as input
    /// </summary>
    public CellObjectType[] ApplicableTypes { get; }

    public Func<FunctionEvaluationContext, Task<FunctionExecutionResult>> Handler { get; }

    /// <summary>
    /// Describes the typical result or return value for this function.
    /// </summary>
    public string? ExpectedOutput { get; set; }

    /// <summary>
    /// Optional example formula that demonstrates how to use the function.
    /// </summary>
    public string? Example { get; set; }

    /// <summary>
    /// Indicates the primary cell type produced when the function succeeds.
    /// </summary>
    public CellObjectType? ResultType { get; set; }

    public bool HasExample => !string.IsNullOrWhiteSpace(Example);

    public bool HasExpectedOutput => !string.IsNullOrWhiteSpace(ExpectedOutput);
    
    /// <summary>
    /// Check if this function can accept the given cell types
    /// </summary>
    public bool CanAccept(params CellObjectType[] types)
    {
        if (types.Length == 0) return true;
        return types.All(t => ApplicableTypes.Contains(t));
    }
}

public class FunctionParameter
{
    public FunctionParameter(
        string name, 
        string description, 
        CellObjectType expectedType, 
        bool isOptional = false,
        params CellObjectType[] additionalAcceptableTypes)
    {
        Name = name;
        Description = description;
        ExpectedType = expectedType;
        IsOptional = isOptional;
        
        // Build list of acceptable types
        var types = new List<CellObjectType> { expectedType };
        if (additionalAcceptableTypes != null && additionalAcceptableTypes.Length > 0)
        {
            types.AddRange(additionalAcceptableTypes);
        }
        AcceptableTypes = types.ToArray();
    }

    public string Name { get; }

    public string Description { get; }

    public CellObjectType ExpectedType { get; }

    public bool IsOptional { get; }
    
    /// <summary>
    /// All cell types that this parameter can accept (supports polymorphism)
    /// </summary>
    public CellObjectType[] AcceptableTypes { get; }
    
    /// <summary>
    /// Check if this parameter can accept the given cell type
    /// </summary>
    public bool CanAccept(CellObjectType type)
    {
        return AcceptableTypes.Contains(type);
    }
}

public record FunctionExecutionResult(
    CellValue Value,
    string? Diagnostics = null,
    CellValue[,]? SpillRange = null,
    IReadOnlyList<CellAddress>? ReferencedCells = null,
    AIResponse? AiResponse = null)
{
    public bool HasSpill => SpillRange is { Length: > 0 };
}

public record FunctionEvaluationContext(WorkbookViewModel Workbook, SheetViewModel Sheet, IReadOnlyList<CellViewModel> Arguments, string RawFormula);
