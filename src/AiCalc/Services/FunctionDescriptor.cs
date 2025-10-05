using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.ViewModels;

namespace AiCalc.Services;

public class FunctionDescriptor
{
    public FunctionDescriptor(string name, string description, Func<FunctionEvaluationContext, Task<FunctionExecutionResult>> handler, params FunctionParameter[] parameters)
    {
        Name = name;
        Description = description;
        Handler = handler;
        Parameters = parameters;
    }

    public string Name { get; }

    public string Description { get; }

    public IReadOnlyList<FunctionParameter> Parameters { get; }

    public Func<FunctionEvaluationContext, Task<FunctionExecutionResult>> Handler { get; }
}

public class FunctionParameter
{
    public FunctionParameter(string name, string description, CellObjectType expectedType, bool isOptional = false)
    {
        Name = name;
        Description = description;
        ExpectedType = expectedType;
        IsOptional = isOptional;
    }

    public string Name { get; }

    public string Description { get; }

    public CellObjectType ExpectedType { get; }

    public bool IsOptional { get; }
}

public record FunctionExecutionResult(CellValue Value, string? Diagnostics = null);

public record FunctionEvaluationContext(WorkbookViewModel Workbook, SheetViewModel Sheet, IReadOnlyList<CellViewModel> Arguments, string RawFormula);
