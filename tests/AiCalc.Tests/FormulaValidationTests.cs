using AiCalc.Services;
using AiCalc.Models;
using Xunit;

namespace AiCalc.Tests;

public class FormulaValidationTests
{
    [Fact]
    public void ValidateParameters_SimpleTypeMismatch()
    {
        var param = new FunctionParameter("value", "number", CellObjectType.Number);
        var desc = new FunctionDescriptor("FAKE", "desc", null!, FunctionCategory.Math, param);
        var tokens = new[] { "\"hello\"" } as string[];
        var result = FormulaValidation.ValidateParameters(desc, tokens, null, "Sheet1");
        Assert.False(result.IsValid);
    }

    [Fact]
    public void ValidateParameters_VarArgsAcceptsMultiple()
    {
        var p1 = new FunctionParameter("first", "first", CellObjectType.Number);
        var p2 = new FunctionParameter("...", "rest", CellObjectType.Number);
        var desc = new FunctionDescriptor("SUM", "sum", null!, FunctionCategory.Math, p1, p2);
        var tokens = new[] { "1", "2", "3" };
        var result = FormulaValidation.ValidateParameters(desc, tokens, null, "Sheet1");
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ValidateParameters_OptionalParam()
    {
        var p1 = new FunctionParameter("a", "a", CellObjectType.Number);
        var p2 = new FunctionParameter("b", "b", CellObjectType.Number, isOptional: true);
        var desc = new FunctionDescriptor("ADD", "add", null!, FunctionCategory.Math, p1, p2);
        var tokens = new[] { "1" };
        var result = FormulaValidation.ValidateParameters(desc, tokens, null, "Sheet1");
        Assert.True(result.IsValid);
    }
}
