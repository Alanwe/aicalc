using System;
using System.Collections.Generic;
using System.Linq;
using AiCalc.Models;

namespace AiCalc.Services;

public static class FormulaValidation
{
    public record ValidationResult(bool IsValid, string? ErrorMessage = null);

    public static ValidationResult ValidateParameters(FunctionDescriptor descriptor, IReadOnlyList<string> tokens, Func<CellAddress, CellObjectType?>? cellTypeLookup, string defaultSheet = "Sheet1")
    {
        if (descriptor == null) throw new ArgumentNullException(nameof(descriptor));
        if (tokens == null) throw new ArgumentNullException(nameof(tokens));

        var parameters = descriptor.Parameters.ToList();
        bool isVarArgs = parameters.Count > 0 && parameters[^1].Name == "...";
        var requiredCount = parameters.Count(p => !p.IsOptional && p.Name != "...");

        if (tokens.Count < requiredCount)
            return new ValidationResult(false, "Not enough arguments for function");

        int validateUpTo = Math.Min(tokens.Count, parameters.Count);
        if (isVarArgs && tokens.Count > parameters.Count)
            validateUpTo = tokens.Count;

        for (int i = 0; i < validateUpTo; i++)
        {
            var token = tokens[i];
            FunctionParameter? param = null;
            if (i < parameters.Count)
            {
                param = parameters[i];
            }
            else if (isVarArgs)
            {
                param = parameters.Last();
            }
            if (param == null) continue;

            // If token is quoted string
            if (token.StartsWith("\"") && token.EndsWith("\""))
            {
                if (!param.CanAccept(CellObjectType.Text))
                    return new ValidationResult(false, $"Parameter {i + 1} expects {param.ExpectedType}; got Text");
                continue;
            }

            // If numeric
            if (double.TryParse(token, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
            {
                if (!param.CanAccept(CellObjectType.Number))
                    return new ValidationResult(false, $"Parameter {i + 1} expects {param.ExpectedType}; got Number");
                continue;
            }

            // If looks like cell reference or sheet reference
            if (CellAddress.TryParse(token, defaultSheet, out var addr))
            {
                var type = cellTypeLookup?.Invoke(addr) ?? CellObjectType.Text;
                if (!param.CanAccept(type))
                    return new ValidationResult(false, $"Parameter {i + 1} expects {param.ExpectedType}; got {type}");
                continue;
            }

            // Unknown token treated as text
            if (!param.CanAccept(CellObjectType.Text))
                return new ValidationResult(false, $"Parameter {i + 1} expects {param.ExpectedType}; got Text");
        }

        return new ValidationResult(true, null);
    }
}
