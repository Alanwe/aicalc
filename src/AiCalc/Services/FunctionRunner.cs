using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.ViewModels;

namespace AiCalc.Services;

public class FunctionRunner
{
    private static readonly Regex FunctionRegex = new(@"^=?(?<name>[A-Z0-9_]+)\((?<args>.*)\)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public FunctionRunner(FunctionRegistry registry)
    {
        Registry = registry;
    }

    public FunctionRegistry Registry { get; }

    public async Task<FunctionExecutionResult?> EvaluateAsync(CellViewModel cell, string formula)
    {
        if (!FunctionRegex.IsMatch(formula))
        {
            return null;
        }

        var match = FunctionRegex.Match(formula);
        var name = match.Groups["name"].Value;
        var args = match.Groups["args"].Value;

        if (!Registry.TryGet(name, out var descriptor))
        {
            throw new InvalidOperationException($"Unknown function '{name}'.");
        }

        var arguments = await ResolveArgumentsAsync(cell.Sheet, args);
        var context = new FunctionEvaluationContext(cell.Sheet.Workbook, cell.Sheet, arguments, formula);
        return await descriptor.Handler(context);
    }

    private async Task<IReadOnlyList<CellViewModel>> ResolveArgumentsAsync(SheetViewModel sheet, string args)
    {
        await Task.CompletedTask;
        var workbook = sheet.Workbook;
        var results = new List<CellViewModel>();

        foreach (var token in SplitArguments(args))
        {
            if (CellAddress.TryParse(token.Trim(), sheet.Name, out var address))
            {
                var targetSheet = workbook.GetSheet(address.SheetName) ?? sheet;
                var cell = targetSheet.GetCell(address.Row, address.Column);
                if (cell is not null)
                {
                    results.Add(cell);
                }
            }
            else if (double.TryParse(token, out var number))
            {
                var temp = new CellViewModel(workbook, sheet, 0, 0)
                {
                    Value = new CellValue(CellObjectType.Number, number.ToString(), number.ToString())
                };
                results.Add(temp);
            }
            else if (token.StartsWith('"') && token.EndsWith('"'))
            {
                var text = token.Trim('"');
                var temp = new CellViewModel(workbook, sheet, 0, 0)
                {
                    Value = new CellValue(CellObjectType.Text, text, text)
                };
                results.Add(temp);
            }
        }

        return results;
    }

    private static IEnumerable<string> SplitArguments(string args)
    {
        if (string.IsNullOrWhiteSpace(args))
        {
            yield break;
        }

        int depth = 0;
        var current = new List<char>();
        foreach (var ch in args)
        {
            if (ch == '(')
            {
                depth++;
            }
            else if (ch == ')')
            {
                depth--;
            }

            if (ch == ',' && depth == 0)
            {
                yield return new string(current.ToArray());
                current.Clear();
            }
            else
            {
                current.Add(ch);
            }
        }

        if (current.Count > 0)
        {
            yield return new string(current.ToArray());
        }
    }
}
