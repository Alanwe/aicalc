using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.ViewModels;

namespace AiCalc.Services;

public class FunctionRegistry
{
    private readonly Dictionary<string, FunctionDescriptor> _functions = new(StringComparer.OrdinalIgnoreCase);

    public FunctionRegistry()
    {
        RegisterBuiltIns();
    }

    public IReadOnlyCollection<FunctionDescriptor> Functions => _functions.Values.OrderBy(f => f.Name).ToList();

    public bool TryGet(string name, out FunctionDescriptor descriptor) => _functions.TryGetValue(name, out descriptor!);

    public void Register(FunctionDescriptor descriptor)
    {
        _functions[descriptor.Name] = descriptor;
    }

    private void RegisterBuiltIns()
    {
        Register(new FunctionDescriptor(
            "SUM",
            "Adds a series of numbers.",
            async ctx =>
            {
                await Task.CompletedTask;
                var sum = ctx.Arguments.Sum(cell => double.TryParse(cell.DisplayValue, out var value) ? value : 0);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, sum.ToString(), sum.ToString()), $"Summed {ctx.Arguments.Count} cells");
            },
            new FunctionParameter("values", "Range of values to add.", CellObjectType.Number)));

        Register(new FunctionDescriptor(
            "CONCAT",
            "Concatenates string values.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = string.Join("", ctx.Arguments.Select(cell => cell.DisplayValue));
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, text, text));
            },
            new FunctionParameter("values", "Values to concatenate.", CellObjectType.Text)));

        Register(new FunctionDescriptor(
            "TEXT_TO_IMAGE",
            "Generates an image description placeholder for a prompt.",
            async ctx =>
            {
                await Task.CompletedTask;
                var prompt = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var metadata = $"image://generated/{Guid.NewGuid():N}.png";
                return new FunctionExecutionResult(new CellValue(CellObjectType.Image, metadata, $"Image from '{prompt}'"));
            },
            new FunctionParameter("prompt", "Prompt text", CellObjectType.Text)));

        Register(new FunctionDescriptor(
            "IMAGE_TO_TEXT",
            "Describes an image as text.",
            async ctx =>
            {
                await Task.CompletedTask;
                var description = $"Caption for {ctx.Arguments.FirstOrDefault()?.DisplayValue}";
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, description, description));
            },
            new FunctionParameter("image", "Image cell", CellObjectType.Image)));

        Register(new FunctionDescriptor(
            "DIRECTORY_TO_TABLE",
            "Expands a directory listing into a table value.",
            async ctx =>
            {
                await Task.CompletedTask;
                var path = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var json = $"[{{"name":"sample.txt","size":1024}},{{"name":"image.png","size":2048}}]";
                return new FunctionExecutionResult(new CellValue(CellObjectType.Table, json, $"Directory snapshot: {path}"));
            },
            new FunctionParameter("directory", "Directory path", CellObjectType.Directory)));
    }
}
