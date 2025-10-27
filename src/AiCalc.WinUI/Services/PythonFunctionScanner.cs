using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.ViewModels;

namespace AiCalc.Services;

/// <summary>
/// Scans Python files for @aicalc_function decorated functions
/// and extracts metadata for registration in AiCalc.
/// </summary>
public class PythonFunctionScanner
{
    private readonly string _pythonExecutablePath;
    private readonly string _discoverScriptPath;

    public PythonFunctionScanner(string pythonExecutablePath)
    {
        _pythonExecutablePath = pythonExecutablePath;
        
        // discover_functions.py should be in same directory as aicalc_sdk
        var sdkPath = Path.GetDirectoryName(typeof(PythonFunctionScanner).Assembly.Location);
        _discoverScriptPath = Path.Combine(sdkPath!, "python-sdk", "aicalc_sdk", "discover_functions.py");
        
        // Fallback: check relative to project root
        if (!File.Exists(_discoverScriptPath))
        {
            var projectRoot = Path.GetFullPath(Path.Combine(sdkPath!, "..", "..", "..", ".."));
            _discoverScriptPath = Path.Combine(projectRoot, "python-sdk", "aicalc_sdk", "discover_functions.py");
        }
    }

    /// <summary>
    /// Scans a directory for Python files and discovers all @aicalc_function decorated functions.
    /// </summary>
    public async Task<List<PythonFunctionInfo>> ScanDirectoryAsync(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
            return new List<PythonFunctionInfo>();

        var allFunctions = new List<PythonFunctionInfo>();

        // Find all .py files in directory and subdirectories
        var pythonFiles = Directory.GetFiles(directoryPath, "*.py", SearchOption.AllDirectories);

        foreach (var filePath in pythonFiles)
        {
            // Skip __init__.py and __pycache__
            if (filePath.Contains("__pycache__") || Path.GetFileName(filePath) == "__init__.py")
                continue;

            var functions = await ScanFileAsync(filePath);
            allFunctions.AddRange(functions);
        }

        return allFunctions;
    }

    /// <summary>
    /// Scans a single Python file for @aicalc_function decorated functions.
    /// </summary>
    public async Task<List<PythonFunctionInfo>> ScanFileAsync(string filePath)
    {
        if (!File.Exists(filePath))
            return new List<PythonFunctionInfo>();

        if (!File.Exists(_discoverScriptPath))
        {
            throw new FileNotFoundException($"Discovery script not found: {_discoverScriptPath}");
        }

        try
        {
            // Run Python discovery script
            var startInfo = new ProcessStartInfo
            {
                FileName = _pythonExecutablePath,
                Arguments = $"\"{_discoverScriptPath}\" \"{filePath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(filePath)
            };

            using var process = Process.Start(startInfo);
            if (process == null)
                return new List<PythonFunctionInfo>();

            var output = await process.StandardOutput.ReadToEndAsync();
            var error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
            {
                Console.WriteLine($"[PythonFunctionScanner] Error scanning {filePath}: {error}");
                return new List<PythonFunctionInfo>();
            }

            // Parse JSON output
            var result = JsonSerializer.Deserialize<DiscoveryResult>(output, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (result == null || !result.Success || result.Functions == null)
            {
                if (!string.IsNullOrEmpty(result?.Error))
                    Console.WriteLine($"[PythonFunctionScanner] Discovery error: {result.Error}");
                return new List<PythonFunctionInfo>();
            }

            return result.Functions;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[PythonFunctionScanner] Exception scanning {filePath}: {ex.Message}");
            return new List<PythonFunctionInfo>();
        }
    }

    /// <summary>
    /// Creates a FunctionDescriptor from a PythonFunctionInfo for registration.
    /// </summary>
    public static FunctionDescriptor CreateDescriptor(PythonFunctionInfo info, string pythonPath)
    {
        var parameters = info.Parameters.Select(p =>
        {
            var paramType = MapPythonTypeToCellType(p.Type);
            return new FunctionParameter(
                name: p.Name,
                description: $"Parameter {p.Name}",
                expectedType: paramType,
                isOptional: !p.Required
            );
        }).ToArray();

        var category = info.Category.ToLower() switch
        {
            "math" => FunctionCategory.Math,
            "text" => FunctionCategory.Text,
            "datetime" => FunctionCategory.DateTime,
            "file" => FunctionCategory.File,
            "directory" => FunctionCategory.Directory,
            "table" => FunctionCategory.Table,
            "image" => FunctionCategory.Image,
            "video" => FunctionCategory.Video,
            "pdf" => FunctionCategory.Pdf,
            "data" => FunctionCategory.Data,
            "ai" => FunctionCategory.AI,
            _ => FunctionCategory.Contrib
        };

        var descriptor = new FunctionDescriptor(
            name: info.Name,
            description: info.Description,
            handler: async context =>
            {
                if (!TryBuildArgumentPayload(info, context, out var argumentsJson, out var validationError))
                {
                    return validationError ?? CreateErrorResult("Unable to build Python argument payload.");
                }

                return await ExecutePythonFunctionAsync(info, pythonPath, argumentsJson);
            },
            category: category,
            parameters: parameters
        )
        {
            ResultType = MapPythonTypeToCellType(info.ReturnType),
            ExpectedOutput = string.IsNullOrWhiteSpace(info.ReturnType)
                ? null
                : $"Returns a {info.ReturnType} value.",
            Example = info.Examples?.FirstOrDefault()
        };

        return descriptor;
    }

    /// <summary>
    /// Maps Python type string to CellObjectType.
    /// </summary>
    private static CellObjectType MapPythonTypeToCellType(string? pythonType)
    {
        return pythonType?.ToLower() switch
        {
            "str" or "string" or "text" => CellObjectType.Text,
            "int" or "float" or "number" or "double" => CellObjectType.Number,
            "bool" or "boolean" => CellObjectType.Text,
            "list" or "array" or "table" => CellObjectType.Table,
            "dict" or "dictionary" or "json" => CellObjectType.Json,
            _ => CellObjectType.Text
        };
    }

    private static bool TryBuildArgumentPayload(PythonFunctionInfo info, FunctionEvaluationContext context, out string argumentsJson, out FunctionExecutionResult? errorResult)
    {
        argumentsJson = "[]";
        errorResult = null;

        var parameters = info.Parameters ?? new List<ParameterInfo>();
        var providedArguments = context.Arguments ?? Array.Empty<CellViewModel>();

        var requiredCount = parameters.Count(p => p.Required);
        if (providedArguments.Count < requiredCount)
        {
            errorResult = CreateErrorResult($"{info.Name} expects at least {requiredCount} argument(s) but received {providedArguments.Count}.");
            return false;
        }

        if (parameters.Count > 0 && providedArguments.Count > parameters.Count)
        {
            errorResult = CreateErrorResult($"{info.Name} accepts at most {parameters.Count} argument(s).");
            return false;
        }

        var payload = new List<object?>();

        for (var index = 0; index < providedArguments.Count; index++)
        {
            var parameter = parameters.Count > index ? parameters[index] : null;
            var cell = providedArguments[index];

            if (!TryConvertCellArgument(cell, parameter, index, out var converted, out var message))
            {
                errorResult = CreateErrorResult(message);
                return false;
            }

            payload.Add(converted);
        }

        argumentsJson = JsonSerializer.Serialize(payload);
        return true;
    }

    private static bool TryConvertCellArgument(CellViewModel cell, ParameterInfo? parameter, int index, out object? value, out string errorMessage)
    {
        var parameterLabel = string.IsNullOrWhiteSpace(parameter?.Name)
            ? $"argument #{index + 1}"
            : $"'{parameter!.Name}'";

        var raw = cell.Value.SerializedValue ?? cell.Value.DisplayValue ?? cell.RawValue ?? string.Empty;
        var normalizedType = parameter?.Type?.ToLowerInvariant() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(raw))
        {
            if (parameter != null && parameter.Required)
            {
                errorMessage = $"{parameterLabel} is required.";
                value = null;
                return false;
            }

            value = normalizedType switch
            {
                "list" or "array" or "table" => JsonNode.Parse("[]"),
                "dict" or "dictionary" or "json" => JsonNode.Parse("{}"),
                _ => string.Empty
            };

            errorMessage = string.Empty;
            return true;
        }

        switch (normalizedType)
        {
            case "int" or "float" or "number" or "double":
                if (TryGetNumericValue(raw, out var numericValue))
                {
                    value = numericValue;
                    errorMessage = string.Empty;
                    return true;
                }

                errorMessage = $"{parameterLabel} must be a number.";
                value = null;
                return false;

            case "bool" or "boolean":
                if (TryGetBooleanValue(raw, out var boolValue))
                {
                    value = boolValue;
                    errorMessage = string.Empty;
                    return true;
                }

                errorMessage = $"{parameterLabel} must be TRUE or FALSE.";
                value = null;
                return false;

            case "list" or "array" or "table":
                if (TryParseJsonNode(raw, out var arrayNode) && arrayNode is JsonArray)
                {
                    value = arrayNode;
                    errorMessage = string.Empty;
                    return true;
                }

                errorMessage = $"{parameterLabel} must be a JSON array.";
                value = null;
                return false;

            case "dict" or "dictionary" or "json":
                if (TryParseJsonNode(raw, out var objectNode) && objectNode is JsonObject or JsonArray)
                {
                    value = objectNode;
                    errorMessage = string.Empty;
                    return true;
                }

                errorMessage = $"{parameterLabel} must be valid JSON.";
                value = null;
                return false;

            default:
                value = raw;
                errorMessage = string.Empty;
                return true;
        }
    }

    private static bool TryParseJsonNode(string json, out JsonNode? node)
    {
        node = null;

        if (string.IsNullOrWhiteSpace(json))
        {
            return false;
        }

        try
        {
            node = JsonNode.Parse(json);
            return node != null;
        }
        catch
        {
            node = null;
            return false;
        }
    }

    private static bool TryGetNumericValue(string input, out double value)
    {
        if (double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
        {
            return true;
        }

        if (double.TryParse(input, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
        {
            return true;
        }

        return false;
    }

    private static bool TryGetBooleanValue(string input, out bool value)
    {
        if (bool.TryParse(input, out value))
        {
            return true;
        }

        var normalized = input.Trim().ToLowerInvariant();
        if (normalized is "yes" or "y" or "1")
        {
            value = true;
            return true;
        }

        if (normalized is "no" or "n" or "0")
        {
            value = false;
            return true;
        }

        value = false;
        return false;
    }

    private static FunctionExecutionResult CreateErrorResult(string message)
    {
        var errorValue = new CellValue(CellObjectType.Error, message, message);
        return new FunctionExecutionResult(errorValue, message);
    }

    private static CellValue ConvertPythonOutput(string output, PythonFunctionInfo info)
    {
        var trimmed = output?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return CellValue.Empty;
        }

        var normalizedReturn = info.ReturnType?.ToLowerInvariant() ?? string.Empty;
        var targetType = MapPythonTypeToCellType(info.ReturnType);

        if (targetType == CellObjectType.Number && TryGetNumericValue(trimmed, out var numericValue))
        {
            var serialized = numericValue.ToString(CultureInfo.InvariantCulture);
            return new CellValue(CellObjectType.Number, serialized, serialized);
        }

        if ((targetType == CellObjectType.Json || targetType == CellObjectType.Table) && TryParseJsonNode(trimmed, out var jsonNode) && jsonNode != null)
        {
            var serializedJson = jsonNode.ToJsonString(new JsonSerializerOptions { WriteIndented = false });
            return new CellValue(targetType, serializedJson, serializedJson);
        }

        if (normalizedReturn is "bool" or "boolean" && TryGetBooleanValue(trimmed, out var boolValue))
        {
            var text = boolValue ? "TRUE" : "FALSE";
            return new CellValue(CellObjectType.Text, text, text);
        }

        return new CellValue(CellObjectType.Text, trimmed, trimmed);
    }

    private static async Task<FunctionExecutionResult> ExecutePythonFunctionAsync(
        PythonFunctionInfo info,
        string pythonPath,
        string argumentsJson)
    {
        try
        {
            var scriptDirectory = Path.GetDirectoryName(info.FilePath);
            if (string.IsNullOrEmpty(scriptDirectory))
            {
                scriptDirectory = Directory.GetCurrentDirectory();
            }

            var moduleName = Path.GetFileNameWithoutExtension(info.FilePath);

            var pythonScript = $@"import json
import os
import sys

sys.path.insert(0, '{scriptDirectory}')

from {moduleName} import {info.FunctionName}

try:
    args_json = os.environ.get('AICALC_ARGS', '[]')
    args = json.loads(args_json)
    result = {info.FunctionName}(*args)
    if isinstance(result, (dict, list)):
        print(json.dumps(result, ensure_ascii=False))
    elif result is None:
        print('')
    else:
        print(result)
except Exception as e:
    print(f'ERROR: {{e}}', file=sys.stderr)
    sys.exit(1)
";

            var startInfo = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = $"-c \"{pythonScript.Replace("\"", "\\\"")}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = scriptDirectory
            };

            startInfo.EnvironmentVariables["AICALC_ARGS"] = argumentsJson;

            using var process = Process.Start(startInfo);
            if (process == null)
            {
                return CreateErrorResult("Failed to start Python process.");
            }

            var output = await process.StandardOutput.ReadToEndAsync();
            var error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
            {
                var message = string.IsNullOrWhiteSpace(error)
                    ? "Python process returned a non-zero exit code."
                    : error.Trim();
                return CreateErrorResult(message);
            }

            var resultValue = ConvertPythonOutput(output, info);
            return new FunctionExecutionResult(resultValue);
        }
        catch (Exception ex)
        {
            return CreateErrorResult($"Execution error: {ex.Message}");
        }
    }

    // JSON models for discovery script output
    private class DiscoveryResult
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("functions")]
        public List<PythonFunctionInfo>? Functions { get; set; }

        [JsonPropertyName("error")]
        public string? Error { get; set; }
    }
}

/// <summary>
/// Metadata for a Python function discovered by the scanner.
/// </summary>
public class PythonFunctionInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("category")]
    public string Category { get; set; } = "Python";

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;

    [JsonPropertyName("file_path")]
    public string FilePath { get; set; } = string.Empty;

    [JsonPropertyName("function_name")]
    public string FunctionName { get; set; } = string.Empty;

    [JsonPropertyName("parameters")]
    public List<ParameterInfo> Parameters { get; set; } = new();

    [JsonPropertyName("return_type")]
    public string ReturnType { get; set; } = "any";

    [JsonPropertyName("examples")]
    public List<string> Examples { get; set; } = new();
}

/// <summary>
/// Parameter metadata for a Python function.
/// </summary>
public class ParameterInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("type")]
    public string Type { get; set; } = "any";

    [JsonPropertyName("required")]
    public bool Required { get; set; }

    [JsonPropertyName("default")]
    public object? Default { get; set; }
}
