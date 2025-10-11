using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.Models.CellObjects;
using AiCalc.Services.AI;
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
        
        // Check if this is an AI function
        if (descriptor.Category == FunctionCategory.AI)
        {
            return await ExecuteAIFunctionAsync(name, arguments, context);
        }
        
        return await descriptor.Handler(context);
    }
    
    /// <summary>
    /// Executes AI functions by routing to the appropriate AI service
    /// </summary>
    private async Task<FunctionExecutionResult> ExecuteAIFunctionAsync(string functionName, IReadOnlyList<CellViewModel> arguments, FunctionEvaluationContext context)
    {
        try
        {
            var client = App.AIServices.GetDefaultClient();
            if (client == null)
            {
                return new FunctionExecutionResult(new CellValue(
                    CellObjectType.Error,
                    "No AI service configured",
                    "Error: No default AI service connection. Please configure an AI service in Settings."));
            }

            AIResponse response;

            switch (functionName.ToUpperInvariant())
            {
                case "IMAGE_TO_CAPTION":
                    response = await ExecuteImageToCaptionAsync(client, arguments);
                    break;

                case "TEXT_TO_IMAGE":
                    response = await ExecuteTextToImageAsync(client, arguments);
                    break;

                case "TRANSLATE":
                    response = await ExecuteTranslateAsync(client, arguments);
                    break;

                case "SUMMARIZE":
                    response = await ExecuteSummarizeAsync(client, arguments);
                    break;

                case "CHAT":
                    response = await ExecuteChatAsync(client, arguments);
                    break;

                case "CODE_REVIEW":
                    response = await ExecuteCodeReviewAsync(client, arguments);
                    break;

                case "JSON_QUERY":
                    response = await ExecuteJsonQueryAsync(client, arguments);
                    break;

                case "AI_EXTRACT":
                    response = await ExecuteAIExtractAsync(client, arguments);
                    break;

                case "SENTIMENT":
                    response = await ExecuteSentimentAsync(client, arguments);
                    break;

                default:
                    return new FunctionExecutionResult(new CellValue(
                        CellObjectType.Error,
                        $"AI function '{functionName}' not implemented",
                        $"Error: AI function '{functionName}' is not yet implemented."));
            }

            if (response.Success)
            {
                // Determine appropriate cell type based on function
                var cellType = functionName.ToUpperInvariant() switch
                {
                    "TEXT_TO_IMAGE" => CellObjectType.Image,
                    _ => CellObjectType.Text
                };

                return new FunctionExecutionResult(new CellValue(cellType, response.Result, response.Result));
            }
            else
            {
                return new FunctionExecutionResult(new CellValue(
                    CellObjectType.Error,
                    response.Error ?? "Unknown error",
                    $"AI Error: {response.Error}"));
            }
        }
        catch (Exception ex)
        {
            return new FunctionExecutionResult(new CellValue(
                CellObjectType.Error,
                ex.Message,
                $"Error executing AI function: {ex.Message}"));
        }
    }

    private async Task<AIResponse> ExecuteImageToCaptionAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count == 0)
            return AIResponse.FromError("No image provided", TimeSpan.Zero);

        var imagePath = arguments[0].DisplayValue;
        if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
            return AIResponse.FromError("Image file not found", TimeSpan.Zero);

        var prompt = arguments.Count > 1 ? arguments[1].DisplayValue : "Describe this image in detail.";
        return await client.GenerateCaptionAsync(imagePath, prompt);
    }

    private async Task<AIResponse> ExecuteTextToImageAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count == 0)
            return AIResponse.FromError("No prompt provided", TimeSpan.Zero);

        var prompt = arguments[0].DisplayValue;
        return await client.GenerateImageAsync(prompt);
    }

    private async Task<AIResponse> ExecuteTranslateAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count < 2)
            return AIResponse.FromError("Text and target language required", TimeSpan.Zero);

        var text = arguments[0].DisplayValue;
        var targetLanguage = arguments[1].DisplayValue;
        return await client.TranslateAsync(text, targetLanguage);
    }

    private async Task<AIResponse> ExecuteSummarizeAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count == 0)
            return AIResponse.FromError("No text provided", TimeSpan.Zero);

        var text = arguments[0].DisplayValue;
        var maxWords = arguments.Count > 1 && int.TryParse(arguments[1].DisplayValue, out var words) ? words : 100;
        return await client.SummarizeAsync(text, maxWords);
    }

    private async Task<AIResponse> ExecuteChatAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count == 0)
            return AIResponse.FromError("No message provided", TimeSpan.Zero);

        var message = arguments[0].DisplayValue;
        var systemPrompt = arguments.Count > 1 ? arguments[1].DisplayValue : null;
        
        var options = new AICompletionOptions
        {
            SystemPrompt = systemPrompt
        };

        return await client.CompleteTextAsync(message, options);
    }

    private async Task<AIResponse> ExecuteCodeReviewAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count == 0)
            return AIResponse.FromError("No code provided", TimeSpan.Zero);

        var code = arguments[0].DisplayValue;
        var systemPrompt = "You are an expert code reviewer. Analyze the following code for bugs, security issues, performance problems, and suggest improvements. Be specific and actionable.";
        
        var options = new AICompletionOptions { SystemPrompt = systemPrompt };
        return await client.CompleteTextAsync($"Review this code:\n\n```\n{code}\n```", options);
    }

    private async Task<AIResponse> ExecuteJsonQueryAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count < 2)
            return AIResponse.FromError("JSON data and query required", TimeSpan.Zero);

        var json = arguments[0].DisplayValue;
        var query = arguments[1].DisplayValue;
        
        var systemPrompt = "You are a JSON query assistant. Return only the requested data from the JSON, formatted cleanly. No explanations.";
        var options = new AICompletionOptions { SystemPrompt = systemPrompt };
        
        return await client.CompleteTextAsync($"JSON Data:\n{json}\n\nQuery: {query}", options);
    }

    private async Task<AIResponse> ExecuteAIExtractAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count < 2)
            return AIResponse.FromError("Text and extraction type required", TimeSpan.Zero);

        var text = arguments[0].DisplayValue;
        var extractType = arguments[1].DisplayValue;
        
        var systemPrompt = $"Extract all {extractType} from the following text. Return only the extracted values, one per line.";
        var options = new AICompletionOptions { SystemPrompt = systemPrompt };
        
        return await client.CompleteTextAsync(text, options);
    }

    private async Task<AIResponse> ExecuteSentimentAsync(IAIServiceClient client, IReadOnlyList<CellViewModel> arguments)
    {
        if (arguments.Count == 0)
            return AIResponse.FromError("No text provided", TimeSpan.Zero);

        var text = arguments[0].DisplayValue;
        var systemPrompt = "Analyze the sentiment of the following text. Respond with only one word: Positive, Negative, or Neutral, followed by a confidence score (0-100%).";
        
        var options = new AICompletionOptions { SystemPrompt = systemPrompt };
        return await client.CompleteTextAsync(text, options);
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
