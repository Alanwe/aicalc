using System;
using System.Collections.Generic;
using System.Globalization;
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

    public bool Unregister(string name)
    {
        return _functions.Remove(name);
    }
    
    /// <summary>
    /// Get functions applicable to specific cell types
    /// </summary>
    public IEnumerable<FunctionDescriptor> GetFunctionsForTypes(params CellObjectType[] types)
    {
        if (types == null || types.Length == 0)
            return Functions;
            
        return Functions.Where(f => f.CanAccept(types));
    }
    
    /// <summary>
    /// Get functions by category
    /// </summary>
    public IEnumerable<FunctionDescriptor> GetFunctionsByCategory(FunctionCategory category)
    {
        return Functions.Where(f => f.Category == category);
    }

    private void RegisterBuiltIns()
    {
        RegisterMathFunctions();
        RegisterTextFunctions();
        RegisterDateTimeFunctions();
        RegisterFileFunctions();
        RegisterDirectoryFunctions();
        RegisterTableFunctions();
        RegisterImageFunctions();
        RegisterPdfFunctions();
        RegisterAIFunctions();
    }

    private void RegisterMathFunctions()
    {
        // SUM
        Register(new FunctionDescriptor(
            "SUM",
            "Adds a series of numbers.",
            async ctx =>
            {
                await Task.CompletedTask;
                var sum = ctx.Arguments.Sum(cell => double.TryParse(cell.DisplayValue, out var value) ? value : 0);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, sum.ToString(CultureInfo.InvariantCulture), sum.ToString(CultureInfo.CurrentCulture)));
            },
            FunctionCategory.Math,
            new FunctionParameter("values", "Range of values to add.", CellObjectType.Number)));

        // AVERAGE
        Register(new FunctionDescriptor(
            "AVERAGE",
            "Calculates the average of a series of numbers.",
            async ctx =>
            {
                await Task.CompletedTask;
                var numbers = ctx.Arguments
                    .Where(cell => double.TryParse(cell.DisplayValue, out _))
                    .Select(cell => double.Parse(cell.DisplayValue))
                    .ToList();
                var avg = numbers.Count > 0 ? numbers.Average() : 0;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, avg.ToString(CultureInfo.InvariantCulture), avg.ToString(CultureInfo.CurrentCulture)));
            },
            FunctionCategory.Math,
            new FunctionParameter("values", "Range of values to average.", CellObjectType.Number)));

        // COUNT
        Register(new FunctionDescriptor(
            "COUNT",
            "Counts the number of numeric values.",
            async ctx =>
            {
                await Task.CompletedTask;
                var count = ctx.Arguments.Count(cell => double.TryParse(cell.DisplayValue, out _));
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, count.ToString(), count.ToString()));
            },
            FunctionCategory.Math,
            new FunctionParameter("values", "Range of values to count.", CellObjectType.Number)));

        // MIN
        Register(new FunctionDescriptor(
            "MIN",
            "Returns the minimum value from a series of numbers.",
            async ctx =>
            {
                await Task.CompletedTask;
                var numbers = ctx.Arguments
                    .Where(cell => double.TryParse(cell.DisplayValue, out _))
                    .Select(cell => double.Parse(cell.DisplayValue));
                var min = numbers.Any() ? numbers.Min() : 0;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, min.ToString(CultureInfo.InvariantCulture), min.ToString(CultureInfo.CurrentCulture)));
            },
            FunctionCategory.Math,
            new FunctionParameter("values", "Range of values.", CellObjectType.Number)));

        // MAX
        Register(new FunctionDescriptor(
            "MAX",
            "Returns the maximum value from a series of numbers.",
            async ctx =>
            {
                await Task.CompletedTask;
                var numbers = ctx.Arguments
                    .Where(cell => double.TryParse(cell.DisplayValue, out _))
                    .Select(cell => double.Parse(cell.DisplayValue));
                var max = numbers.Any() ? numbers.Max() : 0;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, max.ToString(CultureInfo.InvariantCulture), max.ToString(CultureInfo.CurrentCulture)));
            },
            FunctionCategory.Math,
            new FunctionParameter("values", "Range of values.", CellObjectType.Number)));

        // ROUND
        Register(new FunctionDescriptor(
            "ROUND",
            "Rounds a number to a specified number of digits.",
            async ctx =>
            {
                await Task.CompletedTask;
                if (ctx.Arguments.Count == 0) return new FunctionExecutionResult(CellValue.Empty);
                
                var value = double.TryParse(ctx.Arguments[0].DisplayValue, out var v) ? v : 0;
                var digits = ctx.Arguments.Count > 1 && int.TryParse(ctx.Arguments[1].DisplayValue, out var d) ? d : 0;
                var rounded = Math.Round(value, digits);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, rounded.ToString(CultureInfo.InvariantCulture), rounded.ToString(CultureInfo.CurrentCulture)));
            },
            FunctionCategory.Math,
            new FunctionParameter("value", "Number to round.", CellObjectType.Number),
            new FunctionParameter("digits", "Number of decimal places.", CellObjectType.Number, isOptional: true)));

        // ABS
        Register(new FunctionDescriptor(
            "ABS",
            "Returns the absolute value of a number.",
            async ctx =>
            {
                await Task.CompletedTask;
                var value = ctx.Arguments.Count > 0 && double.TryParse(ctx.Arguments[0].DisplayValue, out var v) ? v : 0;
                var abs = Math.Abs(value);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, abs.ToString(CultureInfo.InvariantCulture), abs.ToString(CultureInfo.CurrentCulture)));
            },
            FunctionCategory.Math,
            new FunctionParameter("value", "Number.", CellObjectType.Number)));

        // SQRT
        Register(new FunctionDescriptor(
            "SQRT",
            "Returns the square root of a number.",
            async ctx =>
            {
                await Task.CompletedTask;
                var value = ctx.Arguments.Count > 0 && double.TryParse(ctx.Arguments[0].DisplayValue, out var v) ? v : 0;
                var sqrt = Math.Sqrt(value);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, sqrt.ToString(CultureInfo.InvariantCulture), sqrt.ToString(CultureInfo.CurrentCulture)));
            },
            FunctionCategory.Math,
            new FunctionParameter("value", "Number.", CellObjectType.Number)));

        // POWER
        Register(new FunctionDescriptor(
            "POWER",
            "Returns a number raised to a power.",
            async ctx =>
            {
                await Task.CompletedTask;
                var baseValue = ctx.Arguments.Count > 0 && double.TryParse(ctx.Arguments[0].DisplayValue, out var b) ? b : 0;
                var exponent = ctx.Arguments.Count > 1 && double.TryParse(ctx.Arguments[1].DisplayValue, out var e) ? e : 0;
                var result = Math.Pow(baseValue, exponent);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, result.ToString(CultureInfo.InvariantCulture), result.ToString(CultureInfo.CurrentCulture)));
            },
            FunctionCategory.Math,
            new FunctionParameter("base", "Base number.", CellObjectType.Number),
            new FunctionParameter("exponent", "Exponent.", CellObjectType.Number)));
    }

    private void RegisterTextFunctions()
    {
        // CONCAT
        Register(new FunctionDescriptor(
            "CONCAT",
            "Concatenates string values.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = string.Join("", ctx.Arguments.Select(cell => cell.DisplayValue));
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, text, text));
            },
            FunctionCategory.Text,
            new FunctionParameter("values", "Values to concatenate.", CellObjectType.Text)));

        // UPPER
        Register(new FunctionDescriptor(
            "UPPER",
            "Converts text to uppercase.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var upper = text.ToUpper();
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, upper, upper));
            },
            FunctionCategory.Text,
            new FunctionParameter("text", "Text to convert.", CellObjectType.Text)));

        // LOWER
        Register(new FunctionDescriptor(
            "LOWER",
            "Converts text to lowercase.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var lower = text.ToLower();
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, lower, lower));
            },
            FunctionCategory.Text,
            new FunctionParameter("text", "Text to convert.", CellObjectType.Text)));

        // TRIM
        Register(new FunctionDescriptor(
            "TRIM",
            "Removes leading and trailing whitespace.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var trimmed = text.Trim();
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, trimmed, trimmed));
            },
            FunctionCategory.Text,
            new FunctionParameter("text", "Text to trim.", CellObjectType.Text)));

        // LEN
        Register(new FunctionDescriptor(
            "LEN",
            "Returns the length of a text string.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var length = text.Length;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, length.ToString(), length.ToString()));
            },
            FunctionCategory.Text,
            new FunctionParameter("text", "Text to measure.", CellObjectType.Text)));

        // REPLACE
        Register(new FunctionDescriptor(
            "REPLACE",
            "Replaces text within a string.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.Count > 0 ? ctx.Arguments[0].DisplayValue ?? string.Empty : string.Empty;
                var oldText = ctx.Arguments.Count > 1 ? ctx.Arguments[1].DisplayValue ?? string.Empty : string.Empty;
                var newText = ctx.Arguments.Count > 2 ? ctx.Arguments[2].DisplayValue ?? string.Empty : string.Empty;
                var result = text.Replace(oldText, newText);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, result, result));
            },
            FunctionCategory.Text,
            new FunctionParameter("text", "Original text.", CellObjectType.Text),
            new FunctionParameter("old_text", "Text to replace.", CellObjectType.Text),
            new FunctionParameter("new_text", "Replacement text.", CellObjectType.Text)));

        // SPLIT
        Register(new FunctionDescriptor(
            "SPLIT",
            "Splits text into parts using a delimiter.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.Count > 0 ? ctx.Arguments[0].DisplayValue ?? string.Empty : string.Empty;
                var delimiter = ctx.Arguments.Count > 1 ? ctx.Arguments[1].DisplayValue ?? "," : ",";
                var parts = text.Split(new[] { delimiter }, StringSplitOptions.None);
                var result = string.Join(", ", parts);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, result, result));
            },
            FunctionCategory.Text,
            new FunctionParameter("text", "Text to split.", CellObjectType.Text),
            new FunctionParameter("delimiter", "Separator character.", CellObjectType.Text, isOptional: true)));
    }

    private void RegisterDateTimeFunctions()
    {
        // NOW
        Register(new FunctionDescriptor(
            "NOW",
            "Returns the current date and time.",
            async ctx =>
            {
                await Task.CompletedTask;
                var now = DateTime.Now;
                return new FunctionExecutionResult(new CellValue(CellObjectType.DateTime, now.ToString("O"), now.ToString()));
            },
            FunctionCategory.DateTime));

        // TODAY
        Register(new FunctionDescriptor(
            "TODAY",
            "Returns the current date.",
            async ctx =>
            {
                await Task.CompletedTask;
                var today = DateTime.Today;
                return new FunctionExecutionResult(new CellValue(CellObjectType.DateTime, today.ToString("O"), today.ToShortDateString()));
            },
            FunctionCategory.DateTime));

        // DATE
        Register(new FunctionDescriptor(
            "DATE",
            "Creates a date from year, month, and day.",
            async ctx =>
            {
                await Task.CompletedTask;
                var year = ctx.Arguments.Count > 0 && int.TryParse(ctx.Arguments[0].DisplayValue, out var y) ? y : DateTime.Now.Year;
                var month = ctx.Arguments.Count > 1 && int.TryParse(ctx.Arguments[1].DisplayValue, out var m) ? m : 1;
                var day = ctx.Arguments.Count > 2 && int.TryParse(ctx.Arguments[2].DisplayValue, out var d) ? d : 1;
                var date = new DateTime(year, month, day);
                return new FunctionExecutionResult(new CellValue(CellObjectType.DateTime, date.ToString("O"), date.ToShortDateString()));
            },
            FunctionCategory.DateTime,
            new FunctionParameter("year", "Year.", CellObjectType.Number),
            new FunctionParameter("month", "Month.", CellObjectType.Number),
            new FunctionParameter("day", "Day.", CellObjectType.Number)));
    }

    private void RegisterFileFunctions()
    {
        // FILE_SIZE
        Register(new FunctionDescriptor(
            "FILE_SIZE",
            "Returns the size of a file in bytes.",
            async ctx =>
            {
                await Task.CompletedTask;
                var filePath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
                {
                    return new FunctionExecutionResult(new CellValue(CellObjectType.Error, "File not found", "Error: File not found"));
                }
                var size = new System.IO.FileInfo(filePath).Length;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, size.ToString(), $"{size:N0} bytes"));
            },
            FunctionCategory.File,
            new FunctionParameter("file", "File path.", CellObjectType.File, additionalAcceptableTypes: CellObjectType.Text)));

        // FILE_EXTENSION
        Register(new FunctionDescriptor(
            "FILE_EXTENSION",
            "Returns the extension of a file.",
            async ctx =>
            {
                await Task.CompletedTask;
                var filePath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var extension = System.IO.Path.GetExtension(filePath);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, extension, extension));
            },
            FunctionCategory.File,
            new FunctionParameter("file", "File path.", CellObjectType.File, additionalAcceptableTypes: CellObjectType.Text)));

        // FILE_NAME
        Register(new FunctionDescriptor(
            "FILE_NAME",
            "Returns the name of a file without path.",
            async ctx =>
            {
                await Task.CompletedTask;
                var filePath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var fileName = System.IO.Path.GetFileName(filePath);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, fileName, fileName));
            },
            FunctionCategory.File,
            new FunctionParameter("file", "File path.", CellObjectType.File, additionalAcceptableTypes: CellObjectType.Text)));

        // FILE_READ
        Register(new FunctionDescriptor(
            "FILE_READ",
            "Reads the contents of a text file.",
            async ctx =>
            {
                var filePath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
                {
                    return new FunctionExecutionResult(new CellValue(CellObjectType.Error, "File not found", "Error: File not found"));
                }
                var content = await System.IO.File.ReadAllTextAsync(filePath);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, content, content));
            },
            FunctionCategory.File,
            new FunctionParameter("file", "File path.", CellObjectType.File, additionalAcceptableTypes: CellObjectType.Text)));
    }

    private void RegisterDirectoryFunctions()
    {
        // DIR_LIST
        Register(new FunctionDescriptor(
            "DIR_LIST",
            "Lists files in a directory.",
            async ctx =>
            {
                await Task.CompletedTask;
                var dirPath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                if (string.IsNullOrWhiteSpace(dirPath) || !System.IO.Directory.Exists(dirPath))
                {
                    return new FunctionExecutionResult(new CellValue(CellObjectType.Error, "Directory not found", "Error: Directory not found"));
                }
                var files = System.IO.Directory.GetFiles(dirPath);
                var result = string.Join(", ", files.Select(System.IO.Path.GetFileName));
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, result, result));
            },
            FunctionCategory.Directory,
            new FunctionParameter("directory", "Directory path.", CellObjectType.Directory, additionalAcceptableTypes: CellObjectType.Text)));

        // DIR_SIZE
        Register(new FunctionDescriptor(
            "DIR_SIZE",
            "Calculates the total size of files in a directory.",
            async ctx =>
            {
                await Task.CompletedTask;
                var dirPath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                if (string.IsNullOrWhiteSpace(dirPath) || !System.IO.Directory.Exists(dirPath))
                {
                    return new FunctionExecutionResult(new CellValue(CellObjectType.Error, "Directory not found", "Error: Directory not found"));
                }
                var size = System.IO.Directory.GetFiles(dirPath, "*", System.IO.SearchOption.AllDirectories)
                    .Sum(f => new System.IO.FileInfo(f).Length);
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, size.ToString(), $"{size:N0} bytes"));
            },
            FunctionCategory.Directory,
            new FunctionParameter("directory", "Directory path.", CellObjectType.Directory, additionalAcceptableTypes: CellObjectType.Text)));

        // DIRECTORY_TO_TABLE
        Register(new FunctionDescriptor(
            "DIRECTORY_TO_TABLE",
            "Expands a directory listing into a table value.",
            async ctx =>
            {
                await Task.CompletedTask;
                var path = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var json = $"[{{\"name\":\"sample.txt\",\"size\":1024}},{{\"name\":\"image.png\",\"size\":2048}}]";
                return new FunctionExecutionResult(new CellValue(CellObjectType.Table, json, $"Directory snapshot: {path}"));
            },
            FunctionCategory.Directory,
            new FunctionParameter("directory", "Directory path", CellObjectType.Directory, additionalAcceptableTypes: CellObjectType.Text)));
    }

    private void RegisterTableFunctions()
    {
        // TABLE_FILTER (placeholder implementation)
        Register(new FunctionDescriptor(
            "TABLE_FILTER",
            "Filters rows in a table based on criteria.",
            async ctx =>
            {
                await Task.CompletedTask;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Table, "[]", "Filtered table"));
            },
            FunctionCategory.Table,
            new FunctionParameter("table", "Table to filter.", CellObjectType.Table),
            new FunctionParameter("criteria", "Filter criteria.", CellObjectType.Text)));

        // TABLE_SORT (placeholder implementation)
        Register(new FunctionDescriptor(
            "TABLE_SORT",
            "Sorts a table by column.",
            async ctx =>
            {
                await Task.CompletedTask;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Table, "[]", "Sorted table"));
            },
            FunctionCategory.Table,
            new FunctionParameter("table", "Table to sort.", CellObjectType.Table),
            new FunctionParameter("column", "Column name.", CellObjectType.Text)));
    }

    private void RegisterImageFunctions()
    {
        // IMAGE_INFO - Get image dimensions and metadata
        Register(new FunctionDescriptor(
            "IMAGE_INFO",
            "Returns information about an image (dimensions, format, size).",
            async ctx =>
            {
                await Task.CompletedTask;
                var imagePath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                if (string.IsNullOrWhiteSpace(imagePath) || !System.IO.File.Exists(imagePath))
                {
                    return new FunctionExecutionResult(new CellValue(CellObjectType.Error, "Image not found", "Error: Image not found"));
                }
                var info = new System.IO.FileInfo(imagePath);
                var result = $"{info.Name} - {info.Length:N0} bytes";
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, result, result));
            },
            FunctionCategory.Image,
            new FunctionParameter("image", "Image file.", CellObjectType.Image)));
    }

    private void RegisterPdfFunctions()
    {
        // PDF_TO_TEXT (placeholder)
        Register(new FunctionDescriptor(
            "PDF_TO_TEXT",
            "Extracts text from a PDF file.",
            async ctx =>
            {
                await Task.CompletedTask;
                var pdfPath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, "[PDF Text Content]", $"Extracted text from {System.IO.Path.GetFileName(pdfPath)}"));
            },
            FunctionCategory.Pdf,
            new FunctionParameter("pdf", "PDF file.", CellObjectType.Pdf)));

        // PDF_PAGE_COUNT (placeholder)
        Register(new FunctionDescriptor(
            "PDF_PAGE_COUNT",
            "Returns the number of pages in a PDF.",
            async ctx =>
            {
                await Task.CompletedTask;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Number, "1", "1 page"));
            },
            FunctionCategory.Pdf,
            new FunctionParameter("pdf", "PDF file.", CellObjectType.Pdf)));
    }

    private void RegisterAIFunctions()
    {
        // IMAGE_TO_CAPTION - Generate image captions using AI vision models
        Register(new FunctionDescriptor(
            "IMAGE_TO_CAPTION",
            "Generates a descriptive caption for an image using AI vision models (GPT-4-Vision, LLaVA).",
            async ctx =>
            {
                // Implementation will be handled by FunctionRunner with AIServiceRegistry
                await Task.CompletedTask;
                var imagePath = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, "[AI Processing Required]", $"Processing image: {System.IO.Path.GetFileName(imagePath)}"));
            },
            FunctionCategory.AI,
            new FunctionParameter("image", "Image to caption.", CellObjectType.Image)));

        // TEXT_TO_IMAGE - Generate images from text prompts
        Register(new FunctionDescriptor(
            "TEXT_TO_IMAGE",
            "Generates an image from a text prompt using AI (DALL-E 3).",
            async ctx =>
            {
                await Task.CompletedTask;
                var prompt = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Image, "[AI Processing Required]", $"Generating image from: {prompt}"));
            },
            FunctionCategory.AI,
            new FunctionParameter("prompt", "Text description of the image to generate.", CellObjectType.Text)));

        // TRANSLATE - Translate text between languages
        Register(new FunctionDescriptor(
            "TRANSLATE",
            "Translates text from one language to another using AI.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.Count > 0 ? ctx.Arguments[0].DisplayValue ?? string.Empty : string.Empty;
                var targetLang = ctx.Arguments.Count > 1 ? ctx.Arguments[1].DisplayValue ?? "English" : "English";
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, "[AI Processing Required]", $"Translating to {targetLang}"));
            },
            FunctionCategory.AI,
            new FunctionParameter("text", "Text to translate.", CellObjectType.Text),
            new FunctionParameter("target_language", "Target language (e.g., Spanish, French, Japanese).", CellObjectType.Text)));

        // SUMMARIZE - Summarize long text
        Register(new FunctionDescriptor(
            "SUMMARIZE",
            "Creates a concise summary of a longer text using AI.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                var maxLength = ctx.Arguments.Count > 1 && int.TryParse(ctx.Arguments[1].DisplayValue, out var len) ? len : 100;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, "[AI Processing Required]", $"Summarizing text ({text.Length} chars -> ~{maxLength} words)"));
            },
            FunctionCategory.AI,
            new FunctionParameter("text", "Text to summarize.", CellObjectType.Text),
            new FunctionParameter("max_words", "Maximum summary length in words.", CellObjectType.Number, isOptional: true)));

        // CHAT - Interactive conversation with AI
        Register(new FunctionDescriptor(
            "CHAT",
            "Sends a message to an AI assistant and returns the response.",
            async ctx =>
            {
                await Task.CompletedTask;
                var message = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, "[AI Processing Required]", $"Asking AI: {message}"));
            },
            FunctionCategory.AI,
            new FunctionParameter("message", "Your message to the AI.", CellObjectType.Text),
            new FunctionParameter("system_prompt", "Optional system instructions.", CellObjectType.Text, isOptional: true)));

        // CODE_REVIEW - Review code and provide suggestions
        Register(new FunctionDescriptor(
            "CODE_REVIEW",
            "Reviews code and provides suggestions for improvements, bugs, and best practices.",
            async ctx =>
            {
                await Task.CompletedTask;
                var code = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, "[AI Processing Required]", "Reviewing code..."));
            },
            FunctionCategory.AI,
            new FunctionParameter("code", "Code to review.", CellObjectType.CodePython, false, CellObjectType.CodeCSharp, CellObjectType.CodeJavaScript, CellObjectType.CodeTypeScript, CellObjectType.CodeHtml, CellObjectType.CodeCss)));

        // JSON_QUERY - Query JSON data using natural language
        Register(new FunctionDescriptor(
            "JSON_QUERY",
            "Queries JSON data using natural language and returns the result.",
            async ctx =>
            {
                await Task.CompletedTask;
                var json = ctx.Arguments.Count > 0 ? ctx.Arguments[0].DisplayValue ?? string.Empty : string.Empty;
                var query = ctx.Arguments.Count > 1 ? ctx.Arguments[1].DisplayValue ?? string.Empty : string.Empty;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Json, "[AI Processing Required]", $"Query: {query}"));
            },
            FunctionCategory.AI,
            new FunctionParameter("json", "JSON data to query.", CellObjectType.Json),
            new FunctionParameter("query", "Natural language query.", CellObjectType.Text)));

        // AI_EXTRACT - Extract specific information from text
        Register(new FunctionDescriptor(
            "AI_EXTRACT",
            "Extracts specific information from text using AI (e.g., emails, dates, names).",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.Count > 0 ? ctx.Arguments[0].DisplayValue ?? string.Empty : string.Empty;
                var extractType = ctx.Arguments.Count > 1 ? ctx.Arguments[1].DisplayValue ?? "entities" : "entities";
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, "[AI Processing Required]", $"Extracting {extractType}"));
            },
            FunctionCategory.AI,
            new FunctionParameter("text", "Text to analyze.", CellObjectType.Text),
            new FunctionParameter("extract_type", "What to extract (emails, dates, names, entities).", CellObjectType.Text)));

        // SENTIMENT - Analyze sentiment of text
        Register(new FunctionDescriptor(
            "SENTIMENT",
            "Analyzes the sentiment (positive/negative/neutral) of text using AI.",
            async ctx =>
            {
                await Task.CompletedTask;
                var text = ctx.Arguments.FirstOrDefault()?.DisplayValue ?? string.Empty;
                return new FunctionExecutionResult(new CellValue(CellObjectType.Text, "[AI Processing Required]", "Analyzing sentiment..."));
            },
            FunctionCategory.AI,
            new FunctionParameter("text", "Text to analyze.", CellObjectType.Text)));
    }
}
