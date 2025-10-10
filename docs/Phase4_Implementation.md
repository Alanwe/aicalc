# Phase 4: AI Functions & Service Integration

**Date:** October 2025  
**Status:** 🔄 IN PROGRESS  
**Phase:** Phase 4 - AI Integration

---

## Overview

Phase 4 implements the core AI functionality that makes AiCalc unique - integrating AI services with the spreadsheet through a type-safe, class-based cell system. This phase builds on Phase 3's evaluation engine to enable AI-powered transformations on rich cell types.

**Goals:**
1. **Task 1**: Object-Oriented Cell Type System (CellObjects)
2. **Task 2**: Type-Specific Function Registry
3. **Task 11**: AI Function Configuration System
4. **Task 12**: AI Function Execution & Preview
5. **Task 13**: Tie AI Functions to Classes & Preview

**Estimated Time:** 3-4 weeks  
**Priority:** P0 (Critical Path)

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                     Cell Value System                        │
│                                                              │
│  ┌────────────────┐        ┌──────────────────────────┐    │
│  │   CellValue    │        │     ICellObject          │    │
│  │  (existing)    │───────>│   (interface)            │    │
│  └────────────────┘        └──────────────────────────┘    │
│                                       │                      │
│              ┌───────────────────────┼────────────────┐    │
│              │                        │                │     │
│         ┌────▼────┐            ┌────▼────┐      ┌────▼────┐│
│         │ Number  │            │  Text   │  ... │  Image  ││
│         │  Cell   │            │  Cell   │      │  Cell   ││
│         └─────────┘            └─────────┘      └─────────┘│
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                  AI Service Integration                       │
│                                                              │
│  ┌────────────────────┐      ┌───────────────────────────┐ │
│  │ WorkspaceConnection│      │    AI Service Clients     │ │
│  │   (enhanced)       │      │                           │ │
│  │ - Encrypted Keys   │      │  ┌─────────────────────┐ │ │
│  │ - Model Config     │──────┼─>│ AzureOpenAIClient   │ │ │
│  │ - Timeout          │      │  └─────────────────────┘ │ │
│  │ - Token Tracking   │      │  ┌─────────────────────┐ │ │
│  └────────────────────┘      │  │   OllamaClient      │ │ │
│                               │  └─────────────────────┘ │ │
│                               └───────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                Function-to-Type Binding                       │
│                                                              │
│  ┌────────────────────────────────────────────────────────┐ │
│  │          FunctionDescriptor (enhanced)                 │ │
│  │  - Name, Description, Parameters                       │ │
│  │  - Category: Built-in | AI | Contrib                   │ │
│  │  - AcceptedInputTypes: CellObjectType[]                │ │
│  │  - RequiredService: WorkspaceConnection?               │ │
│  └────────────────────────────────────────────────────────┘ │
│                                                              │
│  Example:                                                    │
│  IMAGE_TO_CAPTION accepts: [Image, File]                    │
│  TEXT_TO_IMAGE produces: Image                              │
│  SUM accepts: [Number, Table]                               │
└─────────────────────────────────────────────────────────────┘
```

---

## Task 1: Object-Oriented Cell Type System

### Goal
Implement robust class-based cell type system where each cell type has specific operations and validation.

### Design

#### ICellObject Interface
```csharp
public interface ICellObject
{
    CellObjectType Type { get; }
    object Value { get; }
    bool IsEmpty { get; }
    
    // Validation
    bool IsValid();
    string? GetValidationError();
    
    // Serialization
    string Serialize();
    void Deserialize(string data);
    
    // Display
    string GetDisplayText();
    string GetFormattedValue();
}
```

#### Base Implementation Classes

**1. EmptyCell**
```csharp
public class EmptyCell : ICellObject
{
    public CellObjectType Type => CellObjectType.Empty;
    public object Value => null;
    public bool IsEmpty => true;
}
```

**2. NumberCell**
```csharp
public class NumberCell : ICellObject
{
    public double NumericValue { get; set; }
    public string? FormatString { get; set; } // "0.00", "$#,##0", etc.
    
    // Operations
    public NumberCell Add(NumberCell other) => new NumberCell(NumericValue + other.NumericValue);
    public NumberCell Multiply(NumberCell other) => new NumberCell(NumericValue * other.NumericValue);
    public bool IsInteger() => Math.Abs(NumericValue % 1) < double.Epsilon;
}
```

**3. TextCell**
```csharp
public class TextCell : ICellObject
{
    public string Text { get; set; }
    public int Length => Text?.Length ?? 0;
    
    // Operations
    public TextCell Concat(TextCell other) => new TextCell(Text + other.Text);
    public TextCell ToUpper() => new TextCell(Text?.ToUpper());
    public string[] Split(string delimiter) => Text?.Split(delimiter) ?? Array.Empty<string>();
}
```

**4. ImageCell**
```csharp
public class ImageCell : ICellObject
{
    public string FilePath { get; set; } // Path or URL
    public byte[]? ImageData { get; set; } // Optional cached data
    public int? Width { get; set; }
    public int? Height { get; set; }
    public string? MimeType { get; set; } // image/png, image/jpeg
    
    // Operations
    public bool LoadImage();
    public void SaveImage(string path);
    public ImageCell Resize(int width, int height);
}
```

**5. DirectoryCell**
```csharp
public class DirectoryCell : ICellObject
{
    public string Path { get; set; }
    public List<FileInfo>? Files { get; set; }
    public List<DirectoryInfo>? Subdirectories { get; set; }
    
    // Operations
    public TableCell ToTable(); // Convert to table of files
    public long GetTotalSize();
    public int GetFileCount();
    public DirectoryCell Filter(string pattern);
}
```

**6. FileCell**
```csharp
public class FileCell : ICellObject
{
    public string FilePath { get; set; }
    public long? Size { get; set; }
    public string? Extension { get; set; }
    public DateTime? LastModified { get; set; }
    
    // Operations
    public string ReadAllText();
    public byte[] ReadAllBytes();
    public bool Exists();
}
```

**7. TableCell**
```csharp
public class TableCell : ICellObject
{
    public string[] ColumnNames { get; set; }
    public List<ICellObject[]> Rows { get; set; }
    
    // Operations
    public TableCell Filter(Func<ICellObject[], bool> predicate);
    public TableCell Sort(string columnName);
    public TableCell Join(TableCell other, string onColumn);
    public ICellObject Aggregate(string columnName, string function); // SUM, AVG, COUNT
    public int RowCount => Rows?.Count ?? 0;
    public int ColumnCount => ColumnNames?.Length ?? 0;
}
```

**8. VideoCell**
```csharp
public class VideoCell : ICellObject
{
    public string FilePath { get; set; }
    public TimeSpan? Duration { get; set; }
    public int? Width { get; set; }
    public int? Height { get; set; }
    public string? Codec { get; set; }
    
    // Operations
    public ImageCell ExtractFrame(TimeSpan timestamp);
    public VideoCell Trim(TimeSpan start, TimeSpan end);
}
```

**9. PdfCell & PdfPageCell**
```csharp
public class PdfCell : ICellObject
{
    public string FilePath { get; set; }
    public int? PageCount { get; set; }
    
    // Operations
    public List<PdfPageCell> ExtractPages();
    public TextCell ExtractAllText();
    public TableCell ExtractTables();
}

public class PdfPageCell : ICellObject
{
    public string SourcePdf { get; set; }
    public int PageNumber { get; set; }
    
    // Operations
    public TextCell ExtractText();
    public ImageCell RenderToImage();
    public List<ImageCell> ExtractImages();
}
```

**10. MarkdownCell**
```csharp
public class MarkdownCell : ICellObject
{
    public string Markdown { get; set; }
    
    // Operations
    public string RenderToHtml();
    public TextCell GetPlainText();
    public List<string> GetHeadings();
}
```

**11. JsonCell & XmlCell**
```csharp
public class JsonCell : ICellObject
{
    public string JsonText { get; set; }
    
    // Operations
    public bool IsValid();
    public string Format(); // Pretty print
    public ICellObject Query(string jsonPath);
    public TableCell ToTable(); // If array of objects
}

public class XmlCell : ICellObject
{
    public string XmlText { get; set; }
    
    // Operations
    public bool IsValid();
    public string Format();
    public ICellObject Query(string xpath);
}
```

**12. CodeCell**
```csharp
public class CodeCell : ICellObject
{
    public string Code { get; set; }
    public string Language { get; set; } // "python", "javascript", "csharp"
    
    // Operations
    public ICellObject Execute();
    public bool HasSyntaxErrors();
    public string GetHighlightedHtml();
}
```

### CellObjectFactory

```csharp
public static class CellObjectFactory
{
    public static ICellObject Create(CellValue cellValue)
    {
        if (cellValue == null || cellValue.IsEmpty)
            return new EmptyCell();
            
        switch (cellValue.Type)
        {
            case CellObjectType.Number:
                return new NumberCell { NumericValue = Convert.ToDouble(cellValue.Value) };
                
            case CellObjectType.String:
            case CellObjectType.Text:
                return new TextCell { Text = cellValue.Value?.ToString() ?? "" };
                
            case CellObjectType.Image:
                return new ImageCell { FilePath = cellValue.Value?.ToString() ?? "" };
                
            case CellObjectType.Directory:
                return new DirectoryCell { Path = cellValue.Value?.ToString() ?? "" };
                
            // ... other types
                
            default:
                return new TextCell { Text = cellValue.Value?.ToString() ?? "" };
        }
    }
    
    public static ICellObject AutoDetect(object value)
    {
        if (value == null) return new EmptyCell();
        
        // Try parse as number
        if (double.TryParse(value.ToString(), out var number))
            return new NumberCell { NumericValue = number };
            
        var str = value.ToString();
        
        // Check if file path
        if (File.Exists(str))
            return new FileCell { FilePath = str };
            
        // Check if directory
        if (Directory.Exists(str))
            return new DirectoryCell { Path = str };
            
        // Check if image URL
        if (str.StartsWith("http") && IsImageUrl(str))
            return new ImageCell { FilePath = str };
            
        // Check if JSON
        if (str.TrimStart().StartsWith("{") || str.TrimStart().StartsWith("["))
            return new JsonCell { JsonText = str };
            
        // Default to text
        return new TextCell { Text = str };
    }
}
```

---

## Task 2: Type-Specific Function Registry

### Enhanced FunctionDescriptor

```csharp
public class FunctionDescriptor
{
    public string Name { get; set; }
    public string Description { get; set; }
    public List<FunctionParameter> Parameters { get; set; }
    
    // NEW: Type binding
    public CellObjectType[] AcceptedInputTypes { get; set; }
    public CellObjectType OutputType { get; set; }
    
    // NEW: Categorization
    public FunctionCategory Category { get; set; } // Built-in, AI, Contrib
    
    // NEW: AI Service requirement
    public string? RequiredServiceType { get; set; } // "AzureOpenAI", "Ollama", etc.
    
    public Func<FunctionEvaluationContext, Task<CellValue>> Implementation { get; set; }
}

public enum FunctionCategory
{
    BuiltIn,    // SUM, AVERAGE, CONCAT
    AI,         // IMAGE_TO_CAPTION, TEXT_TO_IMAGE
    Contrib     // Community-contributed
}

public class FunctionParameter
{
    public string Name { get; set; }
    public string Description { get; set; }
    public CellObjectType[] AcceptedTypes { get; set; }
    public bool IsRequired { get; set; }
    public bool IsRange { get; set; } // Accept A1:A10
}
```

### Function Registry Enhancements

```csharp
public class FunctionRegistry
{
    // NEW: Filter by cell type
    public List<FunctionDescriptor> GetApplicableFunctions(CellObjectType cellType)
    {
        return _functions.Values
            .Where(f => f.AcceptedInputTypes == null || f.AcceptedInputTypes.Contains(cellType))
            .ToList();
    }
    
    // NEW: Filter by category
    public List<FunctionDescriptor> GetFunctionsByCategory(FunctionCategory category)
    {
        return _functions.Values
            .Where(f => f.Category == category)
            .ToList();
    }
    
    // NEW: Get AI functions requiring service
    public List<FunctionDescriptor> GetAIFunctions(string serviceType)
    {
        return _functions.Values
            .Where(f => f.Category == FunctionCategory.AI && 
                       (f.RequiredServiceType == null || f.RequiredServiceType == serviceType))
            .ToList();
    }
}
```

### Example Function Definitions

```csharp
// IMAGE_TO_CAPTION - accepts Image or File, produces Text
new FunctionDescriptor
{
    Name = "IMAGE_TO_CAPTION",
    Description = "Generate a caption for an image using AI vision",
    Category = FunctionCategory.AI,
    RequiredServiceType = "AzureOpenAI",
    AcceptedInputTypes = new[] { CellObjectType.Image, CellObjectType.File },
    OutputType = CellObjectType.Text,
    Parameters = new List<FunctionParameter>
    {
        new() { Name = "image", AcceptedTypes = new[] { CellObjectType.Image }, IsRequired = true },
        new() { Name = "max_words", AcceptedTypes = new[] { CellObjectType.Number }, IsRequired = false }
    },
    Implementation = async (ctx) => await AIFunctions.ImageToCaption(ctx)
}

// TEXT_TO_IMAGE - accepts Text, produces Image
new FunctionDescriptor
{
    Name = "TEXT_TO_IMAGE",
    Description = "Generate an image from text prompt using AI",
    Category = FunctionCategory.AI,
    RequiredServiceType = "AzureOpenAI",
    AcceptedInputTypes = new[] { CellObjectType.Text },
    OutputType = CellObjectType.Image,
    Parameters = new List<FunctionParameter>
    {
        new() { Name = "prompt", AcceptedTypes = new[] { CellObjectType.Text }, IsRequired = true },
        new() { Name = "size", AcceptedTypes = new[] { CellObjectType.Text }, IsRequired = false }
    },
    Implementation = async (ctx) => await AIFunctions.TextToImage(ctx)
}

// SUM - accepts Number or Table, produces Number
new FunctionDescriptor
{
    Name = "SUM",
    Description = "Sum numbers or table column",
    Category = FunctionCategory.BuiltIn,
    AcceptedInputTypes = new[] { CellObjectType.Number, CellObjectType.Table },
    OutputType = CellObjectType.Number,
    Parameters = new List<FunctionParameter>
    {
        new() { Name = "values", AcceptedTypes = new[] { CellObjectType.Number }, IsRequired = true, IsRange = true }
    },
    Implementation = async (ctx) => await BuiltInFunctions.Sum(ctx)
}
```

---

## Task 11: AI Function Configuration System

### Enhanced WorkspaceConnection Model

```csharp
public class WorkspaceConnection
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = "Untitled Connection";
    public string ProviderType { get; set; } = "AzureOpenAI"; // "AzureOpenAI", "Ollama", "OpenAI"
    
    // Connection details
    public string Endpoint { get; set; } = "";
    
    // ENHANCED: Encrypted credential storage
    public string EncryptedApiKey { get; set; } = ""; // Encrypted with DPAPI
    
    // ENHANCED: Model configuration
    public string DefaultModel { get; set; } = "gpt-4";
    public string DefaultVisionModel { get; set; } = "gpt-4-vision";
    public string DefaultImageModel { get; set; } = "dall-e-3";
    
    // ENHANCED: Performance settings
    public int TimeoutSeconds { get; set; } = 100;
    public int MaxRetries { get; set; } = 3;
    public double Temperature { get; set; } = 0.7;
    
    // ENHANCED: Token tracking
    public long TotalTokensUsed { get; set; }
    public long TotalRequests { get; set; }
    public DateTime? LastUsed { get; set; }
    
    // ENHANCED: Status
    public bool IsActive { get; set; } = true;
    public DateTime? LastTested { get; set; }
    public string? LastTestError { get; set; }
}
```

### Credential Security Service

```csharp
public static class CredentialService
{
    // Encrypt API key using Windows DPAPI
    public static string Encrypt(string plainText)
    {
        if (string.IsNullOrEmpty(plainText))
            return "";
            
        var plainBytes = Encoding.UTF8.GetBytes(plainText);
        var encryptedBytes = ProtectedData.Protect(
            plainBytes,
            optionalEntropy: null,
            scope: DataProtectionScope.CurrentUser
        );
        return Convert.ToBase64String(encryptedBytes);
    }
    
    // Decrypt API key
    public static string Decrypt(string encryptedText)
    {
        if (string.IsNullOrEmpty(encryptedText))
            return "";
            
        var encryptedBytes = Convert.FromBase64String(encryptedText);
        var plainBytes = ProtectedData.Unprotect(
            encryptedBytes,
            optionalEntropy: null,
            scope: DataProtectionScope.CurrentUser
        );
        return Encoding.UTF8.GetString(plainBytes);
    }
}
```

### Service Registry

```csharp
public class AIServiceRegistry
{
    private readonly Dictionary<string, WorkspaceConnection> _connections = new();
    private readonly Dictionary<string, IAIServiceClient> _clients = new();
    
    public void RegisterConnection(WorkspaceConnection connection)
    {
        _connections[connection.Id] = connection;
        _clients[connection.Id] = CreateClient(connection);
    }
    
    public IAIServiceClient GetClient(string connectionId)
    {
        return _clients.TryGetValue(connectionId, out var client) ? client : null;
    }
    
    public async Task<bool> TestConnectionAsync(string connectionId)
    {
        var connection = _connections[connectionId];
        var client = _clients[connectionId];
        
        try
        {
            await client.TestConnectionAsync();
            connection.LastTested = DateTime.Now;
            connection.LastTestError = null;
            return true;
        }
        catch (Exception ex)
        {
            connection.LastTestError = ex.Message;
            return false;
        }
    }
    
    private IAIServiceClient CreateClient(WorkspaceConnection connection)
    {
        return connection.ProviderType switch
        {
            "AzureOpenAI" => new AzureOpenAIClient(connection),
            "Ollama" => new OllamaClient(connection),
            "OpenAI" => new OpenAIClient(connection),
            _ => throw new NotSupportedException($"Provider {connection.ProviderType} not supported")
        };
    }
}
```

---

## Task 12: AI Function Execution & Preview

### IAIServiceClient Interface

```csharp
public interface IAIServiceClient
{
    Task<bool> TestConnectionAsync();
    Task<string> CompleteTextAsync(string prompt, CancellationToken cancellationToken = default);
    Task<string> GenerateCaptionAsync(string imagePath, int maxWords = 50, CancellationToken cancellationToken = default);
    Task<string> GenerateImageAsync(string prompt, string size = "1024x1024", CancellationToken cancellationToken = default);
    Task<string> TranslateAsync(string text, string targetLanguage, CancellationToken cancellationToken = default);
    
    // Streaming support
    IAsyncEnumerable<string> StreamCompletionAsync(string prompt, CancellationToken cancellationToken = default);
}

public class AIResponse
{
    public bool Success { get; set; }
    public string Result { get; set; }
    public int TokensUsed { get; set; }
    public TimeSpan Duration { get; set; }
    public string? Error { get; set; }
}
```

### Azure OpenAI Client Implementation

```csharp
public class AzureOpenAIClient : IAIServiceClient
{
    private readonly WorkspaceConnection _connection;
    private readonly HttpClient _httpClient;
    
    public AzureOpenAIClient(WorkspaceConnection connection)
    {
        _connection = connection;
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri(connection.Endpoint),
            Timeout = TimeSpan.FromSeconds(connection.TimeoutSeconds)
        };
        
        var apiKey = CredentialService.Decrypt(connection.EncryptedApiKey);
        _httpClient.DefaultRequestHeaders.Add("api-key", apiKey);
    }
    
    public async Task<string> CompleteTextAsync(string prompt, CancellationToken cancellationToken = default)
    {
        var request = new
        {
            messages = new[]
            {
                new { role = "system", content = "You are a helpful assistant." },
                new { role = "user", content = prompt }
            },
            temperature = _connection.Temperature,
            max_tokens = 1000
        };
        
        var json = JsonSerializer.Serialize(request);
        var content = new StringContent(json, Encoding.UTF8, "application/json");
        
        var response = await _httpClient.PostAsync(
            $"/openai/deployments/{_connection.DefaultModel}/chat/completions?api-version=2024-02-01",
            content,
            cancellationToken
        );
        
        response.EnsureSuccessStatusCode();
        
        var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
        var result = JsonDocument.Parse(responseJson);
        
        // Track usage
        var tokensUsed = result.RootElement.GetProperty("usage").GetProperty("total_tokens").GetInt32();
        _connection.TotalTokensUsed += tokensUsed;
        _connection.TotalRequests++;
        _connection.LastUsed = DateTime.Now;
        
        return result.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString();
    }
    
    public async Task<string> GenerateCaptionAsync(string imagePath, int maxWords = 50, CancellationToken cancellationToken = default)
    {
        // Load image and encode to base64
        var imageBytes = await File.ReadAllBytesAsync(imagePath, cancellationToken);
        var base64Image = Convert.ToBase64String(imageBytes);
        var mimeType = GetMimeType(imagePath);
        
        var request = new
        {
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = new object[]
                    {
                        new { type = "text", text = $"Describe this image in {maxWords} words or less." },
                        new { type = "image_url", image_url = new { url = $"data:{mimeType};base64,{base64Image}" } }
                    }
                }
            },
            max_tokens = maxWords * 2
        };
        
        var json = JsonSerializer.Serialize(request);
        var content = new StringContent(json, Encoding.UTF8, "application/json");
        
        var response = await _httpClient.PostAsync(
            $"/openai/deployments/{_connection.DefaultVisionModel}/chat/completions?api-version=2024-02-01",
            content,
            cancellationToken
        );
        
        response.EnsureSuccessStatusCode();
        
        var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
        var result = JsonDocument.Parse(responseJson);
        
        // Track usage
        var tokensUsed = result.RootElement.GetProperty("usage").GetProperty("total_tokens").GetInt32();
        _connection.TotalTokensUsed += tokensUsed;
        _connection.TotalRequests++;
        _connection.LastUsed = DateTime.Now;
        
        return result.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString();
    }
    
    // Other implementations...
}
```

### Ollama Client Implementation

```csharp
public class OllamaClient : IAIServiceClient
{
    private readonly WorkspaceConnection _connection;
    private readonly HttpClient _httpClient;
    
    public OllamaClient(WorkspaceConnection connection)
    {
        _connection = connection;
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri(connection.Endpoint), // e.g., http://localhost:11434
            Timeout = TimeSpan.FromSeconds(connection.TimeoutSeconds)
        };
    }
    
    public async Task<string> CompleteTextAsync(string prompt, CancellationToken cancellationToken = default)
    {
        var request = new
        {
            model = _connection.DefaultModel, // e.g., "llama2", "mistral"
            prompt = prompt,
            stream = false
        };
        
        var json = JsonSerializer.Serialize(request);
        var content = new StringContent(json, Encoding.UTF8, "application/json");
        
        var response = await _httpClient.PostAsync("/api/generate", content, cancellationToken);
        response.EnsureSuccessStatusCode();
        
        var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
        var result = JsonDocument.Parse(responseJson);
        
        _connection.TotalRequests++;
        _connection.LastUsed = DateTime.Now;
        
        return result.RootElement.GetProperty("response").GetString();
    }
    
    public async Task<string> GenerateCaptionAsync(string imagePath, int maxWords = 50, CancellationToken cancellationToken = default)
    {
        // Ollama vision model (e.g., llava)
        var imageBytes = await File.ReadAllBytesAsync(imagePath, cancellationToken);
        var base64Image = Convert.ToBase64String(imageBytes);
        
        var request = new
        {
            model = "llava", // Vision model
            prompt = $"Describe this image in {maxWords} words or less.",
            images = new[] { base64Image },
            stream = false
        };
        
        var json = JsonSerializer.Serialize(request);
        var content = new StringContent(json, Encoding.UTF8, "application/json");
        
        var response = await _httpClient.PostAsync("/api/generate", content, cancellationToken);
        response.EnsureSuccessStatusCode();
        
        var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
        var result = JsonDocument.Parse(responseJson);
        
        _connection.TotalRequests++;
        _connection.LastUsed = DateTime.Now;
        
        return result.RootElement.GetProperty("response").GetString();
    }
    
    // Other implementations...
}
```

### AI Functions Implementation

```csharp
public static class AIFunctions
{
    public static async Task<CellValue> ImageToCaption(FunctionEvaluationContext ctx)
    {
        // Get image cell
        var imageCell = ctx.GetArgument<ImageCell>(0);
        if (imageCell == null)
            return CellValue.Error("IMAGE_TO_CAPTION requires an image");
        
        // Get optional max_words parameter
        var maxWords = ctx.GetArgument<NumberCell>(1)?.NumericValue ?? 50;
        
        // Get AI service
        var service = ctx.Workbook.GetDefaultAIService();
        if (service == null)
            return CellValue.Error("No AI service configured");
        
        var client = ctx.ServiceRegistry.GetClient(service.Id);
        
        try
        {
            var caption = await client.GenerateCaptionAsync(
                imageCell.FilePath,
                (int)maxWords,
                ctx.CancellationToken
            );
            
            return CellValue.FromText(caption);
        }
        catch (Exception ex)
        {
            return CellValue.Error($"AI error: {ex.Message}");
        }
    }
    
    public static async Task<CellValue> TextToImage(FunctionEvaluationContext ctx)
    {
        // Get prompt
        var promptCell = ctx.GetArgument<TextCell>(0);
        if (promptCell == null)
            return CellValue.Error("TEXT_TO_IMAGE requires a text prompt");
        
        // Get optional size parameter
        var size = ctx.GetArgument<TextCell>(1)?.Text ?? "1024x1024";
        
        // Get AI service
        var service = ctx.Workbook.GetDefaultAIService();
        if (service == null)
            return CellValue.Error("No AI service configured");
        
        var client = ctx.ServiceRegistry.GetClient(service.Id);
        
        try
        {
            var imageUrl = await client.GenerateImageAsync(
                promptCell.Text,
                size,
                ctx.CancellationToken
            );
            
            // Download image and save locally
            var imagePath = await DownloadImageAsync(imageUrl, ctx.Workbook.WorkingDirectory);
            
            return CellValue.FromImage(imagePath);
        }
        catch (Exception ex)
        {
            return CellValue.Error($"AI error: {ex.Message}");
        }
    }
    
    public static async Task<CellValue> Translate(FunctionEvaluationContext ctx)
    {
        var textCell = ctx.GetArgument<TextCell>(0);
        var targetLang = ctx.GetArgument<TextCell>(1)?.Text ?? "English";
        
        if (textCell == null)
            return CellValue.Error("TRANSLATE requires text input");
        
        var service = ctx.Workbook.GetDefaultAIService();
        if (service == null)
            return CellValue.Error("No AI service configured");
        
        var client = ctx.ServiceRegistry.GetClient(service.Id);
        
        try
        {
            var translated = await client.TranslateAsync(
                textCell.Text,
                targetLang,
                ctx.CancellationToken
            );
            
            return CellValue.FromText(translated);
        }
        catch (Exception ex)
        {
            return CellValue.Error($"AI error: {ex.Message}");
        }
    }
    
    public static async Task<CellValue> Summarize(FunctionEvaluationContext ctx)
    {
        var textCell = ctx.GetArgument<TextCell>(0);
        var maxWords = ctx.GetArgument<NumberCell>(1)?.NumericValue ?? 100;
        
        if (textCell == null)
            return CellValue.Error("SUMMARIZE requires text input");
        
        var service = ctx.Workbook.GetDefaultAIService();
        if (service == null)
            return CellValue.Error("No AI service configured");
        
        var client = ctx.ServiceRegistry.GetClient(service.Id);
        
        try
        {
            var prompt = $"Summarize the following text in {maxWords} words or less:\\n\\n{textCell.Text}";
            var summary = await client.CompleteTextAsync(prompt, ctx.CancellationToken);
            
            return CellValue.FromText(summary);
        }
        catch (Exception ex)
        {
            return CellValue.Error($"AI error: {ex.Message}");
        }
    }
}
```

---

## UI Enhancements

### ServiceConnectionDialog Enhancements

**Add to SettingsDialog.xaml** (new "AI Services" tab):

```xml
<PivotItem Header="AI Services">
    <ScrollViewer>
        <StackPanel Spacing="12" Padding="12">
            <!-- Connection List -->
            <TextBlock Text="Configured Services" Style="{StaticResource SubtitleTextBlockStyle}"/>
            
            <ListView ItemsSource="{x:Bind ViewModel.Connections}" 
                     SelectedItem="{x:Bind ViewModel.SelectedConnection, Mode=TwoWay}"
                     Height="200">
                <ListView.ItemTemplate>
                    <DataTemplate x:DataType="models:WorkspaceConnection">
                        <Grid ColumnDefinitions="Auto,*,Auto">
                            <FontIcon Grid.Column="0" Glyph="&#xE9CE;" Margin="0,0,12,0"
                                     Foreground="{ThemeResource SystemAccentColor}"/>
                            <StackPanel Grid.Column="1">
                                <TextBlock Text="{x:Bind Name}" FontWeight="SemiBold"/>
                                <TextBlock Text="{x:Bind ProviderType}" 
                                         Foreground="{ThemeResource TextFillColorSecondaryBrush}"
                                         FontSize="12"/>
                            </StackPanel>
                            <ToggleSwitch Grid.Column="2" IsOn="{x:Bind IsActive, Mode=TwoWay}"
                                        OffContent="" OnContent=""/>
                        </Grid>
                    </DataTemplate>
                </ListView.ItemTemplate>
            </ListView>
            
            <StackPanel Orientation="Horizontal" Spacing="8">
                <Button Content="Add Service" Command="{x:Bind ViewModel.AddServiceCommand}"/>
                <Button Content="Edit" Command="{x:Bind ViewModel.EditServiceCommand}"
                       IsEnabled="{x:Bind ViewModel.HasSelectedConnection, Mode=OneWay}"/>
                <Button Content="Test Connection" Command="{x:Bind ViewModel.TestConnectionCommand}"
                       IsEnabled="{x:Bind ViewModel.HasSelectedConnection, Mode=OneWay}"/>
                <Button Content="Remove" Command="{x:Bind ViewModel.RemoveServiceCommand}"
                       IsEnabled="{x:Bind ViewModel.HasSelectedConnection, Mode=OneWay}"
                       Style="{StaticResource AccentButtonStyle}"/>
            </StackPanel>
            
            <!-- Connection Details -->
            <Border Background="{ThemeResource CardBackgroundFillColorDefaultBrush}"
                   BorderBrush="{ThemeResource CardStrokeColorDefaultBrush}"
                   BorderThickness="1" CornerRadius="4" Padding="12"
                   Visibility="{x:Bind ViewModel.HasSelectedConnection, Mode=OneWay}">
                <StackPanel Spacing="12">
                    <TextBlock Text="Connection Details" FontWeight="SemiBold"/>
                    
                    <TextBox Header="Name" 
                            Text="{x:Bind ViewModel.SelectedConnection.Name, Mode=TwoWay}"/>
                    
                    <ComboBox Header="Provider Type" 
                             SelectedItem="{x:Bind ViewModel.SelectedConnection.ProviderType, Mode=TwoWay}"
                             ItemsSource="{x:Bind ViewModel.ProviderTypes}"/>
                    
                    <TextBox Header="Endpoint" 
                            Text="{x:Bind ViewModel.SelectedConnection.Endpoint, Mode=TwoWay}"
                            PlaceholderText="https://your-resource.openai.azure.com"/>
                    
                    <PasswordBox Header="API Key"
                                Password="{x:Bind ViewModel.ApiKeyPlainText, Mode=TwoWay}"/>
                    
                    <TextBox Header="Default Model" 
                            Text="{x:Bind ViewModel.SelectedConnection.DefaultModel, Mode=TwoWay}"
                            PlaceholderText="gpt-4"/>
                    
                    <NumberBox Header="Timeout (seconds)" 
                              Value="{x:Bind ViewModel.SelectedConnection.TimeoutSeconds, Mode=TwoWay}"
                              Minimum="10" Maximum="600" SpinButtonPlacementMode="Inline"/>
                    
                    <NumberBox Header="Temperature" 
                              Value="{x:Bind ViewModel.SelectedConnection.Temperature, Mode=TwoWay}"
                              Minimum="0" Maximum="2" SmallChange="0.1" LargeChange="0.5"
                              SpinButtonPlacementMode="Inline"/>
                    
                    <!-- Usage Statistics -->
                    <StackPanel Spacing="4" Visibility="{x:Bind ViewModel.SelectedConnection.LastUsed, Mode=OneWay, Converter={StaticResource ObjectToBoolConverter}}">
                        <TextBlock Text="Usage Statistics" FontWeight="SemiBold" FontSize="12"/>
                        <TextBlock FontSize="12">
                            <Run Text="Total Requests:"/>
                            <Run Text="{x:Bind ViewModel.SelectedConnection.TotalRequests, Mode=OneWay}"/>
                        </TextBlock>
                        <TextBlock FontSize="12">
                            <Run Text="Total Tokens:"/>
                            <Run Text="{x:Bind ViewModel.SelectedConnection.TotalTokensUsed, Mode=OneWay}"/>
                        </TextBlock>
                        <TextBlock FontSize="12">
                            <Run Text="Last Used:"/>
                            <Run Text="{x:Bind ViewModel.SelectedConnection.LastUsed, Mode=OneWay}"/>
                        </TextBlock>
                    </StackPanel>
                </StackPanel>
            </Border>
        </StackPanel>
    </ScrollViewer>
</PivotItem>
```

---

## Implementation Plan

### Week 1: Foundation (Task 1 & 2)
**Days 1-2**: Implement ICellObject interface and base classes
- EmptyCell, NumberCell, TextCell, ImageCell
- Basic operations and validation

**Days 3-4**: Implement rich cell types
- DirectoryCell, FileCell, TableCell
- PdfCell, MarkdownCell, JsonCell, XmlCell

**Day 5**: Implement CellObjectFactory
- Auto-detection logic
- Type conversion
- Testing

### Week 2: Type System Integration (Task 2 continued)
**Days 1-2**: Enhance FunctionDescriptor
- Add type binding properties
- Implement function filtering

**Days 3-4**: Update existing functions
- Add type metadata to all built-in functions
- Implement parameter validation

**Day 5**: UI updates
- Function panel filtering by cell type
- Category grouping

### Week 3: AI Service Infrastructure (Task 11 & 12)
**Days 1-2**: Enhance WorkspaceConnection
- Credential encryption
- Model configuration
- Token tracking

**Days 3-4**: Implement AI Service Clients
- AzureOpenAIClient
- OllamaClient
- Service registry

**Day 5**: Connection testing and validation

### Week 4: AI Functions & UI (Task 12 & 13)
**Days 1-2**: Implement AI Functions
- IMAGE_TO_CAPTION
- TEXT_TO_IMAGE
- TRANSLATE, SUMMARIZE

**Days 3-4**: ServiceConnectionDialog enhancements
- Add/Edit/Test connections UI
- Usage statistics display

**Day 5**: Testing, documentation, commit

---

## Testing Strategy

### Unit Tests
1. **Cell Object Tests**
   - Creation and validation
   - Type conversions
   - Operations (add, concat, filter, etc.)

2. **Function Registry Tests**
   - Type filtering
   - Category filtering
   - Parameter validation

3. **AI Service Client Tests**
   - Mock HTTP responses
   - Error handling
   - Token tracking

### Integration Tests
1. **End-to-End AI Function Tests**
   - Real API calls (with test keys)
   - Timeout handling
   - Preview mode

2. **Type Safety Tests**
   - Invalid type combinations
   - Auto-conversion scenarios

---

## Success Criteria

✅ **Task 1 Complete:**
- [ ] All 12 cell object classes implemented
- [ ] CellObjectFactory with auto-detection
- [ ] Unit tests passing

✅ **Task 2 Complete:**
- [ ] FunctionDescriptor enhanced with type binding
- [ ] Function registry filtering working
- [ ] UI shows only applicable functions

✅ **Task 11 Complete:**
- [ ] WorkspaceConnection with encryption
- [ ] AI Service registry
- [ ] Connection testing functional

✅ **Task 12 Complete:**
- [ ] Azure OpenAI client working
- [ ] Ollama client working
- [ ] At least 4 AI functions implemented

✅ **Task 13 Complete:**
- [ ] Functions properly bound to cell types
- [ ] Preview mode functional
- [ ] Usage tracking working

---

## Files to Create/Modify

### New Files:
1. `Models/CellObjects/ICellObject.cs`
2. `Models/CellObjects/EmptyCell.cs`
3. `Models/CellObjects/NumberCell.cs`
4. `Models/CellObjects/TextCell.cs`
5. `Models/CellObjects/ImageCell.cs`
6. `Models/CellObjects/DirectoryCell.cs`
7. `Models/CellObjects/FileCell.cs`
8. `Models/CellObjects/TableCell.cs`
9. `Models/CellObjects/VideoCell.cs`
10. `Models/CellObjects/PdfCell.cs`
11. `Models/CellObjects/PdfPageCell.cs`
12. `Models/CellObjects/MarkdownCell.cs`
13. `Models/CellObjects/JsonCell.cs`
14. `Models/CellObjects/XmlCell.cs`
15. `Models/CellObjects/CodeCell.cs`
16. `Models/CellObjects/CellObjectFactory.cs`
17. `Services/AI/IAIServiceClient.cs`
18. `Services/AI/AzureOpenAIClient.cs`
19. `Services/AI/OllamaClient.cs`
20. `Services/AI/AIServiceRegistry.cs`
21. `Services/AI/CredentialService.cs`
22. `Services/AI/AIFunctions.cs`

### Modified Files:
1. `Models/CellObjectType.cs` - Add new enum values
2. `Models/WorkspaceConnection.cs` - Enhance with encryption, tracking
3. `Services/FunctionDescriptor.cs` - Add type binding
4. `Services/FunctionRegistry.cs` - Add filtering methods
5. `SettingsDialog.xaml` - Add AI Services tab
6. `SettingsDialog.xaml.cs` - Add service management logic

---

## Next Steps

1. **Review and Approve Plan**: Confirm approach and priorities
2. **Start Implementation**: Begin with Task 1 (Cell Objects)
3. **Iterative Development**: Build, test, commit incrementally
4. **Documentation**: Update as features complete

**Ready to start Phase 4 implementation!** 🚀
