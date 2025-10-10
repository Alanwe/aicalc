using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace AiCalc.Services.AI;

/// <summary>
/// Interface for AI service clients (Azure OpenAI, Ollama, OpenAI, etc.)
/// </summary>
public interface IAIServiceClient
{
    /// <summary>
    /// Test if the connection is working
    /// </summary>
    Task<bool> TestConnectionAsync(CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Complete text with AI
    /// </summary>
    Task<AIResponse> CompleteTextAsync(string prompt, AICompletionOptions? options = null, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Generate a caption for an image
    /// </summary>
    Task<AIResponse> GenerateCaptionAsync(string imagePath, int maxWords = 50, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Generate an image from text prompt
    /// </summary>
    Task<AIResponse> GenerateImageAsync(string prompt, string size = "1024x1024", CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Translate text to target language
    /// </summary>
    Task<AIResponse> TranslateAsync(string text, string targetLanguage, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Summarize text
    /// </summary>
    Task<AIResponse> SummarizeAsync(string text, int maxWords = 100, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Stream completion responses (for long-running operations)
    /// </summary>
    IAsyncEnumerable<string> StreamCompletionAsync(string prompt, AICompletionOptions? options = null, CancellationToken cancellationToken = default);
}

/// <summary>
/// Options for AI text completion
/// </summary>
public record AICompletionOptions
{
    public double Temperature { get; init; } = 0.7;
    public int MaxTokens { get; init; } = 1000;
    public string? SystemPrompt { get; init; }
    public List<ChatMessage>? ConversationHistory { get; init; }
}

/// <summary>
/// Chat message for conversation history
/// </summary>
public record ChatMessage(string Role, string Content);

/// <summary>
/// Response from AI service
/// </summary>
public record AIResponse
{
    public bool Success { get; init; }
    public string? Result { get; init; }
    public int TokensUsed { get; init; }
    public TimeSpan Duration { get; init; }
    public string? Error { get; init; }
    public Dictionary<string, object>? Metadata { get; init; }
    
    public static AIResponse FromSuccess(string result, int tokensUsed, TimeSpan duration, Dictionary<string, object>? metadata = null)
    {
        return new AIResponse
        {
            Success = true,
            Result = result,
            TokensUsed = tokensUsed,
            Duration = duration,
            Metadata = metadata
        };
    }
    
    public static AIResponse FromError(string error, TimeSpan duration)
    {
        return new AIResponse
        {
            Success = false,
            Error = error,
            Duration = duration
        };
    }
}
