using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using AiCalc.Models;

namespace AiCalc.Services.AI;

/// <summary>
/// Ollama local AI service client implementation
/// </summary>
public class OllamaClient : IAIServiceClient
{
    private readonly WorkspaceConnection _connection;
    private readonly HttpClient _httpClient;
    
    public OllamaClient(WorkspaceConnection connection)
    {
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri(connection.Endpoint), // e.g., http://localhost:11434
            Timeout = TimeSpan.FromSeconds(connection.TimeoutSeconds)
        };
    }
    
    public async Task<bool> TestConnectionAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            // Test by trying to list models
            var response = await _httpClient.GetAsync("/api/tags", cancellationToken);
            _connection.LastTested = DateTime.Now;
            _connection.LastTestError = response.IsSuccessStatusCode ? null : $"Status: {response.StatusCode}";
            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            _connection.LastTested = DateTime.Now;
            _connection.LastTestError = ex.Message;
            return false;
        }
    }
    
    public async Task<AIResponse> CompleteTextAsync(string prompt, AICompletionOptions? options = null, CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        options ??= new AICompletionOptions();
        
        try
        {
            var request = new
            {
                model = _connection.Model,
                prompt,
                stream = false,
                options = new
                {
                    temperature = options.Temperature,
                    num_predict = options.MaxTokens
                }
            };
            
            var json = JsonSerializer.Serialize(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            
            var response = await _httpClient.PostAsync("/api/generate", content, cancellationToken);
            
            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync(cancellationToken);
                return AIResponse.FromError($"API Error: {response.StatusCode} - {errorContent}", sw.Elapsed);
            }
            
            var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
            var result = JsonDocument.Parse(responseJson);
            
            _connection.TotalRequests++;
            _connection.LastUsed = DateTime.Now;
            
            var resultText = result.RootElement.GetProperty("response").GetString() ?? string.Empty;
            
            // Ollama doesn't always report token count, estimate it
            var tokensUsed = resultText.Split(' ').Length;
            
            return AIResponse.FromSuccess(resultText, tokensUsed, sw.Elapsed);
        }
        catch (Exception ex)
        {
            return AIResponse.FromError($"Exception: {ex.Message}", sw.Elapsed);
        }
    }
    
    public async Task<AIResponse> GenerateCaptionAsync(string imagePath, int maxWords = 50, CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        
        try
        {
            // Ollama vision model (e.g., llava, bakllava)
            var imageBytes = await System.IO.File.ReadAllBytesAsync(imagePath, cancellationToken);
            var base64Image = Convert.ToBase64String(imageBytes);
            
            var request = new
            {
                model = _connection.VisionModel ?? "llava",
                prompt = $"Describe this image in {maxWords} words or less.",
                images = new[] { base64Image },
                stream = false
            };
            
            var json = JsonSerializer.Serialize(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            
            var response = await _httpClient.PostAsync("/api/generate", content, cancellationToken);
            
            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync(cancellationToken);
                return AIResponse.FromError($"API Error: {response.StatusCode} - {errorContent}", sw.Elapsed);
            }
            
            var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
            var result = JsonDocument.Parse(responseJson);
            
            _connection.TotalRequests++;
            _connection.LastUsed = DateTime.Now;
            
            var caption = result.RootElement.GetProperty("response").GetString() ?? string.Empty;
            var tokensUsed = caption.Split(' ').Length;
            
            return AIResponse.FromSuccess(caption, tokensUsed, sw.Elapsed);
        }
        catch (Exception ex)
        {
            return AIResponse.FromError($"Exception: {ex.Message}", sw.Elapsed);
        }
    }
    
    public Task<AIResponse> GenerateImageAsync(string prompt, string size = "1024x1024", CancellationToken cancellationToken = default)
    {
        // Ollama doesn't support image generation natively
        return Task.FromResult(AIResponse.FromError("Image generation not supported by Ollama", TimeSpan.Zero));
    }
    
    public async Task<AIResponse> TranslateAsync(string text, string targetLanguage, CancellationToken cancellationToken = default)
    {
        var prompt = $"Translate the following text to {targetLanguage}. Only provide the translation, no explanations:\n\n{text}";
        return await CompleteTextAsync(prompt, new AICompletionOptions { Temperature = 0.3 }, cancellationToken);
    }
    
    public async Task<AIResponse> SummarizeAsync(string text, int maxWords = 100, CancellationToken cancellationToken = default)
    {
        var prompt = $"Summarize the following text in {maxWords} words or less:\n\n{text}";
        return await CompleteTextAsync(prompt, new AICompletionOptions { Temperature = 0.5 }, cancellationToken);
    }
    
    public async IAsyncEnumerable<string> StreamCompletionAsync(
        string prompt, 
        AICompletionOptions? options = null, 
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        options ??= new AICompletionOptions();
        
        var request = new
        {
            model = _connection.Model,
            prompt,
            stream = true,
            options = new
            {
                temperature = options.Temperature,
                num_predict = options.MaxTokens
            }
        };
        
        var json = JsonSerializer.Serialize(request);
        var content = new StringContent(json, Encoding.UTF8, "application/json");
        
        var requestMessage = new HttpRequestMessage(HttpMethod.Post, "/api/generate")
        {
            Content = content
        };
        
        using var response = await _httpClient.SendAsync(requestMessage, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        
        response.EnsureSuccessStatusCode();
        
        using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var reader = new System.IO.StreamReader(stream);
        
        while (!reader.EndOfStream && !cancellationToken.IsCancellationRequested)
        {
            var line = await reader.ReadLineAsync(cancellationToken);
            if (string.IsNullOrWhiteSpace(line))
                continue;
            
            var jsonDoc = JsonDocument.Parse(line);
            
            if (jsonDoc.RootElement.TryGetProperty("response", out var responseProperty))
            {
                yield return responseProperty.GetString() ?? string.Empty;
            }
            
            if (jsonDoc.RootElement.TryGetProperty("done", out var doneProperty) && 
                doneProperty.GetBoolean())
            {
                break;
            }
        }
    }
}
