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
/// Azure OpenAI Service client implementation
/// </summary>
public class AzureOpenAIClient : IAIServiceClient
{
    private readonly WorkspaceConnection _connection;
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;
    
    public AzureOpenAIClient(WorkspaceConnection connection)
    {
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _apiKey = CredentialService.Decrypt(connection.ApiKey);
        
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri(connection.Endpoint),
            Timeout = TimeSpan.FromSeconds(connection.TimeoutSeconds)
        };
        
        _httpClient.DefaultRequestHeaders.Add("api-key", _apiKey);
    }
    
    public async Task<bool> TestConnectionAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var response = await CompleteTextAsync("Hello", new AICompletionOptions { MaxTokens = 5 }, cancellationToken);
            _connection.LastTested = DateTime.Now;
            _connection.LastTestError = response.Success ? null : response.Error;
            return response.Success;
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
            var messages = new List<object>();
            
            // Add system message if provided
            if (!string.IsNullOrWhiteSpace(options.SystemPrompt))
            {
                messages.Add(new { role = "system", content = options.SystemPrompt });
            }
            
            // Add conversation history if provided
            if (options.ConversationHistory != null)
            {
                foreach (var msg in options.ConversationHistory)
                {
                    messages.Add(new { role = msg.Role.ToLowerInvariant(), content = msg.Content });
                }
            }
            
            // Add user prompt
            messages.Add(new { role = "user", content = prompt });
            
            var request = new
            {
                messages,
                temperature = options.Temperature,
                max_tokens = options.MaxTokens
            };
            
            var json = JsonSerializer.Serialize(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            
            var deployment = _connection.Deployment ?? _connection.Model;
            var response = await _httpClient.PostAsync(
                $"/openai/deployments/{deployment}/chat/completions?api-version=2024-02-01",
                content,
                cancellationToken
            );
            
            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync(cancellationToken);
                return AIResponse.FromError($"API Error: {response.StatusCode} - {errorContent}", sw.Elapsed);
            }
            
            var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
            var result = JsonDocument.Parse(responseJson);
            
            // Track usage
            var tokensUsed = result.RootElement.GetProperty("usage").GetProperty("total_tokens").GetInt32();
            _connection.TotalTokensUsed += tokensUsed;
            _connection.TotalRequests++;
            _connection.LastUsed = DateTime.Now;
            
            var resultText = result.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString() ?? string.Empty;
            
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
            // Load and encode image
            var imageBytes = await System.IO.File.ReadAllBytesAsync(imagePath, cancellationToken);
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
            
            var deployment = _connection.VisionModel ?? _connection.Deployment ?? _connection.Model;
            var response = await _httpClient.PostAsync(
                $"/openai/deployments/{deployment}/chat/completions?api-version=2024-02-01",
                content,
                cancellationToken
            );
            
            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync(cancellationToken);
                return AIResponse.FromError($"API Error: {response.StatusCode} - {errorContent}", sw.Elapsed);
            }
            
            var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
            var result = JsonDocument.Parse(responseJson);
            
            var tokensUsed = result.RootElement.GetProperty("usage").GetProperty("total_tokens").GetInt32();
            _connection.TotalTokensUsed += tokensUsed;
            _connection.TotalRequests++;
            _connection.LastUsed = DateTime.Now;
            
            var caption = result.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString() ?? string.Empty;
            
            return AIResponse.FromSuccess(caption, tokensUsed, sw.Elapsed);
        }
        catch (Exception ex)
        {
            return AIResponse.FromError($"Exception: {ex.Message}", sw.Elapsed);
        }
    }
    
    public async Task<AIResponse> GenerateImageAsync(string prompt, string size = "1024x1024", CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        
        try
        {
            var request = new
            {
                prompt,
                n = 1,
                size
            };
            
            var json = JsonSerializer.Serialize(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            
            var deployment = _connection.ImageModel ?? "dall-e-3";
            var response = await _httpClient.PostAsync(
                $"/openai/deployments/{deployment}/images/generations?api-version=2024-02-01",
                content,
                cancellationToken
            );
            
            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync(cancellationToken);
                return AIResponse.FromError($"API Error: {response.StatusCode} - {errorContent}", sw.Elapsed);
            }
            
            var responseJson = await response.Content.ReadAsStringAsync(cancellationToken);
            var result = JsonDocument.Parse(responseJson);
            
            _connection.TotalRequests++;
            _connection.LastUsed = DateTime.Now;
            
            var imageUrl = result.RootElement
                .GetProperty("data")[0]
                .GetProperty("url")
                .GetString() ?? string.Empty;
            
            return AIResponse.FromSuccess(imageUrl, 0, sw.Elapsed);
        }
        catch (Exception ex)
        {
            return AIResponse.FromError($"Exception: {ex.Message}", sw.Elapsed);
        }
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
        
        var messages = new List<object>
        {
            new { role = "user", content = prompt }
        };
        
        var request = new
        {
            messages,
            temperature = options.Temperature,
            max_tokens = options.MaxTokens,
            stream = true
        };
        
        var json = JsonSerializer.Serialize(request);
        var content = new StringContent(json, Encoding.UTF8, "application/json");
        
        var deployment = _connection.Deployment ?? _connection.Model;
        
        var requestMessage = new HttpRequestMessage(HttpMethod.Post, $"/openai/deployments/{deployment}/chat/completions?api-version=2024-02-01")
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
            if (string.IsNullOrWhiteSpace(line) || !line.StartsWith("data: "))
                continue;
            
            var data = line.Substring(6);
            if (data == "[DONE]")
                break;
            
            var json_data = JsonDocument.Parse(data);
            var delta = json_data.RootElement.GetProperty("choices")[0].GetProperty("delta");
            
            if (delta.TryGetProperty("content", out var contentProperty))
            {
                yield return contentProperty.GetString() ?? string.Empty;
            }
        }
    }
    
    private static string GetMimeType(string filePath)
    {
        var ext = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".jpg" or ".jpeg" => "image/jpeg",
            ".png" => "image/png",
            ".gif" => "image/gif",
            ".webp" => "image/webp",
            _ => "image/jpeg"
        };
    }
}
