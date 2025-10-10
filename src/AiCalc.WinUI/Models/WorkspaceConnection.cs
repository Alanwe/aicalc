using System;

namespace AiCalc.Models;

/// <summary>
/// Represents a connection to an AI service provider
/// </summary>
public class WorkspaceConnection
{
    public Guid Id { get; set; } = Guid.NewGuid();
    
    public string Name { get; set; } = "Local Runtime";

    /// <summary>
    /// Provider type: "AzureOpenAI", "Ollama", "OpenAI", "Local"
    /// </summary>
    public string Provider { get; set; } = "Local";

    public string Endpoint { get; set; } = "http://localhost";

    /// <summary>
    /// Encrypted API key (use CredentialService to encrypt/decrypt)
    /// </summary>
    public string ApiKey { get; set; } = string.Empty;
    
    /// <summary>
    /// Default model for text completion
    /// </summary>
    public string Model { get; set; } = string.Empty;
    
    /// <summary>
    /// Azure OpenAI deployment name
    /// </summary>
    public string? Deployment { get; set; }
    
    /// <summary>
    /// Model for vision tasks (image captioning)
    /// </summary>
    public string? VisionModel { get; set; }
    
    /// <summary>
    /// Model for image generation
    /// </summary>
    public string? ImageModel { get; set; }

    public bool IsDefault { get; set; }
    
    /// <summary>
    /// Timeout in seconds for API calls
    /// </summary>
    public int TimeoutSeconds { get; set; } = 100;
    
    /// <summary>
    /// Maximum retry attempts for failed requests
    /// </summary>
    public int MaxRetries { get; set; } = 3;
    
    /// <summary>
    /// Temperature for AI generation (0.0-2.0)
    /// </summary>
    public double Temperature { get; set; } = 0.7;
    
    /// <summary>
    /// Total tokens used across all requests
    /// </summary>
    public long TotalTokensUsed { get; set; }
    
    /// <summary>
    /// Total number of requests made
    /// </summary>
    public long TotalRequests { get; set; }
    
    /// <summary>
    /// Last time this connection was used
    /// </summary>
    public DateTime? LastUsed { get; set; }
    
    /// <summary>
    /// Whether this connection is active
    /// </summary>
    public bool IsActive { get; set; } = true;
    
    /// <summary>
    /// Last time connection was tested
    /// </summary>
    public DateTime? LastTested { get; set; }
    
    /// <summary>
    /// Error from last test (if any)
    /// </summary>
    public string? LastTestError { get; set; }
}
