using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AiCalc.Models;

namespace AiCalc.Services.AI;

/// <summary>
/// Registry for managing AI service connections and clients
/// </summary>
public class AIServiceRegistry
{
    private readonly Dictionary<Guid, WorkspaceConnection> _connections = new();
    private readonly Dictionary<Guid, IAIServiceClient> _clients = new();
    
    /// <summary>
    /// Register a new connection and create its client
    /// </summary>
    public void RegisterConnection(WorkspaceConnection connection)
    {
        if (connection == null)
            throw new ArgumentNullException(nameof(connection));
        
        _connections[connection.Id] = connection;
        _clients[connection.Id] = CreateClient(connection);
    }
    
    /// <summary>
    /// Update an existing connection
    /// </summary>
    public void UpdateConnection(WorkspaceConnection connection)
    {
        if (connection == null)
            throw new ArgumentNullException(nameof(connection));
        
        _connections[connection.Id] = connection;
        
        // Recreate client with updated connection
        if (_clients.ContainsKey(connection.Id))
        {
            _clients[connection.Id] = CreateClient(connection);
        }
    }
    
    /// <summary>
    /// Remove a connection
    /// </summary>
    public void RemoveConnection(Guid connectionId)
    {
        _connections.Remove(connectionId);
        _clients.Remove(connectionId);
    }
    
    /// <summary>
    /// Get client by connection ID
    /// </summary>
    public IAIServiceClient? GetClient(Guid connectionId)
    {
        return _clients.TryGetValue(connectionId, out var client) ? client : null;
    }
    
    /// <summary>
    /// Get connection by ID
    /// </summary>
    public WorkspaceConnection? GetConnection(Guid connectionId)
    {
        return _connections.TryGetValue(connectionId, out var connection) ? connection : null;
    }
    
    /// <summary>
    /// Get all connections
    /// </summary>
    public IEnumerable<WorkspaceConnection> GetAllConnections()
    {
        return _connections.Values;
    }
    
    /// <summary>
    /// Get default connection (marked as default or first active)
    /// </summary>
    public WorkspaceConnection? GetDefaultConnection()
    {
        return _connections.Values.FirstOrDefault(c => c.IsDefault && c.IsActive) 
            ?? _connections.Values.FirstOrDefault(c => c.IsActive);
    }
    
    /// <summary>
    /// Get default client
    /// </summary>
    public IAIServiceClient? GetDefaultClient()
    {
        var connection = GetDefaultConnection();
        return connection != null ? GetClient(connection.Id) : null;
    }
    
    /// <summary>
    /// Test a connection
    /// </summary>
    public async Task<bool> TestConnectionAsync(Guid connectionId)
    {
        var connection = GetConnection(connectionId);
        var client = GetClient(connectionId);
        
        if (connection == null || client == null)
            return false;
        
        try
        {
            return await client.TestConnectionAsync();
        }
        catch (Exception ex)
        {
            connection.LastTestError = ex.Message;
            connection.LastTested = DateTime.Now;
            return false;
        }
    }
    
    /// <summary>
    /// Get active connections by provider type
    /// </summary>
    public IEnumerable<WorkspaceConnection> GetConnectionsByProvider(string provider)
    {
        return _connections.Values.Where(c => c.IsActive && c.Provider == provider);
    }
    
    /// <summary>
    /// Create appropriate client for connection
    /// </summary>
    private static IAIServiceClient CreateClient(WorkspaceConnection connection)
    {
        return connection.Provider switch
        {
            "AzureOpenAI" => new AzureOpenAIClient(connection),
            "Ollama" => new OllamaClient(connection),
            "OpenAI" => new AzureOpenAIClient(connection), // Similar to Azure but different endpoint
            _ => throw new NotSupportedException($"Provider '{connection.Provider}' is not supported")
        };
    }
    
    /// <summary>
    /// Clear all connections
    /// </summary>
    public void Clear()
    {
        _connections.Clear();
        _clients.Clear();
    }
}
