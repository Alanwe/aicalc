using System;

namespace AiCalc.Models;

public class WorkspaceConnection
{
    public Guid Id { get; set; } = Guid.NewGuid();
    
    public string Name { get; set; } = "Local Runtime";

    public string Provider { get; set; } = "Local";

    public string Endpoint { get; set; } = "http://localhost";

    public string ApiKey { get; set; } = string.Empty;
    
    public string Model { get; set; } = string.Empty;
    
    public string? Deployment { get; set; }

    public bool IsDefault { get; set; }
}
