namespace AiCalc.Models;

public class WorkspaceConnection
{
    public string Name { get; set; } = "Local Runtime";

    public string Provider { get; set; } = "Local";

    public string Endpoint { get; set; } = "http://localhost";

    public string ApiKey { get; set; } = string.Empty;

    public bool IsDefault { get; set; }
}
