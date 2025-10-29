using System;

namespace AiCalc.Models;

public class CloudImportConfiguration
{
    public Guid ConnectionId { get; set; }

    public string ConnectionName { get; set; } = string.Empty;

    public string ContainerName { get; set; } = string.Empty;

    public string FilterPattern { get; set; } = "*.*";

    public CloudImportConfiguration Clone()
    {
        return new CloudImportConfiguration
        {
            ConnectionId = ConnectionId,
            ConnectionName = ConnectionName,
            ContainerName = ContainerName,
            FilterPattern = FilterPattern
        };
    }

    public override string ToString()
    {
        return string.IsNullOrWhiteSpace(ContainerName)
            ? $"Cloud import from {ConnectionName}"
            : $"{ConnectionName} • {ContainerName} ({FilterPattern})";
    }
}
