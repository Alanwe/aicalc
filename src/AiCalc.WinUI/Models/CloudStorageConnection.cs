using System;

namespace AiCalc.Models;

public class CloudStorageConnection
{
    public Guid Id { get; set; } = Guid.NewGuid();

    public string Name { get; set; } = string.Empty;

    public CloudStorageProvider Provider { get; set; } = CloudStorageProvider.AzureBlob;

    public string AzureStorageAccountName { get; set; } = string.Empty;

    public string AzureStorageAccountKey { get; set; } = string.Empty;

    public string AzureConnectionString { get; set; } = string.Empty;

    public string AwsAccessKeyId { get; set; } = string.Empty;

    public string AwsSecretAccessKey { get; set; } = string.Empty;

    public string AwsRegion { get; set; } = string.Empty;

    public string GoogleProjectId { get; set; } = string.Empty;

    public string GoogleClientEmail { get; set; } = string.Empty;

    public string GoogleJsonKey { get; set; } = string.Empty;

    public CloudStorageConnection Clone()
    {
        return new CloudStorageConnection
        {
            Id = Id,
            Name = Name,
            Provider = Provider,
            AzureStorageAccountName = AzureStorageAccountName,
            AzureStorageAccountKey = AzureStorageAccountKey,
            AzureConnectionString = AzureConnectionString,
            AwsAccessKeyId = AwsAccessKeyId,
            AwsSecretAccessKey = AwsSecretAccessKey,
            AwsRegion = AwsRegion,
            GoogleProjectId = GoogleProjectId,
            GoogleClientEmail = GoogleClientEmail,
            GoogleJsonKey = GoogleJsonKey
        };
    }
}
