using AiCalc.Models;
using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace AiCalc.Services;

public sealed record CloudConnectionTestResult(bool IsSuccess, string Message, HttpStatusCode? StatusCode = null);

public static class CloudStorageConnectionTester
{
    private static readonly HttpClient HttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(20)
    };

    public static async Task<CloudConnectionTestResult> TestAsync(CloudStorageConnection connection, CancellationToken cancellationToken = default)
    {
        if (connection == null)
        {
            throw new ArgumentNullException(nameof(connection));
        }

        return connection.Provider switch
        {
            CloudStorageProvider.AzureBlob => await TestAzureAsync(connection, cancellationToken),
            CloudStorageProvider.AwsS3 => await TestAwsAsync(connection, cancellationToken),
            CloudStorageProvider.GoogleCloudStorage => await TestGoogleAsync(connection, cancellationToken),
            _ => new CloudConnectionTestResult(false, "The selected provider is not supported yet.")
        };
    }

    private static async Task<CloudConnectionTestResult> TestAzureAsync(CloudStorageConnection connection, CancellationToken cancellationToken)
    {
        var accountName = ResolveAzureAccountName(connection);
        if (string.IsNullOrWhiteSpace(accountName))
        {
            return new CloudConnectionTestResult(false, "Provide a storage account name or connection string.");
        }

        var uri = new Uri($"https://{accountName}.blob.core.windows.net/?comp=list&maxresults=1");
        return await ProbeEndpointAsync(uri, "Azure Blob Storage", cancellationToken);
    }

    private static async Task<CloudConnectionTestResult> TestAwsAsync(CloudStorageConnection connection, CancellationToken cancellationToken)
    {
        var region = string.IsNullOrWhiteSpace(connection.AwsRegion) ? "us-east-1" : connection.AwsRegion.Trim();
        var uri = new Uri($"https://s3.{region}.amazonaws.com/");
        return await ProbeEndpointAsync(uri, "AWS S3", cancellationToken);
    }

    private static async Task<CloudConnectionTestResult> TestGoogleAsync(CloudStorageConnection connection, CancellationToken cancellationToken)
    {
        var uri = new Uri("https://storage.googleapis.com/");
        return await ProbeEndpointAsync(uri, "Google Cloud Storage", cancellationToken);
    }

    private static async Task<CloudConnectionTestResult> ProbeEndpointAsync(Uri uri, string serviceName, CancellationToken cancellationToken)
    {
        try
        {
            using var headRequest = new HttpRequestMessage(HttpMethod.Head, uri);
            using var headResponse = await HttpClient.SendAsync(headRequest, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

            var headEvaluation = EvaluateResponse(headResponse, serviceName);
            if (headResponse.StatusCode != HttpStatusCode.MethodNotAllowed || headEvaluation.IsSuccess)
            {
                return headEvaluation;
            }

            using var getRequest = new HttpRequestMessage(HttpMethod.Get, uri);
            using var getResponse = await HttpClient.SendAsync(getRequest, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            return EvaluateResponse(getResponse, serviceName);
        }
        catch (HttpRequestException ex)
        {
            return new CloudConnectionTestResult(false, $"Network error contacting {serviceName}: {ex.Message}");
        }
    }

    private static CloudConnectionTestResult EvaluateResponse(HttpResponseMessage response, string serviceName)
    {
        if (response.StatusCode is HttpStatusCode.OK or HttpStatusCode.Forbidden or HttpStatusCode.Unauthorized)
        {
            return new CloudConnectionTestResult(true, $"Reached {serviceName} endpoint (HTTP {(int)response.StatusCode}).", response.StatusCode);
        }

        return new CloudConnectionTestResult(false, $"Received HTTP {(int)response.StatusCode} from {serviceName}.", response.StatusCode);
    }

    private static string? ResolveAzureAccountName(CloudStorageConnection connection)
    {
        if (!string.IsNullOrWhiteSpace(connection.AzureStorageAccountName))
        {
            return connection.AzureStorageAccountName.Trim();
        }

        if (!string.IsNullOrWhiteSpace(connection.AzureConnectionString))
        {
            var segments = connection.AzureConnectionString.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var segment in segments)
            {
                var parts = segment.Split('=', 2);
                if (parts.Length == 2 && parts[0].Equals("AccountName", StringComparison.OrdinalIgnoreCase))
                {
                    return parts[1];
                }
            }
        }

        return null;
    }
}
