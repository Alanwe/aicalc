using Xunit;
using AiCalc.Models;

namespace AiCalc.Tests;

public class WorkbookSettingsTests
{
    [Fact]
    public void Constructor_InitializesDefaults()
    {
        // Arrange & Act
        var settings = new WorkbookSettings();

        // Assert
        Assert.NotNull(settings.Connections);
        Assert.Empty(settings.Connections);
        Assert.Equal(4, settings.MaxEvaluationThreads); // Default to CPU count (typically 4 in tests)
        Assert.Equal(100, settings.DefaultEvaluationTimeoutSeconds);
    }

    [Fact]
    public void MaxEvaluationThreads_CanBeConfigured()
    {
        // Arrange
        var settings = new WorkbookSettings();

        // Act
        settings.MaxEvaluationThreads = 8;

        // Assert
        Assert.Equal(8, settings.MaxEvaluationThreads);
    }

    [Fact]
    public void DefaultEvaluationTimeoutSeconds_CanBeConfigured()
    {
        // Arrange
        var settings = new WorkbookSettings();

        // Act
        settings.DefaultEvaluationTimeoutSeconds = 300;

        // Assert
        Assert.Equal(300, settings.DefaultEvaluationTimeoutSeconds);
    }

    [Fact]
    public void Connections_CanBeAdded()
    {
        // Arrange
        var settings = new WorkbookSettings();
        var connection = new WorkspaceConnection
        {
            Name = "Azure OpenAI",
            Provider = "AzureOpenAI",
            IsActive = true
        };

        // Act
        settings.Connections.Add(connection);

        // Assert
        Assert.Single(settings.Connections);
        Assert.Equal("Azure OpenAI", settings.Connections[0].Name);
    }

    [Fact]
    public void ApplicationTheme_DefaultsToSystem()
    {
        // Arrange & Act
        var settings = new WorkbookSettings();

        // Assert
        Assert.Equal(AppTheme.System, settings.ApplicationTheme);
    }

    [Fact]
    public void SelectedTheme_DefaultsToLight()
    {
        // Arrange & Act
        var settings = new WorkbookSettings();

        // Assert
        Assert.Equal(CellVisualTheme.Light, settings.SelectedTheme);
    }
}

public class WorkspaceConnectionTests
{
    [Fact]
    public void Constructor_InitializesDefaults()
    {
        // Arrange & Act
        var connection = new WorkspaceConnection();

        // Assert
        Assert.Equal(100, connection.TimeoutSeconds);
        Assert.Equal(3, connection.MaxRetries);
        Assert.Equal(0.7, connection.Temperature);
        Assert.True(connection.IsActive); // Default is true
        Assert.Equal(0, connection.TotalTokensUsed);
        Assert.Equal(0, connection.TotalRequests);
    }

    [Fact]
    public void Provider_CanBeSet()
    {
        // Arrange
        var connection = new WorkspaceConnection();

        // Act
        connection.Provider = "Ollama";

        // Assert
        Assert.Equal("Ollama", connection.Provider);
    }

    [Fact]
    public void Models_CanBeConfigured()
    {
        // Arrange
        var connection = new WorkspaceConnection();

        // Act
        connection.Model = "gpt-4";
        connection.VisionModel = "gpt-4-vision";
        connection.ImageModel = "dall-e-3";

        // Assert
        Assert.Equal("gpt-4", connection.Model);
        Assert.Equal("gpt-4-vision", connection.VisionModel);
        Assert.Equal("dall-e-3", connection.ImageModel);
    }

    [Fact]
    public void UsageTracking_IncrementsProperly()
    {
        // Arrange
        var connection = new WorkspaceConnection
        {
            TotalTokensUsed = 100,
            TotalRequests = 5
        };

        // Act
        connection.TotalTokensUsed += 50;
        connection.TotalRequests += 1;

        // Assert
        Assert.Equal(150, connection.TotalTokensUsed);
        Assert.Equal(6, connection.TotalRequests);
    }

    [Fact]
    public void ConnectionTesting_RecordsResults()
    {
        // Arrange
        var connection = new WorkspaceConnection();
        var testTime = DateTime.UtcNow;

        // Act
        connection.LastTested = testTime;
        connection.LastTestError = null;
        connection.IsActive = true;

        // Assert
        Assert.Equal(testTime, connection.LastTested);
        Assert.Null(connection.LastTestError);
        Assert.True(connection.IsActive);
    }

    [Fact]
    public void ConnectionError_IsRecorded()
    {
        // Arrange
        var connection = new WorkspaceConnection();
        var errorMessage = "Connection timeout";

        // Act
        connection.LastTestError = errorMessage;
        connection.IsActive = false;

        // Assert
        Assert.Equal(errorMessage, connection.LastTestError);
        Assert.False(connection.IsActive);
    }
}
