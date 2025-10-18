using AiCalc.Models;
using AiCalc.Services.AI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using System;
using System.Threading.Tasks;

namespace AiCalc;

public sealed partial class ServiceConnectionDialog : ContentDialog
{
    private string _apiKeyPlainText = string.Empty;
    private PasswordBox _apiKeyBox = null!;
    private ComboBox _providerComboBox = null!;
    private TextBox _endpointTextBox = null!;
    private TextBox _modelTextBox = null!;
    private TextBox _visionModelTextBox = null!;
    private TextBox _deploymentTextBox = null!;
    private Button _testConnectionButton = null!;
    private TextBlock _connectionStatusText = null!;

    public WorkspaceConnection Connection { get; private set; }

    public ServiceConnectionDialog() : this(null)
    {
    }

    public ServiceConnectionDialog(WorkspaceConnection? existingConnection)
    {
        // Create a copy to avoid modifying the original until saved
        if (existingConnection != null)
        {
            Connection = existingConnection.Clone();
            if (!string.IsNullOrWhiteSpace(existingConnection.ApiKey))
            {
                _apiKeyPlainText = CredentialService.IsEncrypted(existingConnection.ApiKey)
                    ? CredentialService.Decrypt(existingConnection.ApiKey)
                    : existingConnection.ApiKey;
            }
        }
        else
        {
            Connection = new WorkspaceConnection
            {
                Name = string.Empty,
                Provider = "AzureOpenAI",
                Endpoint = string.Empty,
                ApiKey = string.Empty,
                Model = "gpt-4",
                VisionModel = "gpt-4-vision-preview",
                ImageModel = "dall-e-3",
                TimeoutSeconds = 100,
                MaxRetries = 3,
                Temperature = 0.7,
                IsDefault = false
            };
        }

    this.InitializeComponent();
    _apiKeyBox = GetRequiredElement<PasswordBox>("ApiKeyPasswordBox");
    _providerComboBox = GetRequiredElement<ComboBox>("ProviderComboBox");
    _endpointTextBox = GetRequiredElement<TextBox>("EndpointTextBox");
    _modelTextBox = GetRequiredElement<TextBox>("ModelTextBox");
    _visionModelTextBox = GetRequiredElement<TextBox>("VisionModelTextBox");
    _deploymentTextBox = GetRequiredElement<TextBox>("DeploymentTextBox");
    _testConnectionButton = GetRequiredElement<Button>("TestConnectionButton");
    _connectionStatusText = GetRequiredElement<TextBlock>("ConnectionStatusText");

    _apiKeyBox.Password = _apiKeyPlainText;
    }

    private void SaveButton_Click(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        // Validate required fields
        if (string.IsNullOrWhiteSpace(Connection.Name))
        {
            args.Cancel = true;
            ShowValidationError("Service Name is required");
            return;
        }

        if (string.IsNullOrWhiteSpace(Connection.Endpoint))
        {
            args.Cancel = true;
            ShowValidationError("Endpoint is required");
            return;
        }

        if (string.IsNullOrWhiteSpace(Connection.Provider))
        {
            args.Cancel = true;
            ShowValidationError("Provider is required");
            return;
        }

        // Validate endpoint format
        if (!Uri.TryCreate(Connection.Endpoint, UriKind.Absolute, out _))
        {
            args.Cancel = true;
            ShowValidationError("Endpoint must be a valid URL");
            return;
        }

        Connection.ApiKey = string.IsNullOrWhiteSpace(_apiKeyPlainText)
            ? string.Empty
            : CredentialService.Encrypt(_apiKeyPlainText);
    }

    private async void ShowValidationError(string message)
    {
        var errorDialog = new ContentDialog
        {
            Title = "Validation Error",
            Content = message,
            CloseButtonText = "OK",
            XamlRoot = this.XamlRoot
        };
        await errorDialog.ShowAsync();
    }

    private void PresetAzureOpenAI_Click(object sender, RoutedEventArgs e)
    {
        Connection.Provider = "AzureOpenAI";
        Connection.Endpoint = "https://your-resource.openai.azure.com/";
        Connection.Model = "gpt-4";
        Connection.Deployment = "gpt-4-deployment";
        
        // Trigger UI update
    _providerComboBox.SelectedValue = Connection.Provider;
    _endpointTextBox.Text = Connection.Endpoint;
    _modelTextBox.Text = Connection.Model;
    _deploymentTextBox.Text = Connection.Deployment;
    }

    private void PresetOpenAI_Click(object sender, RoutedEventArgs e)
    {
        ShowValidationError("Native OpenAI support is coming soon. Please configure Azure OpenAI or Ollama.");
        return;
    }

    private void PresetOllama_Click(object sender, RoutedEventArgs e)
    {
        Connection.Provider = "Ollama";
        Connection.Endpoint = "http://localhost:11434";
        Connection.Model = "llama2";
        Connection.VisionModel = "llava";
        Connection.Deployment = null;
        _apiKeyPlainText = string.Empty;
        
        // Trigger UI update
    _providerComboBox.SelectedValue = Connection.Provider;
    _endpointTextBox.Text = Connection.Endpoint;
    _modelTextBox.Text = Connection.Model;
    _visionModelTextBox.Text = Connection.VisionModel;
    _deploymentTextBox.Text = string.Empty;
    _apiKeyBox.Password = string.Empty;
    }

    private async void TestConnection_Click(object sender, RoutedEventArgs e)
    {
    _testConnectionButton.IsEnabled = false;
    _connectionStatusText.Visibility = Visibility.Visible;
    _connectionStatusText.Foreground = new SolidColorBrush(Microsoft.UI.Colors.Gray);
    _connectionStatusText.Text = "Testing connection...";

        try
        {
            // Validate required fields first
            if (string.IsNullOrWhiteSpace(Connection.Endpoint))
            {
                _connectionStatusText.Foreground = new SolidColorBrush(Microsoft.UI.Colors.Red);
                _connectionStatusText.Text = "❌ Error: Endpoint is required";
                return;
            }

            // Create temporary client for testing
            var testConnection = Connection.Clone();
            testConnection.ApiKey = string.IsNullOrWhiteSpace(_apiKeyPlainText)
                ? string.Empty
                : CredentialService.Encrypt(_apiKeyPlainText);

            IAIServiceClient client = testConnection.Provider switch
            {
                "AzureOpenAI" => new AzureOpenAIClient(testConnection),
                "Ollama" => new OllamaClient(testConnection),
                _ => throw new NotSupportedException($"Provider '{testConnection.Provider}' is not yet supported")
            };

            var result = await client.TestConnectionAsync();
            
            if (result.Success)
            {
                _connectionStatusText.Foreground = new SolidColorBrush(Microsoft.UI.Colors.Green);
                _connectionStatusText.Text = $"✅ Connection successful! {result.Result}";
                Connection.IsActive = true;
                Connection.LastTested = DateTime.Now;
                Connection.LastTestError = null;
                Connection.ApiKey = testConnection.ApiKey;
            }
            else
            {
                _connectionStatusText.Foreground = new SolidColorBrush(Microsoft.UI.Colors.Red);
                _connectionStatusText.Text = $"❌ Connection failed: {result.Error}";
                Connection.IsActive = false;
                Connection.LastTested = DateTime.Now;
                Connection.LastTestError = result.Error;
            }
        }
        catch (Exception ex)
        {
            _connectionStatusText.Foreground = new SolidColorBrush(Microsoft.UI.Colors.Red);
            _connectionStatusText.Text = $"❌ Error: {ex.Message}";
            Connection.IsActive = false;
            Connection.LastTested = DateTime.Now;
            Connection.LastTestError = ex.Message;
        }
        finally
        {
            _testConnectionButton.IsEnabled = true;
        }
    }

    private void ApiKeyPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (sender is PasswordBox passwordBox)
        {
            _apiKeyPlainText = passwordBox.Password;
        }
        else
        {
            _apiKeyPlainText = _apiKeyBox.Password;
        }
    }

    private T GetRequiredElement<T>(string name) where T : class
    {
        var element = FindName(name) as T;
        if (element == null)
        {
            throw new InvalidOperationException($"Could not locate element '{name}' in ServiceConnectionDialog template.");
        }

        return element;
    }
}
