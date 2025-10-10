using AiCalc.Models;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;

namespace AiCalc;

public sealed partial class ServiceConnectionDialog : ContentDialog
{
    public WorkspaceConnection Connection { get; private set; }

    public ServiceConnectionDialog() : this(null)
    {
    }

    public ServiceConnectionDialog(WorkspaceConnection? existingConnection)
    {
        // Create a copy to avoid modifying the original until saved
        if (existingConnection != null)
        {
            Connection = new WorkspaceConnection
            {
                Id = existingConnection.Id,
                Name = existingConnection.Name,
                Provider = existingConnection.Provider,
                Endpoint = existingConnection.Endpoint,
                ApiKey = existingConnection.ApiKey,
                Model = existingConnection.Model,
                Deployment = existingConnection.Deployment,
                IsDefault = existingConnection.IsDefault
            };
        }
        else
        {
            Connection = new WorkspaceConnection
            {
                Name = "",
                Provider = "Azure OpenAI",
                Endpoint = "",
                ApiKey = "",
                Model = "",
                IsDefault = false
            };
        }

        InitializeComponent();
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
        Connection.Provider = "Azure OpenAI";
        Connection.Endpoint = "https://your-resource.openai.azure.com/";
        Connection.Model = "gpt-4";
        Connection.Deployment = "gpt-4-deployment";
        
        // Trigger UI update
        ProviderComboBox.SelectedItem = "Azure OpenAI";
        EndpointTextBox.Text = Connection.Endpoint;
        ModelTextBox.Text = Connection.Model;
        DeploymentTextBox.Text = Connection.Deployment;
    }

    private void PresetOpenAI_Click(object sender, RoutedEventArgs e)
    {
        Connection.Provider = "OpenAI";
        Connection.Endpoint = "https://api.openai.com/v1";
        Connection.Model = "gpt-4";
        Connection.Deployment = null;
        
        // Trigger UI update
        ProviderComboBox.SelectedItem = "OpenAI";
        EndpointTextBox.Text = Connection.Endpoint;
        ModelTextBox.Text = Connection.Model;
        DeploymentTextBox.Text = "";
    }

    private void PresetOllama_Click(object sender, RoutedEventArgs e)
    {
        Connection.Provider = "Local (Ollama)";
        Connection.Endpoint = "http://localhost:11434";
        Connection.Model = "llama2";
        Connection.Deployment = null;
        Connection.ApiKey = "";
        
        // Trigger UI update
        ProviderComboBox.SelectedItem = "Local (Ollama)";
        EndpointTextBox.Text = Connection.Endpoint;
        ModelTextBox.Text = Connection.Model;
        DeploymentTextBox.Text = "";
        ApiKeyPasswordBox.Password = "";
    }
}
