using AiCalc.Models;
using AiCalc.Services;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;

namespace AiCalc;

// UserControl for the XAML content
public sealed partial class DataSourcesDialogContent : UserControl, INotifyPropertyChanged
{
    public ObservableCollection<CloudStorageConnection> EditableConnections { get; } = new();
    public Array Providers { get; } = Enum.GetValues(typeof(CloudStorageProvider));
    
    private CloudStorageConnection? _selectedConnection;
    public CloudStorageConnection? SelectedConnection
    {
        get => _selectedConnection;
        set
        {
            if (_selectedConnection == value)
            {
                return;
            }

            _selectedConnection = value;
            UpdatePasswordEditors();
            UpdateNameFollowState();
            OnPropertyChanged(nameof(SelectedConnection));
            OnPropertyChanged(nameof(AzureFieldsVisibility));
            OnPropertyChanged(nameof(AwsFieldsVisibility));
            OnPropertyChanged(nameof(GoogleFieldsVisibility));
            OnPropertyChanged(nameof(HasSelectedConnection));
            OnPropertyChanged(nameof(IsTestConnectionEnabled));
            OnPropertyChanged(nameof(DetailsVisibility));
            OnPropertyChanged(nameof(EmptyStateVisibility));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private bool _nameFollowsAzureAccount = true;
    private bool _suppressDisplayNameChange;
    private bool _isTestingConnection;
    private CancellationTokenSource? _testCancellationSource;

    private PasswordBox? _azureAccessKeyBox;
    private PasswordBox? _azureConnectionStringBox;
    private PasswordBox? _awsSecretKeyBox;
    private ListView? _connectionsList;
    private Button? _testConnectionButton;
    private InfoBar? _testResultInfo;
    private TextBox? _displayNameBox;

    public DataSourcesDialogContent()
    {
        this.InitializeComponent();

        _azureAccessKeyBox = (PasswordBox?)FindName("AzureAccessKeyBox");
        _azureConnectionStringBox = (PasswordBox?)FindName("AzureConnectionStringBox");
        _awsSecretKeyBox = (PasswordBox?)FindName("AwsSecretKeyBox");
        _connectionsList = (ListView?)FindName("ConnectionsList");
        _testConnectionButton = (Button?)FindName("TestConnectionButton");
        _testResultInfo = (InfoBar?)FindName("TestResultInfo");
        _displayNameBox = (TextBox?)FindName("DisplayNameBox");
    }
    
    public void LoadConnections(ObservableCollection<CloudStorageConnection> connections)
    {
        EditableConnections.Clear();
        foreach (var conn in connections)
        {
            EditableConnections.Add(conn);
        }

        if (_connectionsList != null)
        {
            _connectionsList.SelectedIndex = -1;
        }
    }

    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private void UpdatePasswordEditors()
    {
        if (_azureAccessKeyBox != null)
        {
            _azureAccessKeyBox.Password = SelectedConnection?.AzureStorageAccountKey ?? string.Empty;
        }

        if (_azureConnectionStringBox != null)
        {
            _azureConnectionStringBox.Password = SelectedConnection?.AzureConnectionString ?? string.Empty;
        }

        if (_awsSecretKeyBox != null)
        {
            _awsSecretKeyBox.Password = SelectedConnection?.AwsSecretAccessKey ?? string.Empty;
        }

        if (_testResultInfo != null)
        {
            _testResultInfo.IsOpen = false;
        }
    }

    private void AddConnection_Click(object sender, RoutedEventArgs e)
    {
        var connection = new CloudStorageConnection
        {
            Name = "New Data Source",
            Provider = CloudStorageProvider.AzureBlob
        };

        EditableConnections.Add(connection);
        SelectedConnection = connection;
        if (_connectionsList != null)
        {
            _connectionsList.SelectedItem = connection;
        }
    }

    private void RemoveConnection_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedConnection == null)
        {
            return;
        }

        EditableConnections.Remove(SelectedConnection);
        SelectedConnection = EditableConnections.FirstOrDefault();
        if (_connectionsList != null)
        {
            _connectionsList.SelectedItem = SelectedConnection;
        }
    }

    private void AzureAccessKeyBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (SelectedConnection != null && sender is PasswordBox passwordBox)
        {
            SelectedConnection.AzureStorageAccountKey = passwordBox.Password;
            ResetTestResult();
        }
    }

    private void AzureConnectionStringBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (SelectedConnection != null && sender is PasswordBox passwordBox)
        {
            SelectedConnection.AzureConnectionString = passwordBox.Password;
            ResetTestResult();
        }
    }

    private void AwsSecretKeyBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (SelectedConnection != null && sender is PasswordBox passwordBox)
        {
            SelectedConnection.AwsSecretAccessKey = passwordBox.Password;
            ResetTestResult();
        }
    }

    private void ConnectionsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is ListView listView)
        {
            SelectedConnection = listView.SelectedItem as CloudStorageConnection;
        }
    }

    public Visibility AzureFieldsVisibility => SelectedConnection?.Provider == CloudStorageProvider.AzureBlob ? Visibility.Visible : Visibility.Collapsed;

    public Visibility AwsFieldsVisibility => SelectedConnection?.Provider == CloudStorageProvider.AwsS3 ? Visibility.Visible : Visibility.Collapsed;

    public Visibility GoogleFieldsVisibility => SelectedConnection?.Provider == CloudStorageProvider.GoogleCloudStorage ? Visibility.Visible : Visibility.Collapsed;

    public bool HasSelectedConnection => SelectedConnection != null;

    public bool IsTestConnectionEnabled => HasSelectedConnection && !_isTestingConnection;

    public Visibility DetailsVisibility => HasSelectedConnection ? Visibility.Visible : Visibility.Collapsed;

    public Visibility EmptyStateVisibility => HasSelectedConnection ? Visibility.Collapsed : Visibility.Visible;

    private void ProviderBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (SelectedConnection == null)
        {
            return;
        }

        ResetTestResult();
        UpdateNameFollowState();
        OnPropertyChanged(nameof(AzureFieldsVisibility));
        OnPropertyChanged(nameof(AwsFieldsVisibility));
        OnPropertyChanged(nameof(GoogleFieldsVisibility));
        OnPropertyChanged(nameof(IsTestConnectionEnabled));
    }

    private void AzureAccountNameBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (SelectedConnection == null || sender is not TextBox textBox)
        {
            return;
        }

        SelectedConnection.AzureStorageAccountName = textBox.Text?.Trim() ?? string.Empty;

        if (_displayNameBox == null)
        {
            _displayNameBox = (TextBox?)FindName("DisplayNameBox");
        }

        if (_nameFollowsAzureAccount && SelectedConnection.Provider == CloudStorageProvider.AzureBlob && _displayNameBox != null)
        {
            _suppressDisplayNameChange = true;
            SelectedConnection.Name = SelectedConnection.AzureStorageAccountName;
            _displayNameBox.Text = SelectedConnection.AzureStorageAccountName;
            _suppressDisplayNameChange = false;
        }

        ResetTestResult();
    }

    private void DisplayNameBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressDisplayNameChange || SelectedConnection == null || sender is not TextBox textBox)
        {
            return;
        }

        _nameFollowsAzureAccount = SelectedConnection.Provider == CloudStorageProvider.AzureBlob &&
            (string.IsNullOrWhiteSpace(textBox.Text) || string.Equals(textBox.Text, SelectedConnection.AzureStorageAccountName, StringComparison.OrdinalIgnoreCase));

        ResetTestResult();
    }

    private void UpdateNameFollowState()
    {
        if (SelectedConnection == null)
        {
            _nameFollowsAzureAccount = true;
            return;
        }

        _nameFollowsAzureAccount = SelectedConnection.Provider == CloudStorageProvider.AzureBlob &&
            (string.IsNullOrWhiteSpace(SelectedConnection.Name) ||
             string.Equals(SelectedConnection.Name, SelectedConnection.AzureStorageAccountName, StringComparison.OrdinalIgnoreCase));
    }

    private async void TestConnection_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedConnection == null || _testConnectionButton == null)
        {
            return;
        }

        if (_isTestingConnection)
        {
            _testCancellationSource?.Cancel();
            return;
        }

        ResetTestResult();

        _isTestingConnection = true;
        OnPropertyChanged(nameof(IsTestConnectionEnabled));
        _testCancellationSource = new CancellationTokenSource();

        try
        {
            var clone = SelectedConnection.Clone();
            var result = await CloudStorageConnectionTester.TestAsync(clone, _testCancellationSource.Token);

            if (_testResultInfo != null)
            {
                _testResultInfo.Severity = result.IsSuccess ? InfoBarSeverity.Success : InfoBarSeverity.Error;
                _testResultInfo.Title = result.IsSuccess ? "Connection test succeeded" : "Connection test failed";
                _testResultInfo.Message = result.Message;
                _testResultInfo.IsOpen = true;
            }
        }
        catch (OperationCanceledException)
        {
            if (_testResultInfo != null)
            {
                _testResultInfo.Severity = InfoBarSeverity.Informational;
                _testResultInfo.Title = "Connection test canceled";
                _testResultInfo.Message = "The connection test was canceled.";
                _testResultInfo.IsOpen = true;
            }
        }
        catch (Exception ex)
        {
            if (_testResultInfo != null)
            {
                _testResultInfo.Severity = InfoBarSeverity.Error;
                _testResultInfo.Title = "Connection test failed";
                _testResultInfo.Message = ex.Message;
                _testResultInfo.IsOpen = true;
            }
        }
        finally
        {
            _testCancellationSource?.Dispose();
            _testCancellationSource = null;
            _isTestingConnection = false;
            OnPropertyChanged(nameof(IsTestConnectionEnabled));
        }
    }

    private void ResetTestResult()
    {
        if (_testResultInfo != null)
        {
            _testResultInfo.IsOpen = false;
            _testResultInfo.Title = string.Empty;
            _testResultInfo.Message = string.Empty;
        }
    }

    private void CredentialField_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (SelectedConnection == null)
        {
            return;
        }

        ResetTestResult();
    }
}

// Actual Dialog Window class
public sealed class DataSourcesDialog : DialogWindowBase
{
    private readonly WorkbookSettings _settings;
    private readonly DataSourcesDialogContent _content;

    public DataSourcesDialog(WorkbookSettings settings)
    {
        _settings = settings;
        
        // Set dialog properties before initialization
        DialogTitle = "Cloud Data Sources";
        DialogWidth = 1150;
        DialogHeight = 650;
        PrimaryButtonText = "Save";
        CloseButtonText = "Cancel";
        
        // Create the content UserControl
        _content = new DataSourcesDialogContent();
        
        // Load connections into the content
        var editableConnections = new ObservableCollection<CloudStorageConnection>();
        foreach (var connection in settings.CloudStorageConnections)
        {
            editableConnections.Add(connection.Clone());
        }
        _content.LoadConnections(editableConnections);
        
        // Set the content
        SetDialogContent(_content);
    }
    
    protected override bool OnPrimaryButtonClick()
    {
        if (_content.EditableConnections.Any(c => string.IsNullOrWhiteSpace(c.Name)))
        {
            return false; // Keep dialog open if validation fails
        }

        _settings.CloudStorageConnections.Clear();
        foreach (var connection in _content.EditableConnections)
        {
            _settings.CloudStorageConnections.Add(connection.Clone());
        }
        
        return true; // Close dialog on success
    }
}
