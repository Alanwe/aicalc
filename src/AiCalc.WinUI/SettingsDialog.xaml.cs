using AiCalc.Models;
using AiCalc.Services;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.Storage.Pickers;
using Windows.UI;

namespace AiCalc;

public sealed partial class SettingsDialog : ContentDialog
{
    public WorkbookSettings Settings { get; }
    private List<PythonEnvironmentDetector.PythonEnvironment> _pythonEnvironments = new();
    private readonly ObservableCollection<PythonFunctionInfo> _discoveredFunctions = new();
    private FileSystemWatcher? _functionsWatcher;
    private bool _pendingHotReloadScan;
    private readonly HashSet<string> _registeredPythonFunctionNames = new(StringComparer.OrdinalIgnoreCase);

    public SettingsDialog(WorkbookSettings settings)
    {
        Settings = settings;
        InitializeComponent();
        
        // Initialize Performance tab
        var cpuCores = Environment.ProcessorCount;
        MaxThreadsLabel.Text = Settings.MaxEvaluationThreads.ToString();
        MaxThreadsDescription.Text = $"Using {Settings.MaxEvaluationThreads} threads (CPU cores detected: {cpuCores})";
        
        // Initialize Appearance tab - Application Theme
        AppThemeComboBox.SelectedIndex = (int)Settings.ApplicationTheme;
        
        // Initialize Appearance tab - Cell Visual Theme
        ThemeComboBox.SelectedIndex = (int)Settings.SelectedTheme;
        UpdateThemePreview(Settings.SelectedTheme);
        
        // Initialize AutoSave settings (Phase 6)
        var prefs = App.PreferencesService.LoadPreferences();
        AutoSaveToggle.IsOn = prefs.AutoSaveEnabled;
        AutoSaveIntervalSlider.Value = prefs.AutoSaveIntervalMinutes;
        AutoSaveIntervalLabel.Text = $"{prefs.AutoSaveIntervalMinutes} min";
        AutoSaveIntervalPanel.Opacity = prefs.AutoSaveEnabled ? 1.0 : 0.5;
        AutoSaveIntervalSlider.IsEnabled = prefs.AutoSaveEnabled;
        
        // Initialize Python settings (Phase 7)
        PythonBridgeToggle.IsOn = prefs.PythonBridgeEnabled;
        LoadPythonEnvironments();

        // Initialize Python Functions Directory (Phase 7 - Task 21)
        DiscoveredFunctionsList.ItemsSource = _discoveredFunctions;
        _discoveredFunctions.CollectionChanged += DiscoveredFunctions_CollectionChanged;
        InitializePythonFunctionsSection();
        Closed += SettingsDialog_Closed;
    }

    private async void AddService_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new ServiceConnectionDialog();
        dialog.XamlRoot = this.XamlRoot;
        
        var result = await dialog.ShowAsync();
        if (result == ContentDialogResult.Primary && dialog.Connection != null)
        {
            Settings.Connections.Add(dialog.Connection);
        }
    }

    private async void EditService_Click(object sender, RoutedEventArgs e)
    {
        if (ServicesListView.SelectedItem is WorkspaceConnection connection)
        {
            var dialog = new ServiceConnectionDialog(connection);
            dialog.XamlRoot = this.XamlRoot;
            
            var result = await dialog.ShowAsync();
            if (result == ContentDialogResult.Primary && dialog.Connection != null)
            {
                // Update the original connection with the edited values
                connection.Name = dialog.Connection.Name;
                connection.Provider = dialog.Connection.Provider;
                connection.Endpoint = dialog.Connection.Endpoint;
                connection.ApiKey = dialog.Connection.ApiKey;
                connection.Model = dialog.Connection.Model;
                connection.Deployment = dialog.Connection.Deployment;
                connection.IsDefault = dialog.Connection.IsDefault;
                
                // Refresh the list to show updates
                var index = Settings.Connections.IndexOf(connection);
                if (index >= 0)
                {
                    Settings.Connections.RemoveAt(index);
                    Settings.Connections.Insert(index, connection);
                }
            }
        }
    }

    private void RemoveService_Click(object sender, RoutedEventArgs e)
    {
        if (ServicesListView.SelectedItem is WorkspaceConnection connection)
        {
            Settings.Connections.Remove(connection);
        }
    }

    private void SetDefault_Click(object sender, RoutedEventArgs e)
    {
        if (ServicesListView.SelectedItem is WorkspaceConnection connection)
        {
            // Clear all defaults
            foreach (var conn in Settings.Connections)
            {
                conn.IsDefault = false;
            }
            
            // Set selected as default
            connection.IsDefault = true;
            
            // Refresh the list
            var temp = Settings.Connections.ToList();
            Settings.Connections.Clear();
            foreach (var conn in temp)
            {
                Settings.Connections.Add(conn);
            }
        }
    }

    private async void BrowseWorkspace_Click(object sender, RoutedEventArgs e)
    {
        var picker = new FolderPicker();
        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);
        
        picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        picker.FileTypeFilter.Add("*");

        var folder = await picker.PickSingleFolderAsync();
        if (folder != null)
        {
            Settings.WorkspacePath = folder.Path;
        }
    }

    private void MaxThreadsSlider_ValueChanged(object sender, Microsoft.UI.Xaml.Controls.Primitives.RangeBaseValueChangedEventArgs e)
    {
        if (MaxThreadsLabel != null && MaxThreadsDescription != null)
        {
            var value = (int)e.NewValue;
            MaxThreadsLabel.Text = value.ToString();
            var cpuCores = Environment.ProcessorCount;
            MaxThreadsDescription.Text = $"Using {value} threads (CPU cores detected: {cpuCores})";
        }
    }

    private void ThemeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ThemeComboBox.SelectedIndex >= 0)
        {
            var theme = (CellVisualTheme)ThemeComboBox.SelectedIndex;
            Settings.SelectedTheme = theme;
            UpdateThemePreview(theme);
            
            // Apply theme to the application resources
            App.ApplyCellStateTheme(theme);
            
            // Force immediate refresh of the main window grid
            var mainWindow = App.MainWindow?.Content as MainWindow;
            if (mainWindow?.ViewModel?.SelectedSheet != null)
            {
                // Rebuild grid immediately to show theme changes
                mainWindow.BuildSpreadsheetGrid(mainWindow.ViewModel.SelectedSheet);
            }
        }
    }

    private void AppThemeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (AppThemeComboBox.SelectedIndex >= 0)
        {
            var theme = (AppTheme)AppThemeComboBox.SelectedIndex;
            Settings.ApplicationTheme = theme;
            
            // Save to preferences
            var prefs = App.PreferencesService.LoadPreferences();
            prefs.Theme = theme switch
            {
                AppTheme.Light => "Light",
                AppTheme.Dark => "Dark",
                _ => "System"
            };
            App.PreferencesService.SavePreferences(prefs);
            
            // Apply the theme globally
            App.ApplyApplicationTheme(theme);
        }
    }

    private void UpdateThemePreview(CellVisualTheme theme)
    {
        if (PreviewJustUpdated == null) return; // Not yet initialized

        Color justUpdated, calculating, stale, manualUpdate, error, dependency;

        switch (theme)
        {
            case CellVisualTheme.Light:
                justUpdated = Color.FromArgb(0xFF, 0x32, 0xCD, 0x32);    // LimeGreen
                calculating = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);    // Orange
                stale = Color.FromArgb(0xFF, 0x1E, 0x90, 0xFF);          // DodgerBlue
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);   // Orange
                error = Color.FromArgb(0xFF, 0xDC, 0x14, 0x3C);          // Crimson
                dependency = Color.FromArgb(0xFF, 0xFF, 0xD7, 0x00);     // Gold
                break;

            case CellVisualTheme.Dark:
                justUpdated = Color.FromArgb(0xFF, 0x00, 0xFF, 0x7F);    // SpringGreen
                calculating = Color.FromArgb(0xFF, 0xFF, 0x8C, 0x00);    // DarkOrange
                stale = Color.FromArgb(0xFF, 0x87, 0xCE, 0xEB);          // SkyBlue
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0x8C, 0x00);   // DarkOrange
                error = Color.FromArgb(0xFF, 0xFF, 0x44, 0x44);          // Bright Red
                dependency = Color.FromArgb(0xFF, 0xFF, 0xD7, 0x00);     // Gold
                break;

            case CellVisualTheme.HighContrast:
                justUpdated = Color.FromArgb(0xFF, 0x00, 0xFF, 0x00);    // Lime
                calculating = Color.FromArgb(0xFF, 0xFF, 0x66, 0x00);    // Bright Orange
                stale = Color.FromArgb(0xFF, 0x00, 0xFF, 0xFF);          // Cyan
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0x66, 0x00);   // Bright Orange
                error = Color.FromArgb(0xFF, 0xFF, 0x00, 0x00);          // Pure Red
                dependency = Color.FromArgb(0xFF, 0xFF, 0xFF, 0x00);     // Yellow
                break;

            case CellVisualTheme.Custom:
            default:
                justUpdated = Color.FromArgb(0xFF, 0x32, 0xCD, 0x32);    // Default to Light
                calculating = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);
                stale = Color.FromArgb(0xFF, 0x1E, 0x90, 0xFF);
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);
                error = Color.FromArgb(0xFF, 0xDC, 0x14, 0x3C);
                dependency = Color.FromArgb(0xFF, 0xFF, 0xD7, 0x00);
                break;
        }

        PreviewJustUpdated.Background = new SolidColorBrush(justUpdated);
        PreviewCalculating.Background = new SolidColorBrush(calculating);
        PreviewStale.Background = new SolidColorBrush(stale);
        PreviewManualUpdate.Background = new SolidColorBrush(manualUpdate);
        PreviewError.Background = new SolidColorBrush(error);
        PreviewDependency.Background = new SolidColorBrush(dependency);
    }

    private void AutoSaveToggle_Toggled(object sender, RoutedEventArgs e)
    {
        var isEnabled = AutoSaveToggle.IsOn;
        
        // Update UI
        AutoSaveIntervalPanel.Opacity = isEnabled ? 1.0 : 0.5;
        AutoSaveIntervalSlider.IsEnabled = isEnabled;
        
        // Save preference
        var prefs = App.PreferencesService.LoadPreferences();
        prefs.AutoSaveEnabled = isEnabled;
        App.PreferencesService.SavePreferences(prefs);
        
        // Apply to workbook's autosave service
        var mainWindow = App.MainWindow?.Content as MainWindow;
        if (mainWindow?.ViewModel != null)
        {
            mainWindow.ViewModel.SetAutoSaveEnabled(isEnabled);
        }
    }

    private void AutoSaveIntervalSlider_ValueChanged(object sender, Microsoft.UI.Xaml.Controls.Primitives.RangeBaseValueChangedEventArgs e)
    {
        if (AutoSaveIntervalLabel != null)
        {
            var value = (int)e.NewValue;
            AutoSaveIntervalLabel.Text = $"{value} min";
            
            // Save preference
            var prefs = App.PreferencesService.LoadPreferences();
            prefs.AutoSaveIntervalMinutes = value;
            App.PreferencesService.SavePreferences(prefs);
            
            // Apply to workbook's autosave service
            var mainWindow = App.MainWindow?.Content as MainWindow;
            if (mainWindow?.ViewModel != null)
            {
                mainWindow.ViewModel.SetAutoSaveInterval(value);
            }
        }
    }

    // Python Environment Management (Phase 7)
    
    private async void LoadPythonEnvironments()
    {
        try
        {
            SdkStatusText.Text = "Detecting Python environments...";
            
            await Task.Run(() =>
            {
                _pythonEnvironments = PythonEnvironmentDetector.DetectEnvironments();
            });
            
            PythonEnvironmentComboBox.ItemsSource = _pythonEnvironments;
            
            if (_pythonEnvironments.Count == 0)
            {
                SdkStatusText.Text = "No Python environments found";
                SdkStatusIcon.Glyph = "\uE783"; // Warning
                SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Orange);
            }
            else
            {
                // Select previously saved environment or first one
                var prefs = App.PreferencesService.LoadPreferences();
                var savedEnv = _pythonEnvironments.FirstOrDefault(e => e.Path == prefs.PythonEnvironmentPath);
                
                if (savedEnv != null)
                {
                    PythonEnvironmentComboBox.SelectedItem = savedEnv;
                }
                else if (_pythonEnvironments.Count > 0)
                {
                    PythonEnvironmentComboBox.SelectedIndex = 0;
                }
                
                SdkStatusText.Text = $"Found {_pythonEnvironments.Count} Python environment(s)";
                SdkStatusIcon.Glyph = "\uE73E"; // CheckMark
                SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Green);
            }
        }
        catch (Exception ex)
        {
            SdkStatusText.Text = $"Error detecting environments: {ex.Message}";
            SdkStatusIcon.Glyph = "\uE783"; // Warning
            SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Red);
        }
    }
    
    private void RefreshPythonEnvironments_Click(object sender, RoutedEventArgs e)
    {
        LoadPythonEnvironments();
    }
    
    private async void PythonEnvironmentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PythonEnvironmentComboBox.SelectedItem is PythonEnvironmentDetector.PythonEnvironment env)
        {
            PythonPathText.Text = $"Path: {env.Path}";
            
            // Save preference
            var prefs = App.PreferencesService.LoadPreferences();
            prefs.PythonEnvironmentPath = env.Path;
            App.PreferencesService.SavePreferences(prefs);
            
            // Check if SDK is installed
            await CheckSdkInstallation(env);
        }
        else
        {
            PythonPathText.Text = "";
        }
    }
    
    private async Task CheckSdkInstallation(PythonEnvironmentDetector.PythonEnvironment env)
    {
        try
        {
            SdkStatusText.Text = "Checking aicalc-sdk installation...";
            
            bool hasSdk = await Task.Run(() => PythonEnvironmentDetector.HasAiCalcSdk(env.Path));
            
            if (hasSdk)
            {
                SdkStatusText.Text = "✓ aicalc-sdk is installed";
                SdkStatusIcon.Glyph = "\uE73E"; // CheckMark
                SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Green);
                InstallSdkButton.Visibility = Visibility.Collapsed;
            }
            else
            {
                SdkStatusText.Text = "aicalc-sdk is not installed";
                SdkStatusIcon.Glyph = "\uE783"; // Warning
                SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Orange);
                InstallSdkButton.Visibility = Visibility.Visible;
            }
        }
        catch (Exception ex)
        {
            SdkStatusText.Text = $"Error checking SDK: {ex.Message}";
            SdkStatusIcon.Glyph = "\uE783"; // Warning
            SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Red);
        }
    }
    
    private async void InstallSdkButton_Click(object sender, RoutedEventArgs e)
    {
        if (PythonEnvironmentComboBox.SelectedItem is not PythonEnvironmentDetector.PythonEnvironment env)
            return;
        
        try
        {
            InstallSdkButton.IsEnabled = false;
            SdkStatusText.Text = "Installing aicalc-sdk...";
            SdkStatusIcon.Glyph = "\uE895"; // Sync
            
            var (success, message) = await Task.Run(() => PythonEnvironmentDetector.InstallAiCalcSdk(env.Path));
            
            if (success)
            {
                SdkStatusText.Text = "✓ aicalc-sdk installed successfully";
                SdkStatusIcon.Glyph = "\uE73E"; // CheckMark
                SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Green);
                InstallSdkButton.Visibility = Visibility.Collapsed;
            }
            else
            {
                SdkStatusText.Text = $"Installation failed: {message}";
                SdkStatusIcon.Glyph = "\uE783"; // Warning
                SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Red);
            }
        }
        catch (Exception ex)
        {
            SdkStatusText.Text = $"Error: {ex.Message}";
            SdkStatusIcon.Glyph = "\uE783"; // Warning
            SdkStatusIcon.Foreground = new SolidColorBrush(Colors.Red);
        }
        finally
        {
            InstallSdkButton.IsEnabled = true;
        }
    }
    
    private async void TestPythonConnection_Click(object sender, RoutedEventArgs e)
    {
        if (PythonEnvironmentComboBox.SelectedItem is not PythonEnvironmentDetector.PythonEnvironment env)
            return;
        
        try
        {
            var testDialog = new ContentDialog
            {
                Title = "Testing Python Connection",
                Content = "Running test script...",
                CloseButtonText = "Cancel",
                XamlRoot = this.XamlRoot
            };
            
            var _ = testDialog.ShowAsync();
            
            // Run test script
            var testScript = @"
from aicalc_sdk import AiCalcClient
c = AiCalcClient()
c.connect()
c.set_value('A1', 'Test from Settings')
value = c.get_value('A1')
c.disconnect()
print(f'Success! A1 = {value}')
";
            
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = env.Path,
                Arguments = $"-c \"{testScript.Replace("\"", "\\\"")}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            
            using var process = System.Diagnostics.Process.Start(startInfo);
            if (process != null)
            {
                await process.WaitForExitAsync();
                
                var output = await process.StandardOutput.ReadToEndAsync();
                var error = await process.StandardError.ReadToEndAsync();
                
                testDialog.Hide();
                
                var resultDialog = new ContentDialog
                {
                    Title = process.ExitCode == 0 ? "Test Successful" : "Test Failed",
                    Content = !string.IsNullOrEmpty(output) ? output : error,
                    CloseButtonText = "OK",
                    XamlRoot = this.XamlRoot
                };
                
                await resultDialog.ShowAsync();
            }
        }
        catch (Exception ex)
        {
            var errorDialog = new ContentDialog
            {
                Title = "Test Error",
                Content = ex.Message,
                CloseButtonText = "OK",
                XamlRoot = this.XamlRoot
            };
            
            await errorDialog.ShowAsync();
        }
    }
    
    private void PythonBridgeToggle_Toggled(object sender, RoutedEventArgs e)
    {
        var prefs = App.PreferencesService.LoadPreferences();
        prefs.PythonBridgeEnabled = PythonBridgeToggle.IsOn;
        App.PreferencesService.SavePreferences(prefs);
    }

    // Python Functions Directory Management (Phase 7 - Task 21)

    private bool _isScanningFunctions;

    private void DiscoveredFunctions_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        UpdateFunctionCountLabel();
    }

    private void InitializePythonFunctionsSection()
    {
        var prefs = App.PreferencesService.LoadPreferences();
        FunctionsDirectoryTextBox.Text = prefs.PythonFunctionsDirectory ?? string.Empty;
        HotReloadToggle.IsOn = prefs.PythonHotReloadEnabled;
        FunctionCountText.Text = string.IsNullOrWhiteSpace(prefs.PythonFunctionsDirectory)
            ? "(No scan yet)"
            : "(Ready to scan)";

        if (!string.IsNullOrWhiteSpace(prefs.PythonFunctionsDirectory) && Directory.Exists(prefs.PythonFunctionsDirectory))
        {
            if (HotReloadToggle.IsOn)
            {
                StartFunctionsWatcher(prefs.PythonFunctionsDirectory);
            }

            _ = ScanFunctionsAsync(showErrors: false);
        }
        else if (!string.IsNullOrWhiteSpace(prefs.PythonFunctionsDirectory))
        {
            FunctionCountText.Text = "(Directory not found)";
        }
    }

    private async void BrowseFunctionsDirectory_Click(object sender, RoutedEventArgs e)
    {
        var picker = new FolderPicker();
        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

        picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        picker.FileTypeFilter.Add("*");

        var folder = await picker.PickSingleFolderAsync();
        if (folder == null)
        {
            return;
        }

        FunctionsDirectoryTextBox.Text = folder.Path;

        var prefs = App.PreferencesService.LoadPreferences();
        prefs.PythonFunctionsDirectory = folder.Path;
        App.PreferencesService.SavePreferences(prefs);

        if (HotReloadToggle.IsOn)
        {
            StartFunctionsWatcher(folder.Path);
        }
        else
        {
            StopFunctionsWatcher();
        }

        await ScanFunctionsAsync(showErrors: true);
    }

    private void HotReloadToggle_Toggled(object sender, RoutedEventArgs e)
    {
        var prefs = App.PreferencesService.LoadPreferences();
        prefs.PythonHotReloadEnabled = HotReloadToggle.IsOn;
        App.PreferencesService.SavePreferences(prefs);

        var directory = FunctionsDirectoryTextBox.Text?.Trim();
        if (HotReloadToggle.IsOn && !string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
        {
            StartFunctionsWatcher(directory);
        }
        else
        {
            StopFunctionsWatcher();
        }
    }

    private async void ScanFunctions_Click(object sender, RoutedEventArgs e)
    {
        await ScanFunctionsAsync(showErrors: true);
    }

    private async Task ScanFunctionsAsync(bool showErrors)
    {
        if (_isScanningFunctions)
        {
            _pendingHotReloadScan = true;
            return;
        }

        var directory = FunctionsDirectoryTextBox.Text?.Trim();
        if (string.IsNullOrWhiteSpace(directory))
        {
            FunctionCountText.Text = "(Select a directory)";
            if (showErrors)
            {
                await ShowMessageAsync("Select Directory", "Choose a folder that contains your Python functions before scanning.");
            }
            return;
        }

        if (!Directory.Exists(directory))
        {
            FunctionCountText.Text = "(Directory not found)";
            StopFunctionsWatcher();
            if (showErrors)
            {
                await ShowMessageAsync("Directory Not Found", "The selected functions directory no longer exists. Choose a different folder.");
            }
            return;
        }

        var pythonPath = GetActivePythonPath();
        if (string.IsNullOrWhiteSpace(pythonPath) || !File.Exists(pythonPath))
        {
            if (showErrors)
            {
                await ShowMessageAsync("Python Environment Required", "Select a Python environment before scanning for custom functions.");
            }
            return;
        }

        _isScanningFunctions = true;
        _pendingHotReloadScan = false;
        ScanFunctionsButton.IsEnabled = false;
        FunctionCountText.Text = "(Scanning...)";

        List<PythonFunctionInfo> results = new();
        Exception? scanError = null;

        try
        {
            var scanner = new PythonFunctionScanner(pythonPath);
            results = await scanner.ScanDirectoryAsync(directory);
        }
        catch (Exception ex)
        {
            scanError = ex;
        }
        finally
        {
            _isScanningFunctions = false;
            ScanFunctionsButton.IsEnabled = true;
        }

        if (scanError != null)
        {
            FunctionCountText.Text = "(Scan failed)";
            if (showErrors)
            {
                await ShowMessageAsync("Scan Failed", scanError.Message);
            }
        }
        else
        {
            UpdateDiscoveredFunctions(results);
            RegisterDiscoveredFunctions(results, pythonPath);
            UpdateFunctionCountLabel();
        }

        if (_pendingHotReloadScan)
        {
            _pendingHotReloadScan = false;
            await ScanFunctionsAsync(showErrors: false);
        }
    }

    private string? GetActivePythonPath()
    {
        if (PythonEnvironmentComboBox.SelectedItem is PythonEnvironmentDetector.PythonEnvironment env)
        {
            return env.Path;
        }

        var prefs = App.PreferencesService.LoadPreferences();
        return prefs.PythonEnvironmentPath;
    }

    private void UpdateDiscoveredFunctions(IEnumerable<PythonFunctionInfo> functions)
    {
        _discoveredFunctions.Clear();

        foreach (var info in functions.OrderBy(f => f.Name, StringComparer.OrdinalIgnoreCase))
        {
            _discoveredFunctions.Add(info);
        }
    }

    private void UpdateFunctionCountLabel()
    {
        if (_isScanningFunctions)
        {
            FunctionCountText.Text = "(Scanning...)";
            return;
        }

        if (_discoveredFunctions.Count == 0)
        {
            FunctionCountText.Text = "(No functions found)";
        }
        else
        {
            FunctionCountText.Text = $"({_discoveredFunctions.Count} found)";
        }
    }

    private void RegisterDiscoveredFunctions(IEnumerable<PythonFunctionInfo> functions, string pythonPath)
    {
        if (App.MainWindow?.Content is not MainWindow mainWindow)
        {
            return;
        }

        var newFunctionNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var info in functions)
        {
            newFunctionNames.Add(info.Name);
        }

        // Remove functions that disappeared since last scan
        foreach (var existing in _registeredPythonFunctionNames.ToArray())
        {
            if (!newFunctionNames.Contains(existing))
            {
                mainWindow.ViewModel.FunctionRegistry.Unregister(existing);
                _registeredPythonFunctionNames.Remove(existing);
            }
        }

        foreach (var info in functions)
        {
            var descriptor = PythonFunctionScanner.CreateDescriptor(info, pythonPath);
            mainWindow.ViewModel.FunctionRegistry.Register(descriptor);
            _registeredPythonFunctionNames.Add(info.Name);
        }

        mainWindow.DispatcherQueue?.TryEnqueue(mainWindow.RefreshFunctionCatalog);
    }

    private void StartFunctionsWatcher(string directory)
    {
        StopFunctionsWatcher();

        if (!HotReloadToggle.IsOn)
        {
            return;
        }

        try
        {
            _functionsWatcher = new FileSystemWatcher(directory)
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.DirectoryName,
                Filter = "*.py",
                IncludeSubdirectories = true,
                EnableRaisingEvents = true
            };

            _functionsWatcher.Created += OnFunctionsDirectoryChanged;
            _functionsWatcher.Changed += OnFunctionsDirectoryChanged;
            _functionsWatcher.Deleted += OnFunctionsDirectoryChanged;
            _functionsWatcher.Renamed += OnFunctionsDirectoryRenamed;
        }
        catch
        {
            StopFunctionsWatcher();
        }
    }

    private void StopFunctionsWatcher()
    {
        if (_functionsWatcher == null)
        {
            return;
        }

        _functionsWatcher.EnableRaisingEvents = false;
        _functionsWatcher.Created -= OnFunctionsDirectoryChanged;
        _functionsWatcher.Changed -= OnFunctionsDirectoryChanged;
        _functionsWatcher.Deleted -= OnFunctionsDirectoryChanged;
        _functionsWatcher.Renamed -= OnFunctionsDirectoryRenamed;
        _functionsWatcher.Dispose();
        _functionsWatcher = null;
    }

    private void OnFunctionsDirectoryChanged(object sender, FileSystemEventArgs e)
    {
        ScheduleHotReloadScan();
    }

    private void OnFunctionsDirectoryRenamed(object sender, RenamedEventArgs e)
    {
        ScheduleHotReloadScan();
    }

    private void ScheduleHotReloadScan()
    {
        if (!HotReloadToggle.IsOn)
        {
            return;
        }

        DispatcherQueue?.TryEnqueue(async () => await ScanFunctionsAsync(showErrors: false));
    }

    private async void OpenFunctionInVsCode_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not string filePath || string.IsNullOrWhiteSpace(filePath))
        {
            return;
        }

        if (!File.Exists(filePath))
        {
            await ShowMessageAsync("File Not Found", $"The file \"{filePath}\" could not be located.");
            return;
        }

        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "code",
                Arguments = $"\"{filePath}\"",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            System.Diagnostics.Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            await ShowMessageAsync("Open in VS Code Failed", ex.Message);
        }
    }

    private async Task ShowMessageAsync(string title, string message)
    {
        var dialog = new ContentDialog
        {
            Title = title,
            Content = message,
            CloseButtonText = "OK",
            XamlRoot = XamlRoot
        };

        await dialog.ShowAsync();
    }

    private void SettingsDialog_Closed(ContentDialog sender, ContentDialogClosedEventArgs args)
    {
        _discoveredFunctions.CollectionChanged -= DiscoveredFunctions_CollectionChanged;
        StopFunctionsWatcher();
    }
}
