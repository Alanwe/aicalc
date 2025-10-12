using AiCalc.Models;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using System;
using System.Linq;
using Windows.Storage.Pickers;
using Windows.UI;

namespace AiCalc;

public sealed partial class SettingsDialog : ContentDialog
{
    public WorkbookSettings Settings { get; }

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
            
            // Apply theme to the application
            App.ApplyCellStateTheme(theme);
        }
    }

    private void AppThemeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (AppThemeComboBox.SelectedIndex >= 0)
        {
            var theme = (AppTheme)AppThemeComboBox.SelectedIndex;
            Settings.ApplicationTheme = theme;
            ApplyApplicationTheme(theme);
        }
    }

    private void ApplyApplicationTheme(AppTheme theme)
    {
        var window = App.MainWindow;
        if (window?.Content is FrameworkElement rootElement)
        {
            rootElement.RequestedTheme = theme switch
            {
                AppTheme.Light => ElementTheme.Light,
                AppTheme.Dark => ElementTheme.Dark,
                AppTheme.System => ElementTheme.Default,
                _ => ElementTheme.Default
            };
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
}
