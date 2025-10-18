using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
using AiCalc.Models;
using AiCalc.Services;
using AiCalc.Services.AI;
using System;
using Windows.UI;

namespace AiCalc;

public partial class App : Application
{
    private Window? _window;
    
    public static Window? MainWindow { get; private set; }
    
    /// <summary>
    /// Global AI Service Registry for managing AI connections
    /// </summary>
    public static AIServiceRegistry AIServices { get; private set; } = new AIServiceRegistry();

    /// <summary>
    /// User preferences service (Phase 5)
    /// </summary>
    public static UserPreferencesService PreferencesService { get; private set; } = new UserPreferencesService();

    public App()
    {
        try
        {
            InitializeComponent();
            this.UnhandledException += OnUnhandledException;
            
            // NOTE: Theme is applied in OnLaunched after window is created
            // Calling ApplyCellStateTheme here causes crash because Resources aren't ready yet
        }
        catch (Exception ex)
        {
            System.IO.File.WriteAllText(@"C:\Projects\aicalc\app_crash.log", 
                $"App Constructor Exception:\n{ex.Message}\n\n{ex.StackTrace}");
            throw;
        }
    }

    private void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"Unhandled exception: {e.Exception}");
        System.IO.File.WriteAllText(@"C:\Projects\aicalc\app_unhandled.log", 
            $"Unhandled Exception:\n{e.Exception.Message}\n\n{e.Exception.StackTrace}");
        e.Handled = true;
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        try
        {
            base.OnLaunched(args);
            
            // Load user preferences (Phase 5)
            var prefs = PreferencesService.LoadPreferences();
            
            // Apply cell state theme
            ApplyCellStateTheme(CellVisualTheme.Light);
            
            if (_window is null)
            {
                _window = new Window
                {
                    Title = "AiCalc Studio"
                };
                _window.Content = new MainWindow();
                MainWindow = _window;
                
                // Apply saved application theme
                var appTheme = prefs.Theme switch
                {
                    "Light" => AppTheme.Light,
                    "Dark" => AppTheme.Dark,
                    _ => AppTheme.System
                };
                ApplyApplicationTheme(appTheme);
                
                // Restore window size
                if (prefs.WindowWidth > 0 && prefs.WindowHeight > 0)
                {
                    var appWindow = _window.AppWindow;
                    appWindow.Resize(new Windows.Graphics.SizeInt32((int)prefs.WindowWidth, (int)prefs.WindowHeight));
                }
                
                // Handle window closing to save preferences
                _window.Closed += OnWindowClosed;
            }

            _window.Activate();
        }
        catch (Exception ex)
        {
            System.IO.File.WriteAllText(@"C:\Projects\aicalc\onlaunched_crash.log", 
                $"OnLaunched Exception:\n{ex.Message}\n\n{ex.StackTrace}\n\nInner: {ex.InnerException?.Message}\n{ex.InnerException?.StackTrace}");
            throw;
        }
    }

    private void OnWindowClosed(object sender, WindowEventArgs args)
    {
        try
        {
            // Save preferences on close (Phase 5)
            if (_window?.Content is MainWindow mainWindow)
            {
                mainWindow.SavePreferences();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving preferences on close: {ex.Message}");
        }
    }

    /// <summary>
    /// Applies the selected cell state theme (Task 10)
    /// </summary>
    public static void ApplyCellStateTheme(CellVisualTheme theme)
    {
        Color justUpdated, calculating, stale, manualUpdate, error, dependency;
        Color cellBackground;
        Color cellForeground;
        Color cellBorder;

        switch (theme)
        {
            case CellVisualTheme.Light:
                justUpdated = Color.FromArgb(0xFF, 0x32, 0xCD, 0x32);    // LimeGreen
                calculating = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);    // Orange
                stale = Color.FromArgb(0xFF, 0x1E, 0x90, 0xFF);          // DodgerBlue
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);   // Orange
                error = Color.FromArgb(0xFF, 0xDC, 0x14, 0x3C);          // Crimson
                dependency = Color.FromArgb(0xFF, 0xFF, 0xD7, 0x00);     // Gold
                cellBackground = Color.FromArgb(0xFF, 0xF7, 0xFA, 0xFF);
                cellForeground = Color.FromArgb(0xFF, 0x20, 0x2A, 0x36);
                cellBorder = Color.FromArgb(0xFF, 0xC7, 0xD3, 0xE3);
                break;

            case CellVisualTheme.Dark:
                justUpdated = Color.FromArgb(0xFF, 0x00, 0xFF, 0x7F);    // SpringGreen
                calculating = Color.FromArgb(0xFF, 0xFF, 0x8C, 0x00);    // DarkOrange
                stale = Color.FromArgb(0xFF, 0x87, 0xCE, 0xEB);          // SkyBlue
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0x8C, 0x00);   // DarkOrange
                error = Color.FromArgb(0xFF, 0xFF, 0x44, 0x44);          // Bright Red
                dependency = Color.FromArgb(0xFF, 0xFF, 0xD7, 0x00);     // Gold
                cellBackground = Color.FromArgb(0xFF, 0x23, 0x28, 0x31);
                cellForeground = Color.FromArgb(0xFF, 0xF5, 0xF7, 0xFA);
                cellBorder = Color.FromArgb(0xFF, 0x43, 0x4B, 0x55);
                break;

            case CellVisualTheme.HighContrast:
                justUpdated = Color.FromArgb(0xFF, 0x00, 0xFF, 0x00);    // Lime
                calculating = Color.FromArgb(0xFF, 0xFF, 0x66, 0x00);    // Bright Orange
                stale = Color.FromArgb(0xFF, 0x00, 0xFF, 0xFF);          // Cyan
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0x66, 0x00);   // Bright Orange
                error = Color.FromArgb(0xFF, 0xFF, 0x00, 0x00);          // Pure Red
                dependency = Color.FromArgb(0xFF, 0xFF, 0xFF, 0x00);     // Yellow
                cellBackground = Color.FromArgb(0xFF, 0x00, 0x00, 0x00);
                cellForeground = Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF);
                cellBorder = Color.FromArgb(0xFF, 0xFF, 0xFF, 0x00);
                break;

            case CellVisualTheme.Custom:
            default:
                // Custom theme uses defaults (same as Light for now)
                justUpdated = Color.FromArgb(0xFF, 0x32, 0xCD, 0x32);
                calculating = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);
                stale = Color.FromArgb(0xFF, 0x1E, 0x90, 0xFF);
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);
                error = Color.FromArgb(0xFF, 0xDC, 0x14, 0x3C);
                dependency = Color.FromArgb(0xFF, 0xFF, 0xD7, 0x00);
                cellBackground = Color.FromArgb(0xFF, 0xF7, 0xFA, 0xFF);
                cellForeground = Color.FromArgb(0xFF, 0x20, 0x2A, 0x36);
                cellBorder = Color.FromArgb(0xFF, 0xC7, 0xD3, 0xE3);
                break;
        }

        // Update app resources
        Current.Resources["CellStateJustUpdatedBrush"] = new SolidColorBrush(justUpdated);
        Current.Resources["CellStateCalculatingBrush"] = new SolidColorBrush(calculating);
        Current.Resources["CellStateStaleBrush"] = new SolidColorBrush(stale);
        Current.Resources["CellStateManualUpdateBrush"] = new SolidColorBrush(manualUpdate);
        Current.Resources["CellStateErrorBrush"] = new SolidColorBrush(error);
        Current.Resources["CellStateInDependencyChainBrush"] = new SolidColorBrush(dependency);
        Current.Resources["CellStateNormalBrush"] = new SolidColorBrush(cellBackground);
        Current.Resources["CellThemeBackgroundBrush"] = new SolidColorBrush(cellBackground);
        Current.Resources["CellThemeForegroundBrush"] = new SolidColorBrush(cellForeground);
        Current.Resources["CellThemeBorderBrush"] = new SolidColorBrush(cellBorder);
    }

    /// <summary>
    /// Applies the selected application theme (Task 10)
    /// </summary>
    public static void ApplyApplicationTheme(AppTheme theme)
    {
        if (MainWindow?.Content is FrameworkElement rootElement)
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
}
