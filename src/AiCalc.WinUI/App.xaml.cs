using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
using AiCalc.Models;
using Windows.UI;

namespace AiCalc;

public partial class App : Application
{
    private Window? _window;
    
    public static Window? MainWindow { get; private set; }

    public App()
    {
        InitializeComponent();
        this.UnhandledException += OnUnhandledException;
        
        // Load default theme (Light) on startup - Task 10
        ApplyCellStateTheme(CellVisualTheme.Light);
    }

    private void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"Unhandled exception: {e.Exception}" );
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        base.OnLaunched(args);
        if (_window is null)
        {
            _window = new Window
            {
                Title = "AiCalc Studio"
            };
            _window.Content = new MainWindow();
            MainWindow = _window;
        }

        _window.Activate();
    }

    /// <summary>
    /// Applies the selected cell state theme (Task 10)
    /// </summary>
    public static void ApplyCellStateTheme(CellVisualTheme theme)
    {
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
                // Custom theme uses defaults (same as Light for now)
                justUpdated = Color.FromArgb(0xFF, 0x32, 0xCD, 0x32);
                calculating = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);
                stale = Color.FromArgb(0xFF, 0x1E, 0x90, 0xFF);
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0xA5, 0x00);
                error = Color.FromArgb(0xFF, 0xDC, 0x14, 0x3C);
                dependency = Color.FromArgb(0xFF, 0xFF, 0xD7, 0x00);
                break;
        }

        // Update app resources
        Current.Resources["CellStateJustUpdatedBrush"] = new SolidColorBrush(justUpdated);
        Current.Resources["CellStateCalculatingBrush"] = new SolidColorBrush(calculating);
        Current.Resources["CellStateStaleBrush"] = new SolidColorBrush(stale);
        Current.Resources["CellStateManualUpdateBrush"] = new SolidColorBrush(manualUpdate);
        Current.Resources["CellStateErrorBrush"] = new SolidColorBrush(error);
        Current.Resources["CellStateInDependencyChainBrush"] = new SolidColorBrush(dependency);
    }
}
