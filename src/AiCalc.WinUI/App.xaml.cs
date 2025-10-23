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
#if DEBUG
            // Initialize debug console for debug output
            DebugConsole.EnsureInitialized();
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] AiCalc starting...");
#endif
            
            base.OnLaunched(args);
            
            // Load user preferences (Phase 5)
            var prefs = PreferencesService.LoadPreferences();
            
            // Ensure theme is set to Dark if it's somehow empty or System
            if (string.IsNullOrEmpty(prefs.Theme) || prefs.Theme == "System")
            {
                prefs.Theme = "Dark";
                PreferencesService.SavePreferences(prefs);
            }
            
            // Apply saved application theme FIRST (before creating MainWindow)
            var appTheme = prefs.Theme switch
            {
                "Light" => AppTheme.Light,
                "Dark" => AppTheme.Dark,
                _ => AppTheme.Dark  // Default to Dark theme
            };
            ApplyApplicationTheme(appTheme);
            
            // Also update cell theme colors to match application theme
            if (Current.Resources.TryGetValue("CellThemeBackgroundBrush", out var cellBgBrush) && cellBgBrush is SolidColorBrush)
            {
                Current.Resources["CellThemeBackgroundBrush"] = new SolidColorBrush(appTheme == AppTheme.Dark ? Color.FromArgb(0xFF, 0x25, 0x25, 0x26) : Color.FromArgb(0xFF, 0xF8, 0xF8, 0xF8));
                Current.Resources["CellThemeForegroundBrush"] = new SolidColorBrush(appTheme == AppTheme.Dark ? Color.FromArgb(0xFF, 0xCC, 0xCC, 0xCC) : Color.FromArgb(0xFF, 0x1E, 0x1E, 0x1E));
                Current.Resources["CellThemeBorderBrush"] = new SolidColorBrush(appTheme == AppTheme.Dark ? Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30) : Color.FromArgb(0xFF, 0xD0, 0xD0, 0xD0));
            }
            
            // Apply cell state theme to match application theme
            var defaultCellTheme = appTheme == AppTheme.Dark ? CellVisualTheme.Dark : CellVisualTheme.Light;
            ApplyCellStateTheme(defaultCellTheme);
            
            if (_window is null)
            {
                _window = new Window
                {
                    Title = "AiCalc Studio"
                };
                _window.Content = new MainWindow();
                MainWindow = _window;
                
                // Always size to 80% of the primary screen and center it
                var appWindow = _window.AppWindow;
                var displayArea = Microsoft.UI.Windowing.DisplayArea.GetFromWindowId(appWindow.Id, Microsoft.UI.Windowing.DisplayAreaFallback.Primary);
                if (displayArea != null)
                {
                    var workArea = displayArea.WorkArea;

                    int targetWidth = (int)Math.Round(workArea.Width * 0.8);
                    int targetHeight = (int)Math.Round(workArea.Height * 0.8);

                    targetWidth = Math.Max(targetWidth, 800);
                    targetHeight = Math.Max(targetHeight, 600);

                    int x = workArea.X + (workArea.Width - targetWidth) / 2;
                    int y = workArea.Y + (workArea.Height - targetHeight) / 2;

                    var centeringRect = new Windows.Graphics.RectInt32(x, y, targetWidth, targetHeight);
                    appWindow.MoveAndResize(centeringRect);
                }
                else
                {
                    var fallbackRect = new Windows.Graphics.RectInt32(0, 0, 1400, 900);
                    appWindow.MoveAndResize(fallbackRect);
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
                cellBackground = Color.FromArgb(0xFF, 0xF8, 0xF8, 0xF8); // Light gray matching mockup
                cellForeground = Color.FromArgb(0xFF, 0x1E, 0x1E, 0x1E);
                cellBorder = Color.FromArgb(0xFF, 0xD0, 0xD0, 0xD0);
                break;

            case CellVisualTheme.Dark:
                justUpdated = Color.FromArgb(0xFF, 0x00, 0xFF, 0x7F);    // SpringGreen
                calculating = Color.FromArgb(0xFF, 0xFF, 0x8C, 0x00);    // DarkOrange
                stale = Color.FromArgb(0xFF, 0x87, 0xCE, 0xEB);          // SkyBlue
                manualUpdate = Color.FromArgb(0xFF, 0xFF, 0x8C, 0x00);   // DarkOrange
                error = Color.FromArgb(0xFF, 0xFF, 0x44, 0x44);          // Bright Red
                dependency = Color.FromArgb(0xFF, 0xFF, 0xD7, 0x00);     // Gold
                cellBackground = Color.FromArgb(0xFF, 0x25, 0x25, 0x26); // Darker to match theme
                cellForeground = Color.FromArgb(0xFF, 0xCC, 0xCC, 0xCC);
                cellBorder = Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
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

    // Update app resources in-place so theme swaps propagate immediately
    var resources = Current.Resources;
    UpdateResourceBrush(resources, "CellStateJustUpdatedBrush", justUpdated);
    UpdateResourceBrush(resources, "CellStateCalculatingBrush", calculating);
    UpdateResourceBrush(resources, "CellStateStaleBrush", stale);
    UpdateResourceBrush(resources, "CellStateManualUpdateBrush", manualUpdate);
    UpdateResourceBrush(resources, "CellStateErrorBrush", error);
    UpdateResourceBrush(resources, "CellStateInDependencyChainBrush", dependency);
    UpdateResourceBrush(resources, "CellStateNormalBrush", cellBackground);
    UpdateResourceBrush(resources, "CellThemeBackgroundBrush", cellBackground);
    UpdateResourceBrush(resources, "CellThemeForegroundBrush", cellForeground);
    UpdateResourceBrush(resources, "CellThemeBorderBrush", cellBorder);
    }

    /// <summary>
    /// Applies the selected application theme - professional VS Code inspired dark theme
    /// </summary>
    public static void ApplyApplicationTheme(AppTheme theme)
    {
        // Swap resource brushes based on application theme
        bool useLight = theme == AppTheme.Light;

        var resources = Current.Resources;
        
        // Define color palettes - professional dark theme matching mockup
        Color appBg = useLight ? Color.FromArgb(0xFF, 0xE8, 0xE8, 0xE8) : Color.FromArgb(0xFF, 0x1E, 0x1E, 0x1E);
        Color cardBg = useLight ? Color.FromArgb(0xFF, 0xF5, 0xF5, 0xF5) : Color.FromArgb(0xFF, 0x25, 0x25, 0x26);
        Color accent = useLight ? Color.FromArgb(0xFF, 0x00, 0x78, 0xD4) : Color.FromArgb(0xFF, 0x0E, 0x63, 0x9C);
        Color accentAlt = useLight ? Color.FromArgb(0xFF, 0x10, 0x7C, 0x10) : Color.FromArgb(0xFF, 0x00, 0xA8, 0x76);
        Color textPrimary = useLight ? Color.FromArgb(0xFF, 0x1E, 0x1E, 0x1E) : Color.FromArgb(0xFF, 0xCC, 0xCC, 0xCC);
        Color textSecondary = useLight ? Color.FromArgb(0xFF, 0x61, 0x61, 0x61) : Color.FromArgb(0xFF, 0x80, 0x80, 0x80);
        Color border = useLight ? Color.FromArgb(0xFF, 0xD0, 0xD0, 0xD0) : Color.FromArgb(0xFF, 0x3E, 0x3E, 0x42);
        Color gridLine = useLight ? Color.FromArgb(0xFF, 0xD0, 0xD0, 0xD0) : Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
        Color headerBg = useLight ? Color.FromArgb(0xFF, 0xE0, 0xE0, 0xE0) : Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
        Color cellBg = useLight ? Color.FromArgb(0xFF, 0xF8, 0xF8, 0xF8) : Color.FromArgb(0xFF, 0x25, 0x25, 0x26);
        
        // Update resource brushes
    UpdateResourceBrush(resources, "AppBackgroundBrush", appBg);
    UpdateResourceBrush(resources, "CardBackgroundBrush", cardBg);
    UpdateResourceBrush(resources, "AccentBrush", accent);
    UpdateResourceBrush(resources, "AccentBrushAlt", accentAlt);
    UpdateResourceBrush(resources, "TextPrimaryBrush", textPrimary);
    UpdateResourceBrush(resources, "TextSecondaryBrush", textSecondary);
    UpdateResourceBrush(resources, "BorderBrushColor", border);
    UpdateResourceBrush(resources, "GridLineColor", gridLine);
    UpdateResourceBrush(resources, "HeaderBackgroundBrush", headerBg);
    UpdateResourceBrush(resources, "CellBackgroundBrush", cellBg);

    // Also update cell theme brushes to match application theme
    UpdateResourceBrush(resources, "CellThemeBackgroundBrush", cellBg);
    UpdateResourceBrush(resources, "CellThemeForegroundBrush", textPrimary);
    UpdateResourceBrush(resources, "CellThemeBorderBrush", gridLine);
        
        // Also update the root element's RequestedTheme for WinUI controls
        if (MainWindow?.Content is FrameworkElement rootElement)
        {
            rootElement.RequestedTheme = theme switch
            {
                AppTheme.Light => ElementTheme.Light,
                AppTheme.Dark => ElementTheme.Dark,
                AppTheme.System => ElementTheme.Default,
                _ => ElementTheme.Dark  // Default to dark
            };
        }
    }

    private static void UpdateResourceBrush(ResourceDictionary resources, string key, Color color)
    {
        if (resources.TryGetValue(key, out var existing) && existing is SolidColorBrush brush)
        {
            brush.Color = color;
        }
        else
        {
            resources[key] = new SolidColorBrush(color);
        }
    }
}
