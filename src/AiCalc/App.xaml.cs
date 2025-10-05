using Microsoft.UI.Xaml;

namespace AiCalc;

public partial class App : Application
{
    private Window? _window;

    public App()
    {
        InitializeComponent();
        this.UnhandledException += OnUnhandledException;
    }

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
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
            _window.Content = new MainPage();
        }

        _window.Activate();
    }
}
