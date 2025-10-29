using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Windowing;
using Microsoft.UI;
using System;
using Windows.Foundation;
using WinRT.Interop;

namespace AiCalc;

/// <summary>
/// Base class for custom dialog windows that provides consistent sizing, styling, and modal behavior.
/// Avoids ContentDialog's sizing constraints and provides full control over layout.
/// </summary>
public abstract class DialogWindowBase : Window
{
    private Grid? _rootGrid;
    private Border? _contentBorder;
    private StackPanel? _buttonPanel;
    private Button? _primaryButton;
    private Button? _closeButton;
    
    protected ContentPresenter? DialogContent;
    
    public new event TypedEventHandler<DialogWindowBase, DialogResult>? Closed;
    
    public enum DialogResult
    {
        None,
        Primary,
        Cancel
    }
    
    /// <summary>
    /// Gets or sets the dialog title.
    /// </summary>
    public string DialogTitle { get; set; } = string.Empty;
    
    /// <summary>
    /// Gets or sets the primary button text (e.g., "Save", "OK").
    /// </summary>
    public string PrimaryButtonText { get; set; } = "OK";
    
    /// <summary>
    /// Gets or sets the close button text (e.g., "Cancel", "Close").
    /// </summary>
    public string CloseButtonText { get; set; } = "Cancel";
    
    /// <summary>
    /// Gets or sets the desired dialog width. Default is 800.
    /// </summary>
    public double DialogWidth { get; set; } = 800;
    
    /// <summary>
    /// Gets or sets the desired dialog height. Default is 600.
    /// </summary>
    public double DialogHeight { get; set; } = 600;
    
    /// <summary>
    /// Gets or sets whether the primary button is enabled.
    /// </summary>
    public bool IsPrimaryButtonEnabled { get; set; } = true;
    
    public DialogResult Result { get; private set; } = DialogResult.None;
    
    protected DialogWindowBase()
    {
        InitializeWindow();
        BuildUI();
    }
    
    private void InitializeWindow()
    {
        // Set up the window chrome and sizing
        var hWnd = WindowNative.GetWindowHandle(this);
        var windowId = Win32Interop.GetWindowIdFromWindow(hWnd);
        var appWindow = AppWindow.GetFromWindowId(windowId);
        
        // Remove default title bar for custom styling
        if (appWindow.TitleBar is { } titleBar)
        {
            titleBar.ExtendsContentIntoTitleBar = true;
            titleBar.ButtonBackgroundColor = Colors.Transparent;
            titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;
        }
        
        // Set window size
        appWindow.Resize(new Windows.Graphics.SizeInt32 
        { 
            Width = (int)DialogWidth, 
            Height = (int)DialogHeight 
        });
        
        // Center the window on the screen
        CenterWindow(appWindow);
    }
    
    private void CenterWindow(AppWindow appWindow)
    {
        var displayArea = DisplayArea.GetFromWindowId(appWindow.Id, DisplayAreaFallback.Primary);
        if (displayArea is not null)
        {
            var centerX = (displayArea.WorkArea.Width - DialogWidth) / 2;
            var centerY = (displayArea.WorkArea.Height - DialogHeight) / 2;
            
            appWindow.Move(new Windows.Graphics.PointInt32 
            { 
                X = (int)centerX + displayArea.WorkArea.X, 
                Y = (int)centerY + displayArea.WorkArea.Y 
            });
        }
    }
    
    private void BuildUI()
    {
        _rootGrid = new Grid
        {
            Background = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["LayerFillColorDefaultBrush"]
        };
        
        _rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Auto) }); // Title
        _rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
        _rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Auto) }); // Buttons
        
        // Title bar
        var titleBar = new Border
        {
            Padding = new Thickness(20, 12, 20, 12),
            Background = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["LayerFillColorDefaultBrush"],
            BorderBrush = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["CardStrokeColorDefaultBrush"],
            BorderThickness = new Thickness(0, 0, 0, 1)
        };
        
        var titleText = new TextBlock
        {
            Text = DialogTitle,
            FontSize = 20,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        };
        
        titleBar.Child = titleText;
        Grid.SetRow(titleBar, 0);
        _rootGrid.Children.Add(titleBar);
        
        // Content area with ScrollViewer
        var scrollViewer = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Padding = new Thickness(20)
        };
        
        DialogContent = new ContentPresenter();
        scrollViewer.Content = DialogContent;
        
        _contentBorder = new Border
        {
            Child = scrollViewer
        };
        
        Grid.SetRow(_contentBorder, 1);
        _rootGrid.Children.Add(_contentBorder);
        
        // Button panel
        _buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing = 8,
            Padding = new Thickness(20, 12, 20, 16),
            Background = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["LayerFillColorDefaultBrush"],
            BorderBrush = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["CardStrokeColorDefaultBrush"],
            BorderThickness = new Thickness(0, 1, 0, 0)
        };
        
        _primaryButton = new Button
        {
            Content = PrimaryButtonText,
            Style = (Style)Application.Current.Resources["AccentButtonStyle"],
            MinWidth = 120,
            IsEnabled = IsPrimaryButtonEnabled
        };
        _primaryButton.Click += PrimaryButton_Click;
        
        _closeButton = new Button
        {
            Content = CloseButtonText,
            MinWidth = 120
        };
        _closeButton.Click += CloseButton_Click;
        
        _buttonPanel.Children.Add(_primaryButton);
        _buttonPanel.Children.Add(_closeButton);
        
        Grid.SetRow(_buttonPanel, 2);
        _rootGrid.Children.Add(_buttonPanel);
        
        Content = _rootGrid;
    }
    
    /// <summary>
    /// Called when the primary button is clicked. Override to add validation logic.
    /// Return true to close the dialog, false to keep it open.
    /// </summary>
    protected virtual bool OnPrimaryButtonClick()
    {
        return true;
    }
    
    /// <summary>
    /// Called when the close/cancel button is clicked. Override to add custom logic.
    /// Return true to close the dialog, false to keep it open.
    /// </summary>
    protected virtual bool OnCloseButtonClick()
    {
        return true;
    }
    
    private void PrimaryButton_Click(object sender, RoutedEventArgs e)
    {
        if (OnPrimaryButtonClick())
        {
            Result = DialogResult.Primary;
            Closed?.Invoke(this, Result);
            Close();
        }
    }
    
    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        if (OnCloseButtonClick())
        {
            Result = DialogResult.Cancel;
            Closed?.Invoke(this, Result);
            Close();
        }
    }
    
    /// <summary>
    /// Sets the content of the dialog. Call this from your derived class constructor.
    /// </summary>
    protected void SetDialogContent(UIElement content)
    {
        if (DialogContent is not null)
        {
            DialogContent.Content = content;
        }
    }
    
    /// <summary>
    /// Updates the primary button's enabled state.
    /// </summary>
    protected void UpdatePrimaryButtonState(bool isEnabled)
    {
        if (_primaryButton is not null)
        {
            _primaryButton.IsEnabled = isEnabled;
        }
    }
}
