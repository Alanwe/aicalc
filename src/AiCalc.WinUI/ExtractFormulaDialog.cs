using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace AiCalc;

public sealed class ExtractFormulaDialog : ContentDialog
{
    private readonly TextBox _targetTextBox;
    private readonly TextBlock _validationText;

    public string SourceAddress { get; }
    public string TargetAddress { get; private set; } = string.Empty;

    public ExtractFormulaDialog(string sourceAddress)
    {
        SourceAddress = sourceAddress;

        Title = "Extract Formula";
        PrimaryButtonText = "Extract";
        CloseButtonText = "Cancel";
        DefaultButton = ContentDialogButton.Primary;

        var root = new StackPanel
        {
            Spacing = 12,
            Width = 340
        };

        root.Children.Add(new TextBlock
        {
            Text = $"Move the formula from {sourceAddress} to a new cell.",
            TextWrapping = TextWrapping.Wrap,
            FontWeight = FontWeights.SemiBold
        });

        root.Children.Add(new TextBlock
        {
            Text = "Enter the destination address (e.g. Sheet2!B4 or C5).",
            TextWrapping = TextWrapping.Wrap
        });

        _targetTextBox = new TextBox
        {
            PlaceholderText = "Destination cell",
            Text = sourceAddress,
            SelectionStart = sourceAddress.Length
        };
        _targetTextBox.TextChanged += TargetTextBox_TextChanged;
        root.Children.Add(_targetTextBox);

        _validationText = new TextBlock
        {
            Foreground = new SolidColorBrush(Microsoft.UI.Colors.OrangeRed),
            TextWrapping = TextWrapping.Wrap,
            Visibility = Visibility.Collapsed,
            FontSize = 12
        };
        root.Children.Add(_validationText);

        root.Children.Add(new TextBlock
        {
            Text = "Use sheet-qualified references to move the formula to another sheet. Existing contents in the destination cell will be replaced.",
            TextWrapping = TextWrapping.Wrap,
            Opacity = 0.7,
            FontSize = 12
        });

        Content = root;

        IsPrimaryButtonEnabled = !string.IsNullOrWhiteSpace(_targetTextBox.Text);
        PrimaryButtonClick += OnPrimaryButtonClick;
    }

    private void TargetTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _validationText.Visibility = Visibility.Collapsed;
        IsPrimaryButtonEnabled = !string.IsNullOrWhiteSpace(_targetTextBox.Text);
    }

    private void OnPrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        var target = _targetTextBox.Text.Trim();
        if (string.IsNullOrEmpty(target))
        {
            _validationText.Text = "Destination address is required.";
            _validationText.Visibility = Visibility.Visible;
            IsPrimaryButtonEnabled = false;
            args.Cancel = true;
            _targetTextBox.Focus(FocusState.Programmatic);
            return;
        }

        TargetAddress = target;
    }
}
