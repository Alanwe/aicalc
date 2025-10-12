using System;
using AiCalc.Models;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.UI;

namespace AiCalc;

public sealed class FormatCellDialog : ContentDialog
{
    private readonly ColorPicker _backgroundPicker;
    private readonly ColorPicker _foregroundPicker;
    private readonly Slider _fontSizeSlider;
    private readonly TextBlock _fontSizeValueText;
    private readonly CheckBox _boldToggle;
    private readonly CheckBox _italicToggle;
    private readonly Slider _borderSlider;
    private readonly TextBlock _borderValueText;

    public FormatCellDialog(CellFormat current)
    {
        Title = "Format Cell";
        PrimaryButtonText = "Apply";
        CloseButtonText = "Cancel";
        DefaultButton = ContentDialogButton.Primary;

        SelectedFormat = CellFormat.From(current);

        _backgroundPicker = new ColorPicker
        {
            IsAlphaEnabled = true
        };

        _foregroundPicker = new ColorPicker
        {
            IsAlphaEnabled = true
        };

        _fontSizeSlider = new Slider
        {
            Minimum = 8,
            Maximum = 36,
            Width = 160
        };

        _fontSizeValueText = new TextBlock
        {
            VerticalAlignment = VerticalAlignment.Center
        };

        _boldToggle = new CheckBox
        {
            Content = "Bold"
        };

        _italicToggle = new CheckBox
        {
            Content = "Italic"
        };

        _borderSlider = new Slider
        {
            Minimum = 0,
            Maximum = 4,
            Width = 160
        };

        _borderValueText = new TextBlock
        {
            VerticalAlignment = VerticalAlignment.Center
        };

        Content = BuildLayout();

        InitializePickers();
        UpdateSliderText();

        _fontSizeSlider.ValueChanged += (_, args) => _fontSizeValueText.Text = args.NewValue.ToString("F0");
        _borderSlider.ValueChanged += (_, args) => _borderValueText.Text = args.NewValue.ToString("F1");

        PrimaryButtonClick += OnPrimaryButtonClick;
    }

    public CellFormat SelectedFormat { get; private set; }

    private UIElement BuildLayout()
    {
        var layout = new StackPanel
        {
            Spacing = 16,
            Width = 320
        };

        layout.Children.Add(new TextBlock
        {
            Text = "Background",
            FontWeight = FontWeights.SemiBold
        });
        layout.Children.Add(_backgroundPicker);

        layout.Children.Add(new TextBlock
        {
            Text = "Text",
            FontWeight = FontWeights.SemiBold
        });
        layout.Children.Add(_foregroundPicker);

        layout.Children.Add(BuildFontSizeRow());

        var styleRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 12
        };
        styleRow.Children.Add(_boldToggle);
        styleRow.Children.Add(_italicToggle);
        layout.Children.Add(styleRow);

        layout.Children.Add(BuildBorderRow());

        return layout;
    }

    private UIElement BuildFontSizeRow()
    {
        var row = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 12,
            VerticalAlignment = VerticalAlignment.Center
        };

        row.Children.Add(new TextBlock
        {
            Text = "Font Size",
            VerticalAlignment = VerticalAlignment.Center
        });

        row.Children.Add(_fontSizeSlider);
        row.Children.Add(_fontSizeValueText);

        return row;
    }

    private UIElement BuildBorderRow()
    {
        var row = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 12,
            VerticalAlignment = VerticalAlignment.Center
        };

        row.Children.Add(new TextBlock
        {
            Text = "Border",
            VerticalAlignment = VerticalAlignment.Center
        });

        row.Children.Add(_borderSlider);
        row.Children.Add(_borderValueText);

        return row;
    }

    private void InitializePickers()
    {
        _backgroundPicker.Color = ToColor(SelectedFormat.Background);
        _foregroundPicker.Color = ToColor(SelectedFormat.Foreground);
        _fontSizeSlider.Value = SelectedFormat.FontSize;
        _borderSlider.Value = SelectedFormat.BorderThickness;
        _boldToggle.IsChecked = SelectedFormat.IsBold;
        _italicToggle.IsChecked = SelectedFormat.IsItalic;
    }

    private void UpdateSliderText()
    {
        _fontSizeValueText.Text = _fontSizeSlider.Value.ToString("F0");
        _borderValueText.Text = _borderSlider.Value.ToString("F1");
    }

    private void OnPrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        SelectedFormat = new CellFormat
        {
            Background = ToHex(_backgroundPicker.Color),
            Foreground = ToHex(_foregroundPicker.Color),
            BorderBrush = SelectedFormat.BorderBrush,
            BorderThickness = _borderSlider.Value,
            FontSize = _fontSizeSlider.Value,
            FontFamily = SelectedFormat.FontFamily,
            IsBold = _boldToggle.IsChecked ?? false,
            IsItalic = _italicToggle.IsChecked ?? false,
            HorizontalAlignment = SelectedFormat.HorizontalAlignment,
            VerticalAlignment = SelectedFormat.VerticalAlignment
        };
    }

    private static Color ToColor(string hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
        {
            return Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
        }

        hex = hex.TrimStart('#');
        try
        {
            byte a = 0xFF;
            byte r;
            byte g;
            byte b;

            if (hex.Length == 6)
            {
                r = Convert.ToByte(hex.Substring(0, 2), 16);
                g = Convert.ToByte(hex.Substring(2, 2), 16);
                b = Convert.ToByte(hex.Substring(4, 2), 16);
            }
            else if (hex.Length == 8)
            {
                a = Convert.ToByte(hex.Substring(0, 2), 16);
                r = Convert.ToByte(hex.Substring(2, 2), 16);
                g = Convert.ToByte(hex.Substring(4, 2), 16);
                b = Convert.ToByte(hex.Substring(6, 2), 16);
            }
            else
            {
                return Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
            }

            return Color.FromArgb(a, r, g, b);
        }
        catch
        {
            return Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
        }
    }

    private static string ToHex(Color color)
    {
        return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
    }
}
