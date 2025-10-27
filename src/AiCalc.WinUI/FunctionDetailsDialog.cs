using System;
using System.Linq;
using AiCalc.Models;
using AiCalc.Services;
using Microsoft.UI;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Windows.UI;
using Windows.UI.Text;

namespace AiCalc;

public sealed class FunctionDetailsDialog : ContentDialog
{
    public FunctionDetailsDialog(FunctionDescriptor descriptor)
    {
        ArgumentNullException.ThrowIfNull(descriptor);

        Title = $"{descriptor.Name} Details";
        CloseButtonText = "Close";
        DefaultButton = ContentDialogButton.Close;

        Content = BuildLayout(descriptor);
    }

    private UIElement BuildLayout(FunctionDescriptor descriptor)
    {
        var root = new ScrollViewer
        {
            MaxHeight = 540,
            MinWidth = 420,
            Content = BuildContent(descriptor)
        };

        return root;
    }

    private UIElement BuildContent(FunctionDescriptor descriptor)
    {
        var stack = new StackPanel
        {
            Spacing = 12,
            MaxWidth = 520
        };

        stack.Children.Add(new TextBlock
        {
            Text = $"{GetCategoryGlyph(descriptor.Category)} {descriptor.Name}",
            FontSize = 20,
            FontWeight = FontWeights.SemiBold,
            Foreground = GetBrush("TextPrimaryBrush", Microsoft.UI.Colors.White)
        });

        stack.Children.Add(new TextBlock
        {
            Text = descriptor.Description,
            TextWrapping = TextWrapping.Wrap,
            Foreground = GetBrush("TextSecondaryBrush", Microsoft.UI.Colors.LightGray)
        });

        stack.Children.Add(BuildMetadataSection(descriptor));

        stack.Children.Add(BuildDivider());
        stack.Children.Add(BuildParametersSection(descriptor));

        if (descriptor.HasExpectedOutput)
        {
            stack.Children.Add(BuildDivider());
            stack.Children.Add(BuildExpectedOutputSection(descriptor.ExpectedOutput!));
        }

        if (descriptor.HasExample)
        {
            stack.Children.Add(BuildDivider());
            stack.Children.Add(BuildExampleSection(descriptor.Example!));
        }

        return stack;
    }

    private UIElement BuildMetadataSection(FunctionDescriptor descriptor)
    {
        var section = new StackPanel
        {
            Spacing = 4
        };

        section.Children.Add(new TextBlock
        {
            Text = $"Category: {descriptor.Category}",
            FontWeight = FontWeights.SemiBold,
            Foreground = GetBrush("TextPrimaryBrush", Microsoft.UI.Colors.White)
        });

        if (descriptor.ResultType.HasValue)
        {
            section.Children.Add(new TextBlock
            {
                Text = $"Returns: {GetCellTypeLabel(descriptor.ResultType.Value)}",
                Foreground = GetBrush("TextPrimaryBrush", Microsoft.UI.Colors.White)
            });
        }

        if (descriptor.ApplicableTypes.Length > 0)
        {
            var uniqueTypes = descriptor.ApplicableTypes
                .Distinct()
                .Select(GetCellTypeLabel)
                .ToArray();

            section.Children.Add(new TextBlock
            {
                Text = $"Accepts input cells: {string.Join(", ", uniqueTypes)}",
                Foreground = GetBrush("TextSecondaryBrush", Microsoft.UI.Colors.LightGray)
            });
        }

        return section;
    }

    private UIElement BuildDivider()
    {
        return new Border
        {
            Height = 1,
            Background = GetBrush("BorderBrushColor", Color.FromArgb(255, 96, 96, 96)),
            Opacity = 0.6,
            Margin = new Thickness(0, 4, 0, 4)
        };
    }

    private UIElement BuildParametersSection(FunctionDescriptor descriptor)
    {
        var section = new StackPanel
        {
            Spacing = 8
        };

        section.Children.Add(new TextBlock
        {
            Text = "Parameters",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = GetBrush("TextPrimaryBrush", Microsoft.UI.Colors.White)
        });

        if (descriptor.Parameters.Count == 0)
        {
            section.Children.Add(new TextBlock
            {
                Text = "This function does not require any parameters.",
                Foreground = GetBrush("TextSecondaryBrush", Microsoft.UI.Colors.LightGray)
            });
            return section;
        }

        foreach (var parameter in descriptor.Parameters)
        {
            section.Children.Add(BuildParameterCard(parameter));
        }

        return section;
    }

    private UIElement BuildParameterCard(FunctionParameter parameter)
    {
        var card = new StackPanel
        {
            Spacing = 2,
            Margin = new Thickness(0, 0, 0, 4)
        };

        var title = parameter.IsOptional
            ? $"{parameter.Name} (optional)"
            : parameter.Name;

        card.Children.Add(new TextBlock
        {
            Text = title,
            FontWeight = FontWeights.SemiBold,
            Foreground = GetBrush("TextPrimaryBrush", Microsoft.UI.Colors.White)
        });

        card.Children.Add(new TextBlock
        {
            Text = parameter.Description,
            TextWrapping = TextWrapping.Wrap,
            Foreground = GetBrush("TextSecondaryBrush", Microsoft.UI.Colors.LightGray)
        });

        var accepted = string.Join(", ", parameter.AcceptableTypes.Select(GetCellTypeLabel));
        card.Children.Add(new TextBlock
        {
            Text = $"Accepts: {accepted}",
            FontStyle = Windows.UI.Text.FontStyle.Italic,
            Foreground = GetBrush("TextSecondaryBrush", Microsoft.UI.Colors.LightGray)
        });

        return card;
    }

    private UIElement BuildExpectedOutputSection(string expectedOutput)
    {
        var section = new StackPanel
        {
            Spacing = 4
        };

        section.Children.Add(new TextBlock
        {
            Text = "Typical Output",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = GetBrush("TextPrimaryBrush", Microsoft.UI.Colors.White)
        });

        section.Children.Add(new TextBlock
        {
            Text = expectedOutput,
            TextWrapping = TextWrapping.Wrap,
            Foreground = GetBrush("TextSecondaryBrush", Microsoft.UI.Colors.LightGray)
        });

        return section;
    }

    private UIElement BuildExampleSection(string example)
    {
        var section = new StackPanel
        {
            Spacing = 4
        };

        section.Children.Add(new TextBlock
        {
            Text = "Example",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = GetBrush("TextPrimaryBrush", Microsoft.UI.Colors.White)
        });

        section.Children.Add(new Border
        {
            Background = GetBrush("CardBackgroundBrush", Color.FromArgb(255, 32, 32, 32)),
            BorderBrush = GetBrush("BorderBrushColor", Color.FromArgb(255, 64, 64, 64)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8),
            Child = new TextBlock
            {
                Text = example,
                FontFamily = new FontFamily("Consolas"),
                TextWrapping = TextWrapping.Wrap,
                Foreground = GetBrush("TextPrimaryBrush", Microsoft.UI.Colors.White)
            }
        });

        return section;
    }

    private static SolidColorBrush GetBrush(string key, Color fallback)
    {
        if (Application.Current.Resources.TryGetValue(key, out var resource) && resource is SolidColorBrush brush)
        {
            return brush;
        }

        return new SolidColorBrush(fallback);
    }

    private static string GetCellTypeLabel(CellObjectType type)
    {
        return type switch
        {
            CellObjectType.Text => "Text",
            CellObjectType.Number => "Number",
            CellObjectType.Image => "Image",
            CellObjectType.Directory => "Directory",
            CellObjectType.File => "File",
            CellObjectType.Table => "Table",
            CellObjectType.DateTime => "Date/Time",
            CellObjectType.Json => "JSON",
            CellObjectType.CodePython => "Python Code",
            CellObjectType.CodeCSharp => "C# Code",
            CellObjectType.CodeJavaScript => "JavaScript Code",
            CellObjectType.CodeTypeScript => "TypeScript Code",
            CellObjectType.CodeHtml => "HTML",
            CellObjectType.CodeCss => "CSS",
            _ => type.ToString()
        };
    }

    private static string GetCategoryGlyph(FunctionCategory category)
    {
        return category switch
        {
            FunctionCategory.Math => "∑",
            FunctionCategory.Text => "✂",
            FunctionCategory.DateTime => "🕒",
            FunctionCategory.File => "📄",
            FunctionCategory.Directory => "📁",
            FunctionCategory.Table => "📊",
            FunctionCategory.Image => "🖼",
            FunctionCategory.Video => "🎬",
            FunctionCategory.Pdf => "📕",
            FunctionCategory.Data => "📈",
            FunctionCategory.AI => "🤖",
            FunctionCategory.Contrib => "✨",
            _ => "ƒ"
        };
    }
}
