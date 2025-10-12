using System;
using System.Text.Json.Serialization;

namespace AiCalc.Models;

/// <summary>
/// Describes formatting applied to a cell (Phase 5 Task 16 requirement).
/// </summary>
public class CellFormat
{
    private const string DefaultBackground = "#22222222";
    private const string DefaultForeground = "#FFFFFFFF";
    private const string DefaultBorder = "#44FFFFFF";

    public static string DefaultBackgroundColor => DefaultBackground;
    public static string DefaultForegroundColor => DefaultForeground;
    public static string DefaultBorderColor => DefaultBorder;

    public string Background { get; set; } = DefaultBackground;

    public string Foreground { get; set; } = DefaultForeground;

    public string BorderBrush { get; set; } = DefaultBorder;

    public double BorderThickness { get; set; } = 1.0;

    public double FontSize { get; set; } = 13.0;

    public string FontFamily { get; set; } = "Segoe UI";

    public bool IsBold { get; set; }

    public bool IsItalic { get; set; }

    public string HorizontalAlignment { get; set; } = "Left";

    public string VerticalAlignment { get; set; } = "Center";

    [JsonIgnore]
    public static CellFormat Default => new();

    public CellFormat Clone()
    {
        return new CellFormat
        {
            Background = Background,
            Foreground = Foreground,
            BorderBrush = BorderBrush,
            BorderThickness = BorderThickness,
            FontSize = FontSize,
            FontFamily = FontFamily,
            IsBold = IsBold,
            IsItalic = IsItalic,
            HorizontalAlignment = HorizontalAlignment,
            VerticalAlignment = VerticalAlignment
        };
    }

    public override string ToString()
    {
        return $"Bg={Background},Fg={Foreground},Border={BorderBrush},Font={FontFamily} {FontSize}";
    }

    public static CellFormat From(CellFormat? format)
    {
        return format is null ? Default : format.Clone();
    }
}
