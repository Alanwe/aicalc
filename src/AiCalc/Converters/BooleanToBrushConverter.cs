using System;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;

namespace AiCalc.Converters;

public class BooleanToBrushConverter : IValueConverter
{
    public Brush? TrueBrush { get; set; }

    public Brush? FalseBrush { get; set; }

    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is bool flag)
        {
            return flag ? TrueBrush ?? new SolidColorBrush(Windows.UI.Color.FromArgb(255, 79, 156, 255)) : FalseBrush ?? new SolidColorBrush(Windows.UI.Color.FromArgb(255, 30, 42, 51));
        }

        return FalseBrush ?? new SolidColorBrush(Windows.UI.Color.FromArgb(255, 30, 42, 51));
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotSupportedException();
}
