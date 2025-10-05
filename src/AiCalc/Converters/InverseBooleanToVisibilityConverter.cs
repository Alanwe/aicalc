using System;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;

namespace AiCalc.Converters;

public class InverseBooleanToVisibilityConverter : IValueConverter
{
    public bool CollapseWhenTrue { get; set; } = true;

    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is bool flag && flag)
        {
            return CollapseWhenTrue ? Visibility.Collapsed : Visibility.Hidden;
        }

        return Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotSupportedException();
}
