using System;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;

namespace AiCalc.Converters;

public class BooleanToVisibilityConverter : IValueConverter
{
    public bool CollapseWhenFalse { get; set; } = true;

    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is bool flag && flag)
        {
            return Visibility.Visible;
        }

        return CollapseWhenFalse ? Visibility.Collapsed : Visibility.Hidden;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotSupportedException();
}
