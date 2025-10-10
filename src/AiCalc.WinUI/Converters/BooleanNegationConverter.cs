using System;
using Microsoft.UI.Xaml.Data;

namespace AiCalc.Converters;

public class BooleanNegationConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is bool flag)
        {
            return !flag;
        }

        return true;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotSupportedException();
}
