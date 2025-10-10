using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using System;

namespace AiCalc.Converters;

public class ObjectToBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        return value != null;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotImplementedException();
    }
}
