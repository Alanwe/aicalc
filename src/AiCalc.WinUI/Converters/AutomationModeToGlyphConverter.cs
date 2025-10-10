using System;
using AiCalc.Models;
using Microsoft.UI.Xaml.Data;

namespace AiCalc.Converters;

public class AutomationModeToGlyphConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is CellAutomationMode mode)
        {
            return mode switch
            {
                CellAutomationMode.AutoOnOpen => "",
                CellAutomationMode.AutoOnDependencyChange => "",
                _ => ""
            };
        }

        return "";
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotImplementedException();
}
