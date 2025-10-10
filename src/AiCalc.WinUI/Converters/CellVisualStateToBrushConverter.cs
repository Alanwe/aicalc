using System;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;
using AiCalc.Models;

namespace AiCalc.Converters;

public class CellVisualStateToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is CellVisualState state)
        {
            // Try to get theme brush from app resources
            var resourceKey = state switch
            {
                CellVisualState.JustUpdated => "CellStateJustUpdatedBrush",
                CellVisualState.Calculating => "CellStateCalculatingBrush",
                CellVisualState.Stale => "CellStateStaleBrush",
                CellVisualState.ManualUpdate => "CellStateManualUpdateBrush",
                CellVisualState.Error => "CellStateErrorBrush",
                CellVisualState.InDependencyChain => "CellStateInDependencyChainBrush",
                _ => "CellStateNormalBrush"
            };

            // Try to get from resources, fallback to hardcoded colors
            if (Application.Current.Resources.TryGetValue(resourceKey, out var resource) && resource is Brush brush)
            {
                return brush;
            }

            // Fallback to hardcoded colors if theme not loaded
            return state switch
            {
                CellVisualState.JustUpdated => new SolidColorBrush(Colors.LimeGreen),
                CellVisualState.Calculating => new SolidColorBrush(Colors.Orange),
                CellVisualState.Stale => new SolidColorBrush(Colors.DodgerBlue),
                CellVisualState.ManualUpdate => new SolidColorBrush(Colors.Orange),
                CellVisualState.Error => new SolidColorBrush(Colors.Red),
                CellVisualState.InDependencyChain => new SolidColorBrush(Colors.Yellow),
                _ => new SolidColorBrush(Colors.Transparent)
            };
        }
        
        return new SolidColorBrush(Colors.Transparent);
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotImplementedException();
    }
}

