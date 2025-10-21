using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Windows.UI;
using AiCalc.Services;

namespace AiCalc
{
    public class ParameterHintItem
    {
    public string Name { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public Windows.UI.Text.FontWeight Weight { get; init; } = Microsoft.UI.Text.FontWeights.Normal;
    public Brush Brush { get; init; } = new SolidColorBrush(Microsoft.UI.Colors.White);
    public Visibility DescriptionVisibility => string.IsNullOrWhiteSpace(Description) ? Visibility.Collapsed : Visibility.Visible;

        public static ParameterHintItem From(FunctionParameter param, bool isActive)
        {
            return new ParameterHintItem
            {
                Name = param.IsOptional ? $"{param.Name} (optional)" : param.Name,
                Description = param.Description,
                Weight = isActive ? Microsoft.UI.Text.FontWeights.SemiBold : Microsoft.UI.Text.FontWeights.Normal,
                Brush = new SolidColorBrush(isActive ? Microsoft.UI.Colors.DeepSkyBlue : Microsoft.UI.Colors.White)
            };
        }
    }
}