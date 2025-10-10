using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;

namespace AiCalc;

public sealed partial class EvaluationSettingsDialog : ContentDialog
{
    public int MaxThreadCount { get; private set; }
    public int TimeoutSeconds { get; private set; }

    public EvaluationSettingsDialog()
    {
        InitializeComponent();

        // Set defaults from environment
        var cpuCores = Environment.ProcessorCount;
        MaxThreadCount = cpuCores;
        TimeoutSeconds = 100;

        // Initialize UI
        ThreadCountSlider.Value = MaxThreadCount;
        ThreadCountSlider.Maximum = Math.Max(32, cpuCores * 2);
        TimeoutNumberBox.Value = TimeoutSeconds;

        UpdateThreadCountDescription(cpuCores);
    }

    public EvaluationSettingsDialog(int currentMaxThreads, int currentTimeoutSeconds) : this()
    {
        MaxThreadCount = currentMaxThreads;
        TimeoutSeconds = currentTimeoutSeconds;

        ThreadCountSlider.Value = MaxThreadCount;
        TimeoutNumberBox.Value = TimeoutSeconds;
    }

    private void ThreadCountSlider_ValueChanged(object sender, Microsoft.UI.Xaml.Controls.Primitives.RangeBaseValueChangedEventArgs e)
    {
        var value = (int)e.NewValue;
        ThreadCountLabel.Text = value.ToString();
        UpdateThreadCountDescription(Environment.ProcessorCount);
    }

    private void UpdateThreadCountDescription(int cpuCores)
    {
        var selectedThreads = (int)ThreadCountSlider.Value;
        ThreadCountDescription.Text = $"Using {selectedThreads} threads (CPU cores detected: {cpuCores})";
    }

    private void ContentDialog_PrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        // Save values
        MaxThreadCount = (int)ThreadCountSlider.Value;
        TimeoutSeconds = (int)TimeoutNumberBox.Value;
    }
}
