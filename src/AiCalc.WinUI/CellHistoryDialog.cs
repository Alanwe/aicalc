using System.Collections.Generic;
using System.Linq;
using AiCalc.Models;
using AiCalc.ViewModels;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Markup;

namespace AiCalc;

public sealed class CellHistoryDialog : ContentDialog
{
    private readonly TextBlock _headerText;
    private readonly TextBlock _emptyText;
    private readonly ListView _historyList;
    private static readonly DataTemplate HistoryTemplate = (DataTemplate)XamlReader.Load(
        "<DataTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>" +
        "<StackPanel Spacing='4' Padding='4'>" +
        "  <TextBlock Text='{Binding Summary}' TextWrapping='Wrap' FontWeight='SemiBold'/>" +
        "  <TextBlock Text='{Binding Notes}' TextWrapping='Wrap'/>" +
        "</StackPanel>" +
        "</DataTemplate>");

    public CellHistoryDialog()
    {
        Title = "Cell History";
        PrimaryButtonText = "Close";
        IsPrimaryButtonEnabled = true;

        var root = new Grid
        {
            Width = 400
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        _headerText = new TextBlock
        {
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 12)
        };
        root.Children.Add(_headerText);

        var bodyGrid = new Grid();
        Grid.SetRow(bodyGrid, 1);
        root.Children.Add(bodyGrid);

        _emptyText = new TextBlock
        {
            Text = "No history available yet.",
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        bodyGrid.Children.Add(_emptyText);

        _historyList = new ListView
        {
            ItemTemplate = HistoryTemplate
        };
        bodyGrid.Children.Add(_historyList);

        Content = root;
    }

    public CellHistoryDialog(CellViewModel cell) : this()
    {
        Initialize(cell);
    }

    public void Initialize(CellViewModel cell)
    {
        Initialize(cell.Address, cell.History);
    }

    public void Initialize(CellAddress address, IEnumerable<CellHistoryEntry> history)
    {
        var snapshot = history?.Select(entry => entry.Clone()).ToList() ?? new List<CellHistoryEntry>();

        _headerText.Text = $"Cell {address} history";
        _historyList.ItemsSource = snapshot;

        var hasEntries = snapshot.Count > 0;
        _emptyText.Visibility = hasEntries ? Visibility.Collapsed : Visibility.Visible;
        _historyList.Visibility = hasEntries ? Visibility.Visible : Visibility.Collapsed;
    }
}
