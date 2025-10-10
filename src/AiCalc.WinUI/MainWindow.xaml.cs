using AiCalc.Models;
using AiCalc.ViewModels;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using System;
using System.Linq;
using Windows.UI;

namespace AiCalc;

public sealed partial class MainWindow : Page
{
    public WorkbookViewModel ViewModel { get; }
    private CellViewModel? _selectedCell;
    private Button? _selectedButton;
    private bool _isUpdatingCell = false;

    public MainWindow()
    {
        ViewModel = new WorkbookViewModel();
        InitializeComponent();
        
        // Initialize with a default sheet
        if (ViewModel.Sheets.Count == 0)
        {
            ViewModel.NewSheetCommand.Execute(null);
        }
        
        // Add F9 keyboard shortcut for Recalculate All (Task 10)
        this.KeyDown += MainWindow_KeyDown;
        
        LoadFunctionsList();
        RefreshSheetTabs();
    }

    private void LoadFunctionsList()
    {
        FunctionsList.Children.Clear();
        foreach (var func in ViewModel.FunctionRegistry.Functions)
        {
            var border = new Border
            {
                Background = new SolidColorBrush(Microsoft.UI.Colors.Transparent),
                Padding = new Thickness(8),
                CornerRadius = new CornerRadius(4)
            };
            
            var stack = new StackPanel { Spacing = 4 };
            stack.Children.Add(new TextBlock 
            { 
                Text = func.Name, 
                Foreground = new SolidColorBrush(Microsoft.UI.Colors.White),
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
            });
            stack.Children.Add(new TextBlock 
            { 
                Text = func.Description, 
                Foreground = new SolidColorBrush(Color.FromArgb(0xAA, 0xFF, 0xFF, 0xFF)),
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap
            });
            
            border.Child = stack;
            FunctionsList.Children.Add(border);
        }
    }

    private void RefreshSheetTabs()
    {
        SheetTabs.TabItems.Clear();
        foreach (var sheet in ViewModel.Sheets)
        {
            var tab = new TabViewItem
            {
                Header = sheet.Name,
                Tag = sheet
            };
            SheetTabs.TabItems.Add(tab);
        }
        
        if (SheetTabs.TabItems.Count > 0)
        {
            SheetTabs.SelectedIndex = 0;
            BuildSpreadsheetGrid(ViewModel.Sheets[0]);
        }
    }

    private void BuildSpreadsheetGrid(SheetViewModel sheet)
    {
        SpreadsheetGrid.Children.Clear();
        SpreadsheetGrid.RowDefinitions.Clear();
        SpreadsheetGrid.ColumnDefinitions.Clear();
        
        var rows = sheet.Rows.Count;
        var cols = sheet.ColumnHeaders.Count;
        
        // Add header row + data rows
        SpreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(32) });
        for (int i = 0; i < rows; i++)
        {
            SpreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(70) });
        }
        
        // Add row header column + data columns
        SpreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(48) });
        for (int i = 0; i < cols; i++)
        {
            SpreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        }
        
        // Top-left corner
        AddCell(new Border 
        { 
            Background = new SolidColorBrush(Color.FromArgb(0xFF, 0x25, 0x25, 0x26)),
            CornerRadius = new CornerRadius(4),
            Margin = new Thickness(2)
        }, 0, 0);
        
        // Column headers
        for (int col = 0; col < cols; col++)
        {
            var header = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(0xFF, 0x25, 0x25, 0x26)),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(2),
                Child = new TextBlock
                {
                    Text = sheet.ColumnHeaders[col],
                    Foreground = new SolidColorBrush(Color.FromArgb(0xCC, 0xFF, 0xFF, 0xFF)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
                }
            };
            AddCell(header, 0, col + 1);
        }
        
        // Row headers and cells
        for (int row = 0; row < rows; row++)
        {
            var rowVm = sheet.Rows[row];
            
            // Row header
            var rowHeader = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(0xFF, 0x25, 0x25, 0x26)),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(2),
                Child = new TextBlock
                {
                    Text = rowVm.Label,
                    Foreground = new SolidColorBrush(Color.FromArgb(0xCC, 0xFF, 0xFF, 0xFF)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
                }
            };
            AddCell(rowHeader, row + 1, 0);
            
            // Data cells
            for (int col = 0; col < rowVm.Cells.Count; col++)
            {
                var cellVm = rowVm.Cells[col];
                var button = CreateCellButton(cellVm);
                AddCell(button, row + 1, col + 1);
            }
        }
    }

    private Button CreateCellButton(CellViewModel cellVm)
    {
        var button = new Button
        {
            Background = new SolidColorBrush(Color.FromArgb(0x22, 0xFF, 0xFF, 0xFF)),
            BorderBrush = new SolidColorBrush(Color.FromArgb(0x44, 0xFF, 0xFF, 0xFF)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Margin = new Thickness(2),
            Padding = new Thickness(10),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Stretch,
            VerticalContentAlignment = VerticalAlignment.Stretch,
            Tag = cellVm
        };
        
        var stack = new StackPanel { Spacing = 6 };
        
        // Show value directly without cell label
        var valueText = new TextBlock
        {
            Text = cellVm.DisplayValue,
            Foreground = new SolidColorBrush(Microsoft.UI.Colors.White),
            FontSize = 13,
            TextWrapping = TextWrapping.Wrap,
            MaxLines = 3,
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center
        };
        
        stack.Children.Add(valueText);
        
        button.Content = stack;
        button.Click += (s, e) => SelectCell(cellVm, button);
        button.DoubleTapped += (s, e) => StartDirectEdit(cellVm, button);
        
        return button;
    }

    private void AddCell(FrameworkElement element, int row, int col)
    {
        Grid.SetRow(element, row);
        Grid.SetColumn(element, col);
        SpreadsheetGrid.Children.Add(element);
    }

    private void SelectCell(CellViewModel cell, Button button)
    {
        // Deselect previous cell
        if (_selectedCell != null)
        {
            _selectedCell.IsSelected = false;
        }
        
        // Reset all button backgrounds
        foreach (var child in SpreadsheetGrid.Children.OfType<Button>())
        {
            child.Background = new SolidColorBrush(Color.FromArgb(0x22, 0xFF, 0xFF, 0xFF));
        }
        
        // Select new cell
        _selectedCell = cell;
        _selectedButton = button;
        cell.IsSelected = true;
        button.Background = new SolidColorBrush(Color.FromArgb(0x4A, 0x00, 0xD4, 0xFF));
        
        // Update inspector
        _isUpdatingCell = true;
        CellInspectorPanel.Visibility = Visibility.Visible;
        NoCellPanel.Visibility = Visibility.Collapsed;
        
        CellLabel.Text = cell.DisplayLabel;
        CellValueBox.Text = cell.RawValue;
        CellFormulaBox.Text = cell.Formula;
        CellNotesBox.Text = cell.Notes;
        AutomationModeBox.SelectedIndex = (int)cell.AutomationMode;
        
        _isUpdatingCell = false;
    }

    private void StartDirectEdit(CellViewModel cell, Button button)
    {
        // Only allow direct edit for text/empty cells (not images, directories, etc.)
        if (cell.Value.ObjectType != CellObjectType.Text && 
            cell.Value.ObjectType != CellObjectType.Number &&
            cell.Value.ObjectType != CellObjectType.Empty)
        {
            return;
        }

        // Position the edit box over the cell
        var buttonPos = button.TransformToVisual(SpreadsheetGrid).TransformPoint(new Windows.Foundation.Point(0, 0));
        
        CellEditBox.Width = button.ActualWidth;
        CellEditBox.Height = button.ActualHeight;
        CellEditBox.Margin = new Thickness(buttonPos.X, buttonPos.Y, 0, 0);
        CellEditBox.HorizontalAlignment = HorizontalAlignment.Left;
        CellEditBox.VerticalAlignment = VerticalAlignment.Top;
        
        CellEditBox.Text = cell.RawValue;
        CellEditBox.Tag = cell;
        CellEditBox.Visibility = Visibility.Visible;
        CellEditBox.Focus(FocusState.Programmatic);
        CellEditBox.SelectAll();
    }

    private void CellEditBox_LostFocus(object sender, RoutedEventArgs e)
    {
        CommitCellEdit();
    }

    private void CellEditBox_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        if (e.Key == Windows.System.VirtualKey.Enter)
        {
            CommitCellEdit();
            e.Handled = true;
        }
        else if (e.Key == Windows.System.VirtualKey.Escape)
        {
            CellEditBox.Visibility = Visibility.Collapsed;
            e.Handled = true;
        }
    }

    private void CommitCellEdit()
    {
        if (CellEditBox.Tag is CellViewModel cell)
        {
            cell.RawValue = CellEditBox.Text;
            CellEditBox.Visibility = Visibility.Collapsed;
            
            // Refresh the grid
            if (ViewModel.SelectedSheet != null)
            {
                BuildSpreadsheetGrid(ViewModel.SelectedSheet);
            }
            
            // Update inspector if this cell is selected
            if (_selectedCell == cell)
            {
                _isUpdatingCell = true;
                CellValueBox.Text = cell.RawValue;
                _isUpdatingCell = false;
            }
        }
    }

    private void SheetTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (SheetTabs.SelectedItem is TabViewItem tab && tab.Tag is SheetViewModel sheet)
        {
            ViewModel.SelectedSheet = sheet;
            BuildSpreadsheetGrid(sheet);
        }
    }

    private void AddTab_Click(TabView sender, object args)
    {
        NewSheetButton_Click(sender, null!);
    }

    private void NewSheetButton_Click(object sender, RoutedEventArgs e)
    {
        ViewModel.NewSheetCommand.Execute(null);
        RefreshSheetTabs();
        SheetTabs.SelectedIndex = SheetTabs.TabItems.Count - 1;
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        ViewModel.SaveCommand.Execute(null);
    }

    private void LoadButton_Click(object sender, RoutedEventArgs e)
    {
        ViewModel.LoadCommand.Execute(null);
        RefreshSheetTabs();
    }

    private async void RecalculateButton_Click(object sender, RoutedEventArgs e)
    {
        // Recalculate All - same as F9 (Task 10)
        await ViewModel.EvaluateWorkbookCommand.ExecuteAsync(null);
        
        // Refresh display
        if (ViewModel.SelectedSheet != null)
        {
            BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        }
    }

    private void CellValue_Changed(object sender, TextChangedEventArgs e)
    {
        if (_isUpdatingCell || _selectedCell == null) return;
        _selectedCell.RawValue = CellValueBox.Text;
    }

    private void CellFormula_Changed(object sender, TextChangedEventArgs e)
    {
        if (_isUpdatingCell || _selectedCell == null) return;
        _selectedCell.Formula = CellFormulaBox.Text;
    }

    private void CellNotes_Changed(object sender, TextChangedEventArgs e)
    {
        if (_isUpdatingCell || _selectedCell == null) return;
        _selectedCell.Notes = CellNotesBox.Text;
    }

    private void AutomationMode_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingCell || _selectedCell == null) return;
        _selectedCell.AutomationMode = (CellAutomationMode)AutomationModeBox.SelectedIndex;
    }

    private async void EvaluateCell_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell != null && ViewModel.SelectedSheet != null)
        {
            await _selectedCell.EvaluateAsync();
            
            // Refresh the cell display
            if (SheetTabs.SelectedItem is TabViewItem tab && tab.Tag is SheetViewModel sheet)
            {
                BuildSpreadsheetGrid(sheet);
            }
        }
    }

    private async void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        // Show full settings dialog with Performance and Appearance tabs (Task 9 & 10)
        var dialog = new SettingsDialog(ViewModel.Settings)
        {
            XamlRoot = this.Content.XamlRoot
        };

        var result = await dialog.ShowAsync();
        
        if (result == ContentDialogResult.Primary)
        {
            // Apply evaluation settings
            ViewModel.UpdateEvaluationSettings();
            
            // Apply theme
            App.ApplyCellStateTheme(ViewModel.Settings.SelectedTheme);
            
            // Refresh UI to show new theme
            if (ViewModel.SelectedSheet != null)
            {
                BuildSpreadsheetGrid(ViewModel.SelectedSheet);
            }
            
            ViewModel.StatusMessage = $"Settings saved: {ViewModel.Settings.MaxEvaluationThreads} threads, {ViewModel.Settings.DefaultEvaluationTimeoutSeconds}s timeout, {ViewModel.Settings.SelectedTheme} theme";
        }
    }

    /// <summary>
    /// Handle F9 keyboard shortcut for Recalculate All (Task 10)
    /// </summary>
    private async void MainWindow_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        // F9: Recalculate all (skip Manual cells per Task 10)
        if (e.Key == Windows.System.VirtualKey.F9)
        {
            e.Handled = true;
            await ViewModel.EvaluateWorkbookCommand.ExecuteAsync(null);
            
            // Refresh display
            if (ViewModel.SelectedSheet != null)
            {
                BuildSpreadsheetGrid(ViewModel.SelectedSheet);
            }
        }
    }
}
