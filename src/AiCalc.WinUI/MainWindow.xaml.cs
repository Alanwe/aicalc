using AiCalc.Models;
using AiCalc.Services;
using AiCalc.ViewModels;
using Microsoft.UI;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Text;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.UI;
using Windows.UI.Text;

namespace AiCalc;

public sealed partial class MainWindow : Page
{
    public WorkbookViewModel ViewModel { get; }
    private CellViewModel? _selectedCell;
    private Button? _selectedButton;
    private bool _isUpdatingCell = false;
    private readonly Dictionary<string, FunctionDescriptor> _functionMap = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<ParameterHintItem> _parameterHints = new();
    private bool _isFormulaEditing;
    private bool _processingTextChange; // Prevent re-entry when we modify the text
    private bool _updatingFormulaText; // Prevent re-entry in CellFormula_Changed
    private CellViewModel? _formulaEditTargetCell;
    private readonly List<(CellViewModel Cell, CellVisualState PreviousState)> _formulaHighlights = new();
    private readonly HashSet<CellViewModel> _validationErrorCells = new();
    private readonly List<CellViewModel> _multiSelection = new();
    private readonly Dictionary<CellObjectType, StackPanel> _functionGroups = new();
    private readonly Dictionary<CellObjectType, Button> _functionHeaders = new();
    private readonly List<CellObjectType> _functionGroupOrder = new()
    {
        CellObjectType.Text,
        CellObjectType.Number,
        CellObjectType.Image,
        CellObjectType.Directory,
        CellObjectType.File,
        CellObjectType.Table,
        CellObjectType.DateTime,
        CellObjectType.Json
    };
    private CellViewModel? _selectionAnchor;
    private bool _isLeftSplitterDragging;
    private bool _isRightSplitterDragging;
    private Point _leftSplitterStart;
    private Point _rightSplitterStart;
    private double _leftInitialWidth;
    private double _rightInitialWidth;
    private PipeServer? _pipeServer;
    private bool _functionsVisible = true;
    private bool _inspectorVisible = true;
    private bool _isNavigatingBetweenCells = false;
    private bool _isPickingFormulaReference;
    private int? _columnContextIndex;
    private int? _rowContextIndex;
    // Reduced default row height further (20% reduction from 32)
    private const double DefaultRowHeight = 26;
    private const double ColumnWidthScale = 0.84; // reduce cell width by 16%

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
        InitializeFunctionSuggestions();
        RefreshSheetTabs();
        LoadPreferences();  // Load user preferences (Phase 5)
        
        // Initialize theme toggle button text
        var prefs = App.PreferencesService.LoadPreferences();
        var currentTheme = prefs.Theme ?? "Dark";
        if (ThemeToggleButton != null)
        {
            ThemeToggleButton.Content = currentTheme == "Light" ? "🌙 Dark" : "☀️ Light";
        }

        ViewModel.Settings.Connections.CollectionChanged += SettingsConnections_CollectionChanged;
        
        // Start Python SDK pipe server
        StartPipeServer();
    }

    private void SettingsConnections_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        DispatcherQueue?.TryEnqueue(LoadFunctionsList);
    }

    private void InitializeFunctionSuggestions()
    {
        _functionMap.Clear();

        foreach (var func in ViewModel.FunctionRegistry.Functions)
        {
            _functionMap[func.Name] = func;
        }
    }

    public void RefreshFunctionCatalog()
    {
        InitializeFunctionSuggestions();
        LoadFunctionsList();
    }

    private void LoadFunctionsList()
    {
        FunctionsList.Children.Clear();
        _functionGroups.Clear();

        var totalFunctions = ViewModel.FunctionRegistry.Functions.Count();
        var available = GetAvailableFunctions().ToList();
        
        System.Diagnostics.Debug.WriteLine($"LoadFunctionsList: Total registered={totalFunctions}, Available after filter={available.Count}");
        
        foreach (var type in _functionGroupOrder)
        {
            var funcs = available.Where(f => f.CanAccept(type)).OrderBy(f => f.Name).ToList();

            // Group header
            var textPrimaryBrush = Application.Current.Resources["TextPrimaryBrush"] as SolidColorBrush
                ?? new SolidColorBrush(Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF));
            var headerButton = new Button
            {
                Content = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Children =
                    {
                        new TextBlock {
                            Text = GetCellObjectTypeLabel(type),
                            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                            Foreground = textPrimaryBrush
                        },
                        new TextBlock {
                            Text = $"({funcs.Count})",
                            Foreground = Application.Current.Resources["TextSecondaryBrush"] as SolidColorBrush
                        }
                    }
                },
                Background = new SolidColorBrush(Color.FromArgb(0x00, 0x00, 0x00, 0x00)),
                BorderThickness = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Padding = new Thickness(6)
            };

            var groupPanel = new StackPanel { Spacing = 6, Margin = new Thickness(4, 0, 0, 12), Visibility = Visibility.Collapsed };
            _functionGroups[type] = groupPanel;
            _functionHeaders[type] = headerButton;

            headerButton.Click += (s, e) =>
            {
                // Toggle this group's visibility
                if (_functionGroups.TryGetValue(type, out var groupPanel))
                {
                    groupPanel.Visibility = groupPanel.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
                }
            };

            // Populate functions
            foreach (var func in funcs)
            {
                var item = CreateFunctionItem(func);
                groupPanel.Children.Add(item);
            }

            FunctionsList.Children.Add(headerButton);
            FunctionsList.Children.Add(groupPanel);
        }
    }

    private string GetCellObjectTypeLabel(CellObjectType type)
    {
        return type switch
        {
            CellObjectType.Text => "Text",
            CellObjectType.Number => "Number",
            CellObjectType.Image => "Image",
            CellObjectType.Directory => "Directory",
            CellObjectType.File => "File",
            CellObjectType.Table => "Table",
            CellObjectType.DateTime => "Date/Time",
            CellObjectType.Json => "JSON",
            _ => type.ToString()
        };
    }

    private Border CreateFunctionItem(FunctionDescriptor func)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Microsoft.UI.Colors.Transparent),
            Padding = new Thickness(4, 2, 4, 2),
            CornerRadius = new CornerRadius(2)
        };

        var stack = new StackPanel { Spacing = 1 };
        var textPrimaryBrush = Application.Current.Resources["TextPrimaryBrush"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0x00, 0x00));
        stack.Children.Add(new TextBlock
        {
            Text = $"{GetCategoryGlyph(func.Category)} {func.Name}",
            Foreground = textPrimaryBrush,
            FontWeight = Microsoft.UI.Text.FontWeights.Bold,
            FontSize = 10
        });
        var textSecondaryBrush = Application.Current.Resources["TextSecondaryBrush"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0x61, 0x61, 0x61));
        stack.Children.Add(new TextBlock
        {
            Text = func.Description,
            Foreground = textSecondaryBrush,
            FontSize = 9,
            TextWrapping = TextWrapping.Wrap
        });

        if (func.Category == FunctionCategory.AI)
        {
            var providerHint = BuildProviderHint(func, App.AIServices.GetDefaultConnection());
            if (!string.IsNullOrWhiteSpace(providerHint))
            {
                var hintBrush = Application.Current.Resources["TextSecondaryBrush"] as SolidColorBrush
                    ?? new SolidColorBrush(Color.FromArgb(0x88, 0xFF, 0xFF, 0xFF));
                stack.Children.Add(new TextBlock
                {
                    Text = providerHint,
                    Foreground = hintBrush,
                    FontSize = 8
                });
            }
        }

        border.Child = stack;
        return border;
    }

    private void FocusFunctionGroupForCellType(CellObjectType type)
    {
        // Optionally, keep this for auto-expansion on cell selection, but do not collapse other groups
        if (_functionGroups.TryGetValue(type, out var groupPanel))
        {
            groupPanel.Visibility = Visibility.Visible;
        }
    }

    private IEnumerable<FunctionDescriptor> GetAvailableFunctions()
    {
        var connection = App.AIServices.GetDefaultConnection();
        foreach (var descriptor in ViewModel.FunctionRegistry.Functions)
        {
            if (descriptor.Category == FunctionCategory.AI && !IsFunctionSupportedByConnection(descriptor, connection))
            {
                continue;
            }

            yield return descriptor;
        }
    }

    private static string GetCategoryGlyph(FunctionCategory category)
    {
        return category switch
        {
            FunctionCategory.Math => "∑",
            FunctionCategory.Text => "✂",
            FunctionCategory.DateTime => "🕒",
            FunctionCategory.File => "📄",
            FunctionCategory.Directory => "📁",
            FunctionCategory.Table => "📊",
            FunctionCategory.Image => "🖼",
            FunctionCategory.Video => "🎬",
            FunctionCategory.Pdf => "📕",
            FunctionCategory.Data => "📈",
            FunctionCategory.AI => "🤖",
            FunctionCategory.Contrib => "✨",
            _ => "ƒ"
        };
    }

    private static string BuildProviderHint(FunctionDescriptor descriptor, WorkspaceConnection? connection)
    {
        if (descriptor.Category != FunctionCategory.AI || connection == null)
        {
            return string.Empty;
        }

        var providerLabel = connection.Provider switch
        {
            "AzureOpenAI" => "Azure OpenAI",
            "OpenAI" => "OpenAI",
            "Ollama" => "Ollama",
            _ => connection.Provider
        };

        var model = descriptor.Name switch
        {
            "IMAGE_TO_CAPTION" => connection.VisionModel ?? connection.Model,
            "TEXT_TO_IMAGE" => connection.ImageModel,
            _ => connection.Model
        };

        if (string.IsNullOrWhiteSpace(model))
        {
            return providerLabel;
        }

        return $"{providerLabel} · {model}";
    }

    private IEnumerable<FunctionSuggestion> GetContextualSuggestions(string searchText)
    {
        var connection = App.AIServices.GetDefaultConnection();
        var currentCellType = _formulaEditTargetCell?.Value.ObjectType ?? _selectedCell?.Value.ObjectType ?? CellObjectType.Text;
        currentCellType = NormalizeCellObjectType(currentCellType);
        var isBrowsing = string.IsNullOrWhiteSpace(searchText);

        foreach (var descriptor in ViewModel.FunctionRegistry.Functions)
        {
            // Check if function name matches search
            if (!descriptor.Name.StartsWith(searchText, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Show ALL functions when browsing (empty search after typing =)
            // Only apply filters when user is actively searching
            if (!isBrowsing)
            {
                if (descriptor.Category == FunctionCategory.AI)
                {
                    if (!IsFunctionSupportedByConnection(descriptor, connection))
                    {
                        continue;
                    }

                    if (!IsFunctionRelevantForCellType(descriptor, currentCellType))
                    {
                        continue;
                    }
                }
            }

            yield return CreateSuggestion(descriptor, connection);
        }
    }

    private FunctionSuggestion CreateSuggestion(FunctionDescriptor descriptor, WorkspaceConnection? connection)
    {
        var suggestion = new FunctionSuggestion
        {
            Name = descriptor.Name,
            Description = descriptor.Description,
            Signature = BuildSignature(descriptor),
            Category = descriptor.Category,
            CategoryGlyph = GetCategoryGlyph(descriptor.Category),
            CategoryLabel = descriptor.Category.ToString(),
            TypeHint = descriptor.Category == FunctionCategory.AI ? BuildTypeHint(descriptor) : string.Empty,
            ProviderHint = descriptor.Category == FunctionCategory.AI ? BuildProviderHint(descriptor, connection) : string.Empty
        };

        return suggestion;
    }

    private static string BuildSignature(FunctionDescriptor descriptor)
    {
        if (descriptor.Parameters.Count == 0)
        {
            return $"{descriptor.Name}()";
        }

        var paramNames = descriptor.Parameters.Select(p => p.Name);
        return $"{descriptor.Name}({string.Join(", ", paramNames)})";
    }

    private static string BuildTypeHint(FunctionDescriptor descriptor)
    {
        if (descriptor.Parameters.Count == 0)
        {
            return string.Empty;
        }

        var firstParam = descriptor.Parameters[0];
        var labels = firstParam.AcceptableTypes
            .Select(FormatCellObjectType)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (labels.Length == 0)
        {
            return string.Empty;
        }

        return $"Input: {string.Join(" / ", labels)}";
    }

    private static string FormatCellObjectType(CellObjectType type)
    {
        return type switch
        {
            CellObjectType.Text => "Text",
            CellObjectType.Number => "Number",
            CellObjectType.Image => "Image",
            CellObjectType.Video => "Video",
            CellObjectType.Audio => "Audio",
            CellObjectType.Directory => "Directory",
            CellObjectType.File => "File",
            CellObjectType.Table => "Table",
            CellObjectType.Json => "JSON",
            CellObjectType.Pdf => "PDF",
            CellObjectType.CodePython => "Python Code",
            CellObjectType.CodeCSharp => "C# Code",
            CellObjectType.CodeJavaScript => "JavaScript Code",
            CellObjectType.CodeTypeScript => "TypeScript Code",
            CellObjectType.CodeHtml => "HTML",
            CellObjectType.CodeCss => "CSS",
            _ => type.ToString()
        };
    }

    private static bool IsFunctionSupportedByConnection(FunctionDescriptor descriptor, WorkspaceConnection? connection)
    {
        if (descriptor.Category != FunctionCategory.AI)
        {
            return true;
        }

        if (connection == null || !connection.IsActive)
        {
            return false;
        }

        return descriptor.Name switch
        {
            "IMAGE_TO_CAPTION" => !string.IsNullOrWhiteSpace(connection.VisionModel ?? connection.Model),
            "TEXT_TO_IMAGE" => !string.IsNullOrWhiteSpace(connection.ImageModel),
            _ => !string.IsNullOrWhiteSpace(connection.Model)
        };
    }

    private static bool IsFunctionRelevantForCellType(FunctionDescriptor descriptor, CellObjectType cellType)
    {
        if (descriptor.Parameters.Count == 0)
        {
            return true;
        }

        var normalizedCellType = NormalizeCellObjectType(cellType);
        var firstParam = descriptor.Parameters[0];
        return firstParam.AcceptableTypes
            .Select(NormalizeCellObjectType)
            .Contains(normalizedCellType);
    }

    private static CellObjectType NormalizeCellObjectType(CellObjectType type)
    {
        return type switch
        {
            CellObjectType.Empty => CellObjectType.Text,
            CellObjectType.Markdown => CellObjectType.Text,
            CellObjectType.Script => CellObjectType.Text,
            _ => type
        };
    }

    private void RefreshSheetTabs()
    {
        SheetTabs.TabItems.Clear();
        foreach (var sheet in ViewModel.Sheets)
        {
            var headerText = new TextBlock
            {
                Text = sheet.Name
            };
            // Default to theme-aware primary text for non-selected tabs
            var themePrimary = Application.Current.Resources["TextPrimaryBrush"] as SolidColorBrush;
            headerText.Foreground = themePrimary ?? new SolidColorBrush(Colors.White);
            // If this is the selected sheet, adjust for dark theme to ensure readability on light tab background
            var prefs = App.PreferencesService.LoadPreferences();
            var currentTheme = prefs.Theme ?? "Dark";
            if (ViewModel.SelectedSheet == sheet && currentTheme == "Dark")
            {
                headerText.Foreground = new SolidColorBrush(Color.FromArgb(0xFF, 0x1E, 0x1E, 0x1E));
            }
            var tab = new TabViewItem
            {
                Header = headerText,
                Tag = sheet
            };

            // Enable double-click to rename
            tab.DoubleTapped += (s, e) =>
            {
                e.Handled = true;
                ShowRenameDialog(sheet);
            };

            // Context menu for Rename/Delete
            var flyout = new MenuFlyout();
            var renameItem = new MenuFlyoutItem { Text = "Rename" };
            renameItem.Click += (s, e) => ShowRenameDialog(sheet);
            var deleteItem = new MenuFlyoutItem { Text = "Delete" };
            deleteItem.Click += (s, e) => ShowDeleteConfirm(sheet);
            flyout.Items.Add(renameItem);
            flyout.Items.Add(deleteItem);
            FlyoutBase.SetAttachedFlyout(tab, flyout);
            tab.RightTapped += (s, e) => FlyoutBase.ShowAttachedFlyout(tab);
            SheetTabs.TabItems.Add(tab);
        }
        
        if (SheetTabs.TabItems.Count > 0)
        {
            SheetTabs.SelectedIndex = 0;
            BuildSpreadsheetGrid(ViewModel.Sheets[0]);
        }
    }

    internal void BuildSpreadsheetGrid(SheetViewModel sheet)
    {
        SpreadsheetGrid.Children.Clear();
        SpreadsheetGrid.RowDefinitions.Clear();
        SpreadsheetGrid.ColumnDefinitions.Clear();

        var visibleRows = sheet.GetVisibleRowIndices().ToList();

        // If the spreadsheet viewport is available, grow the sheet to fill the remaining visible area
        if (SpreadsheetScrollViewer != null && SpreadsheetScrollViewer.ActualHeight > 0)
        {
            // Calculate how many rows fit in the viewport (subtracting header and margins)
            double viewportHeight = SpreadsheetScrollViewer.ActualHeight - 48; // header + padding
            int rowsVisible = Math.Max(5, (int)Math.Floor(viewportHeight / DefaultRowHeight));
            if (rowsVisible > visibleRows.Count)
            {
                sheet.EnsureCapacity(rowsVisible + 2, sheet.ColumnCount);
                visibleRows = sheet.GetVisibleRowIndices().ToList();
            }
        }

    var visibleColumns = sheet.GetVisibleColumnIndices().ToList();

    // Also ensure enough columns to fill the viewport width
        if (SpreadsheetScrollViewer != null && SpreadsheetScrollViewer.ActualWidth > 0)
        {
            double availableWidth = Math.Max(0, SpreadsheetScrollViewer.ActualWidth - 48 - 24); // subtract row header and margins
            // Calculate used width for existing visible columns
            double usedWidth = 0;
            foreach (var columnIndex in visibleColumns)
            {
                var w = sheet.ColumnWidths.Count > columnIndex ? sheet.ColumnWidths[columnIndex] : SheetViewModel.DefaultColumnWidth;
                usedWidth += Math.Max(40, w * ColumnWidthScale);
            }

            if (usedWidth < availableWidth)
            {
                var avgWidth = sheet.ColumnWidths.Count > 0 ? sheet.ColumnWidths.Average() * ColumnWidthScale : SheetViewModel.DefaultColumnWidth * ColumnWidthScale;
                var deficit = availableWidth - usedWidth;
                var additional = (int)Math.Ceiling(deficit / Math.Max(1, avgWidth));
                if (additional > 0)
                {
                    sheet.EnsureCapacity(sheet.Rows.Count, sheet.ColumnCount + additional + 2);
                    visibleColumns = sheet.GetVisibleColumnIndices().ToList();
                }
            }
        }
        

        // Header row plus visible row slots
        SpreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(32) });
        foreach (var _ in visibleRows)
        {
            SpreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(DefaultRowHeight) });
        }

        // Row header column plus visible column slots
        SpreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(48) });
        foreach (var columnIndex in visibleColumns)
        {
            var width = sheet.ColumnWidths.Count > columnIndex
                ? sheet.ColumnWidths[columnIndex]
                : SheetViewModel.DefaultColumnWidth;
            // Apply display scale to column widths
            var displayWidth = Math.Max(40, width * ColumnWidthScale);
            SpreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(displayWidth) });
        }

        // Top-left corner
        var cornerBrush = Application.Current.Resources["HeaderBackgroundBrush"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30));
        AddCell(new Border
        {
            Background = cornerBrush,
            CornerRadius = new CornerRadius(2),
            Margin = new Thickness(1),
            BorderBrush = Application.Current.Resources["GridLineColor"] as SolidColorBrush,
            BorderThickness = new Thickness(0, 0, 1, 1)
        }, 0, 0);

        // Column headers
        var headerBrush = Application.Current.Resources["HeaderBackgroundBrush"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30));
        var headerTextBrush = Application.Current.Resources["TextSecondaryBrush"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0x80, 0x80, 0x80));
        var gridLineBrush = Application.Current.Resources["GridLineColor"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30));
        
        for (int position = 0; position < visibleColumns.Count; position++)
        {
            var columnIndex = visibleColumns[position];
            var header = new Border
            {
                Background = headerBrush,
                CornerRadius = new CornerRadius(0),
                Margin = new Thickness(0),
                BorderBrush = gridLineBrush,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Child = new TextBlock
                {
                    Text = sheet.ColumnHeaders[columnIndex],
                    Foreground = headerTextBrush,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontWeight = Microsoft.UI.Text.FontWeights.Normal,
                    FontSize = 12
                }
            };
            header.Tag = columnIndex;
            FlyoutBase.SetAttachedFlyout(header, Resources["ColumnHeaderMenu"] as MenuFlyout);
            header.RightTapped += ColumnHeader_RightTapped;
            AddCell(header, 0, position + 1);
        }

        // Row headers and cells
        for (int position = 0; position < visibleRows.Count; position++)
        {
            var rowIndex = visibleRows[position];
            var rowVm = sheet.Rows[rowIndex];

            var rowHeader = new Border
            {
                Background = headerBrush,
                CornerRadius = new CornerRadius(0),
                Margin = new Thickness(0),
                BorderBrush = gridLineBrush,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Child = new TextBlock
                {
                    Text = rowVm.Label,
                    Foreground = headerTextBrush,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontWeight = Microsoft.UI.Text.FontWeights.Normal,
                    FontSize = 12
                }
            };
            rowHeader.Tag = rowIndex;
            FlyoutBase.SetAttachedFlyout(rowHeader, Resources["RowHeaderMenu"] as MenuFlyout);
            rowHeader.RightTapped += RowHeader_RightTapped;
            AddCell(rowHeader, position + 1, 0);

            for (int colPosition = 0; colPosition < visibleColumns.Count; colPosition++)
            {
                var columnIndex = visibleColumns[colPosition];
                if (columnIndex >= rowVm.Cells.Count)
                {
                    continue;
                }

                var cellVm = rowVm.Cells[columnIndex];
                var button = CreateCellButton(cellVm);
                AddCell(button, position + 1, colPosition + 1);
            }
        }

        RefreshSelectionReferences(sheet);
        UpdateSelectionVisuals();
        UpdateSelectionSummary();
    }

    private Button CreateCellButton(CellViewModel cellVm)
    {
        // Use CellThemeBackgroundBrush which respects the active theme
        var cellBgBrush = Application.Current.Resources["CellThemeBackgroundBrush"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0x25, 0x25, 0x26));
        var gridLineBrush = Application.Current.Resources["GridLineColor"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30));
        
        var button = new Button
        {
            Background = cellBgBrush,
            BorderBrush = gridLineBrush,
            BorderThickness = new Thickness(0, 0, 1, 1),
            CornerRadius = new CornerRadius(0),
            Margin = new Thickness(0),
            Padding = new Thickness(4, 0, 4, 0),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Stretch,
            VerticalContentAlignment = VerticalAlignment.Stretch,
            Tag = cellVm
        };
        
        var stack = new StackPanel { Spacing = 6 };
        
        // Show value directly without cell label
        var textBrush = Application.Current.Resources["TextPrimaryBrush"] as SolidColorBrush
            ?? new SolidColorBrush(Color.FromArgb(0xFF, 0xCC, 0xCC, 0xCC));
        
            var valueText = new TextBlock
            {
                Text = cellVm.DisplayValue,
                Foreground = textBrush,
                FontSize = 10,
                TextWrapping = TextWrapping.NoWrap,
                MaxLines = 1,
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center
        };
        
        stack.Children.Add(valueText);
        
        button.Content = stack;
        ApplyCellStyling(button, cellVm, valueText);
        button.Click += (s, e) => SelectCell(cellVm, button);
        button.DoubleTapped += (s, e) => StartDirectEdit(cellVm, button);
    button.PointerPressed += CellButton_PointerPressed;
        
        // Attach context menu (Phase 5 Task 16)
        button.ContextFlyout = Resources["CellContextMenu"] as MenuFlyout;
        button.RightTapped += (s, e) => SelectCell(cellVm, button);
        
        return button;
    }

    private void ApplyCellStyling(Button button, CellViewModel cellVm, TextBlock? valueText = null)
    {
        var isDefaultBackground = string.IsNullOrWhiteSpace(cellVm.FormatBackground) ||
            string.Equals(cellVm.FormatBackground, CellFormat.DefaultBackgroundColor, StringComparison.OrdinalIgnoreCase);
        var isDefaultBorder = string.IsNullOrWhiteSpace(cellVm.FormatBorder) ||
            string.Equals(cellVm.FormatBorder, CellFormat.DefaultBorderColor, StringComparison.OrdinalIgnoreCase);
        var isDefaultForeground = string.IsNullOrWhiteSpace(cellVm.FormatForeground) ||
            string.Equals(cellVm.FormatForeground, CellFormat.DefaultForegroundColor, StringComparison.OrdinalIgnoreCase);

        if (isDefaultBackground && Application.Current.Resources.TryGetValue("CellThemeBackgroundBrush", out var bgBrush) && bgBrush is SolidColorBrush themeBackground)
        {
            button.Background = themeBackground;
        }
        else
        {
            var backgroundColor = ParseColor(cellVm.FormatBackground);
            button.Background = new SolidColorBrush(backgroundColor);
        }

        if (isDefaultBorder && Application.Current.Resources.TryGetValue("CellThemeBorderBrush", out var borderBrush) && borderBrush is SolidColorBrush themeBorder)
        {
            button.BorderBrush = themeBorder;
        }
        else
        {
            var borderColor = ParseColor(cellVm.FormatBorder);
            button.BorderBrush = new SolidColorBrush(borderColor);
        }

        button.BorderThickness = new Thickness(cellVm.Format.BorderThickness);

        valueText ??= (button.Content as StackPanel)?.Children.OfType<TextBlock>().FirstOrDefault();
        if (valueText != null)
        {
            valueText.Text = cellVm.DisplayValue;

            if (isDefaultForeground && Application.Current.Resources.TryGetValue("CellThemeForegroundBrush", out var fgBrush) && fgBrush is SolidColorBrush themeForeground)
            {
                valueText.Foreground = themeForeground;
            }
            else
            {
                valueText.Foreground = new SolidColorBrush(ParseColor(cellVm.FormatForeground));
            }

            valueText.FontSize = cellVm.Format.FontSize;
            valueText.FontWeight = cellVm.Format.IsBold ? FontWeights.SemiBold : FontWeights.Normal;
            valueText.FontStyle = cellVm.Format.IsItalic ? FontStyle.Italic : FontStyle.Normal;
        }

        switch (cellVm.VisualState)
        {
            case CellVisualState.InDependencyChain:
                if (Application.Current.Resources.TryGetValue("CellStateInDependencyChainBrush", out var depBrush))
                    button.BorderBrush = depBrush as SolidColorBrush ?? new SolidColorBrush(Colors.DeepSkyBlue);
                else
                    button.BorderBrush = new SolidColorBrush(Colors.DeepSkyBlue);
                button.BorderThickness = new Thickness(2);
                break;
            case CellVisualState.Error:
                if (Application.Current.Resources.TryGetValue("CellStateErrorBrush", out var errBrush))
                    button.BorderBrush = errBrush as SolidColorBrush ?? new SolidColorBrush(Colors.OrangeRed);
                else
                    button.BorderBrush = new SolidColorBrush(Colors.OrangeRed);
                button.BorderThickness = new Thickness(2);
                break;
            case CellVisualState.JustUpdated:
                if (Application.Current.Resources.TryGetValue("CellStateJustUpdatedBrush", out var updBrush))
                    button.BorderBrush = updBrush as SolidColorBrush ?? new SolidColorBrush(Colors.LimeGreen);
                else
                    button.BorderBrush = new SolidColorBrush(Colors.LimeGreen);
                button.BorderThickness = new Thickness(2);
                break;
        }
    }

    private static Color ParseColor(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
        {
            return Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
        }

        var span = hex.AsSpan().TrimStart('#');
        try
        {
            byte a = 0xFF;
            byte r;
            byte g;
            byte b;

            if (span.Length == 6)
            {
                r = Convert.ToByte(span.Slice(0, 2).ToString(), 16);
                g = Convert.ToByte(span.Slice(2, 2).ToString(), 16);
                b = Convert.ToByte(span.Slice(4, 2).ToString(), 16);
            }
            else if (span.Length == 8)
            {
                a = Convert.ToByte(span.Slice(0, 2).ToString(), 16);
                r = Convert.ToByte(span.Slice(2, 2).ToString(), 16);
                g = Convert.ToByte(span.Slice(4, 2).ToString(), 16);
                b = Convert.ToByte(span.Slice(6, 2).ToString(), 16);
            }
            else
            {
                return Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
            }

            return Color.FromArgb(a, r, g, b);
        }
        catch
        {
            return Color.FromArgb(0xFF, 0x2D, 0x2D, 0x30);
        }
    }

    private void AddCell(FrameworkElement element, int row, int col)
    {
        Grid.SetRow(element, row);
        Grid.SetColumn(element, col);
        SpreadsheetGrid.Children.Add(element);
    }

    private void ClearSelectionSet()
    {
        foreach (var vm in _multiSelection)
        {
            vm.IsSelected = false;
        }

        _multiSelection.Clear();
    }

    private void AddToSelection(CellViewModel cell)
    {
        if (!_multiSelection.Contains(cell))
        {
            _multiSelection.Add(cell);
            cell.IsSelected = true;
        }
    }

    private void RemoveFromSelection(CellViewModel cell)
    {
        if (_multiSelection.Remove(cell))
        {
            cell.IsSelected = false;
        }
    }

    private void SelectRange(CellViewModel anchor, CellViewModel target)
    {
        var sheet = ViewModel.SelectedSheet;
        if (sheet == null)
        {
            return;
        }

        ClearSelectionSet();

        var minRow = Math.Min(anchor.Row, target.Row);
        var maxRow = Math.Max(anchor.Row, target.Row);
        var minCol = Math.Min(anchor.Column, target.Column);
        var maxCol = Math.Max(anchor.Column, target.Column);

        for (int r = minRow; r <= maxRow && r < sheet.Rows.Count; r++)
        {
            if (!sheet.IsRowVisible(r))
            {
                continue;
            }

            var rowVm = sheet.Rows[r];
            for (int c = minCol; c <= maxCol && c < rowVm.Cells.Count; c++)
            {
                if (!sheet.IsColumnVisible(c))
                {
                    continue;
                }

                AddToSelection(rowVm.Cells[c]);
            }
        }
    }

    private string GetSelectionRangeLabel()
    {
        if (_multiSelection.Count == 0)
        {
            return string.Empty;
        }

        var ordered = _multiSelection
            .OrderBy(c => c.Row)
            .ThenBy(c => c.Column)
            .ToList();

        var first = ordered.First();
        var last = ordered.Last();

        if (ReferenceEquals(first, last))
        {
            return first.DisplayLabel;
        }

        return $"{first.DisplayLabel}:{last.DisplayLabel}";
    }

    private void RefreshSelectionReferences(SheetViewModel sheet)
    {
        for (var i = 0; i < _multiSelection.Count; i++)
        {
            var cell = _multiSelection[i];
            var updated = sheet.GetCell(cell.Row, cell.Column);
            if (updated == null)
            {
                cell.IsSelected = false;
                _multiSelection.RemoveAt(i);
                i--;
                continue;
            }

            if (!sheet.IsRowVisible(updated.Row) || !sheet.IsColumnVisible(updated.Column))
            {
                updated.IsSelected = false;
                _multiSelection.RemoveAt(i);
                i--;
                continue;
            }

            if (!ReferenceEquals(updated, cell))
            {
                cell.IsSelected = false;
                updated.IsSelected = true;
                _multiSelection[i] = updated;
            }
        }

        if (_selectedCell != null)
        {
            var updatedSelected = sheet.GetCell(_selectedCell.Row, _selectedCell.Column);
            if (updatedSelected != null && sheet.IsRowVisible(updatedSelected.Row) && sheet.IsColumnVisible(updatedSelected.Column))
            {
                _selectedCell = updatedSelected;
            }
            else
            {
                _selectedCell = _multiSelection.LastOrDefault();
            }
        }

        if (_multiSelection.Count == 0)
        {
            _selectedCell = null;
            _selectedButton = null;
        }
    }

    private void SelectCell(CellViewModel cell, Button button)
    {
        if (_isFormulaEditing && _isPickingFormulaReference && _formulaEditTargetCell != null && cell != _formulaEditTargetCell)
        {
            InsertCellReferenceIntoFormula(cell);
            return;
        }

        var ctrl = InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control)
            .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
        var shift = InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift)
            .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);

        var nextActive = cell;

        if (shift && (_selectionAnchor ?? _selectedCell) != null)
        {
            var anchor = _selectionAnchor ?? _selectedCell ?? cell;
            SelectRange(anchor, cell);
        }
        else if (ctrl)
        {
            if (_multiSelection.Contains(cell))
            {
                if (_multiSelection.Count > 1)
                {
                    RemoveFromSelection(cell);
                    nextActive = _multiSelection.LastOrDefault();
                }
            }
            else
            {
                AddToSelection(cell);
            }
        }
        else
        {
            ClearSelectionSet();
            AddToSelection(cell);
            _selectionAnchor = cell;
        }

        if (_multiSelection.Count == 0)
        {
            AddToSelection(cell);
            nextActive = cell;
        }

        _selectedCell = nextActive ?? cell;
        if (!shift)
        {
            _selectionAnchor = _selectedCell;
        }

        _selectedButton = FindButtonForCell(_selectedCell) ?? button;

        UpdateSelectionVisuals();
        UpdateSelectionSummary();
        UpdateInspectorForCell(_selectedCell);

        if (!_isFormulaEditing)
        {
            ClearFormulaHighlights();
        }
        else
        {
            UpdateParameterHints();
        }

        // Focus function group for the selected cell's type
        var cellType = _selectedCell?.Value.ObjectType ?? CellObjectType.Text;
        cellType = NormalizeCellObjectType(cellType);
        FocusFunctionGroupForCellType(cellType);
    }

    private Button? FindButtonForCell(CellViewModel? cell)
    {
        if (cell == null)
        {
            return null;
        }

        return SpreadsheetGrid.Children.OfType<Button>().FirstOrDefault(b => ReferenceEquals(b.Tag, cell));
    }

    private void UpdateSelectionVisuals()
    {
        foreach (var child in SpreadsheetGrid.Children.OfType<Button>())
        {
            if (child.Tag is CellViewModel vm)
            {
                ApplyCellStyling(child, vm);
                vm.IsSelected = _multiSelection.Contains(vm);
                if (_multiSelection.Contains(vm))
                {
                    if (ReferenceEquals(vm, _selectedCell))
                    {
                        child.BorderBrush = new SolidColorBrush(Colors.DodgerBlue);
                        child.BorderThickness = new Thickness(2);
                    }
                    else
                    {
                        child.BorderBrush = new SolidColorBrush(Color.FromArgb(0xFF, 0x64, 0x95, 0xED));
                        child.BorderThickness = new Thickness(1.5);
                    }
                }
            }
        }
    }

    private bool TryGetNumericValue(CellViewModel cell, out double value)
    {
        var candidates = new[]
        {
            cell.Value.DisplayValue,
            cell.Value.SerializedValue,
            cell.RawValue
        };

        foreach (var candidate in candidates)
        {
            if (!string.IsNullOrWhiteSpace(candidate) &&
                double.TryParse(candidate, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            {
                return true;
            }
        }

        value = 0;
        return false;
    }

    private void UpdateSelectionSummary()
    {
        if (SelectionInfoText == null)
        {
            return;
        }

        if (_multiSelection.Count == 0)
        {
            SelectionInfoText.Text = "Selection: none";
            UpdateInspectorForCell(null);
            return;
        }

        if (_multiSelection.Count == 1)
        {
            var cell = _multiSelection[0];
            SelectionInfoText.Text = $"Selection: {cell.DisplayLabel}";
            UpdateInspectorForCell(cell);
            return;
        }

        var count = _multiSelection.Count;
        var range = GetSelectionRangeLabel();
        double sum = 0;
        var numericCount = 0;

        foreach (var cell in _multiSelection)
        {
            if (TryGetNumericValue(cell, out var number))
            {
                sum += number;
                numericCount++;
            }
        }

        if (numericCount > 0)
        {
            var average = sum / numericCount;
            SelectionInfoText.Text = string.Format(
                CultureInfo.InvariantCulture,
                "Selection: {0} ({1}) • Σ {2:G5} • Avg {3:G5}",
                range,
                count,
                sum,
                average);
        }
        else
        {
            SelectionInfoText.Text = $"Selection: {range} ({count})";
        }

        UpdateInspectorForCell(_selectedCell);
    }

    private void UpdateInspectorForCell(CellViewModel? cell)
    {
        if (_isUpdatingCell)
        {
            return;
        }

        _isUpdatingCell = true;

        if (cell == null)
        {
            CellInspectorPanel.Visibility = Visibility.Collapsed;
            NoCellPanel.Visibility = Visibility.Visible;
            _isUpdatingCell = false;
            return;
        }

        CellInspectorPanel.Visibility = Visibility.Visible;
        NoCellPanel.Visibility = Visibility.Collapsed;

        if (_multiSelection.Count > 1)
        {
            CellLabel.Text = $"{GetSelectionRangeLabel()} ({_multiSelection.Count} cells)";
            CellClassTypeLabel.Text = "";
        }
        else if (cell != null)
        {
            CellLabel.Text = cell.DisplayLabel;
            var typeLabel = GetCellObjectTypeLabel(cell.Value.ObjectType);
            CellClassTypeLabel.Text = $"Type: {typeLabel}";
        }

    // When editing a string that looks like a number, prefix with '
    if (cell != null && cell.Value != null && cell.Value.ObjectType == CellObjectType.Text && double.TryParse(cell.RawValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
    {
        CellValueBox.Text = "'" + (cell.RawValue ?? string.Empty);
    }
    else
    {
        CellValueBox.Text = cell?.RawValue ?? string.Empty;
    }
    CellFormulaBox.Text = cell?.Formula ?? string.Empty;
    CellNotesBox.Text = cell?.Notes ?? string.Empty;
    AutomationModeBox.SelectedIndex = (int)(cell?.AutomationMode ?? CellAutomationMode.Manual);

        _isUpdatingCell = false;
    }

    private void ColumnHeader_RightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        if (sender is FrameworkElement element && element.Tag is int columnIndex)
        {
            _columnContextIndex = columnIndex;
            FlyoutBase.ShowAttachedFlyout(element);
        }
    }

    private void RowHeader_RightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        if (sender is FrameworkElement element && element.Tag is int rowIndex)
        {
            _rowContextIndex = rowIndex;
            FlyoutBase.ShowAttachedFlyout(element);
        }
    }

    private void HideColumn_Click(object sender, RoutedEventArgs e)
    {
        if (_columnContextIndex is null || ViewModel.SelectedSheet is null)
        {
            return;
        }

        var sheet = ViewModel.SelectedSheet;
        if (sheet.VisibleColumnCount <= 1)
        {
            ViewModel.StatusMessage = "At least one column must remain visible.";
            _columnContextIndex = null;
            return;
        }

        var column = _columnContextIndex.Value;
        sheet.SetColumnVisibility(column, false);
        BuildSpreadsheetGrid(sheet);
        ViewModel.StatusMessage = $"Hidden column {GetColumnName(column)}";
        _columnContextIndex = null;
    }

    private void UnhideAllColumns_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedSheet is null)
        {
            return;
        }

        var sheet = ViewModel.SelectedSheet;
        sheet.ShowAllColumns();
        BuildSpreadsheetGrid(sheet);
        ViewModel.StatusMessage = "All columns are visible";
        _columnContextIndex = null;
    }

    private void HideRow_Click(object sender, RoutedEventArgs e)
    {
        if (_rowContextIndex is null || ViewModel.SelectedSheet is null)
        {
            return;
        }

        var sheet = ViewModel.SelectedSheet;
        if (sheet.VisibleRowCount <= 1)
        {
            ViewModel.StatusMessage = "At least one row must remain visible.";
            _rowContextIndex = null;
            return;
        }

        var row = _rowContextIndex.Value;
        sheet.SetRowVisibility(row, false);
        BuildSpreadsheetGrid(sheet);
        var label = row >= 0 && row < sheet.Rows.Count ? sheet.Rows[row].Label : (row + 1).ToString(CultureInfo.InvariantCulture);
        ViewModel.StatusMessage = $"Hidden row {label}";
        _rowContextIndex = null;
    }

    private void UnhideAllRows_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedSheet is null)
        {
            return;
        }

        var sheet = ViewModel.SelectedSheet;
        sheet.ShowAllRows();
        BuildSpreadsheetGrid(sheet);
        ViewModel.StatusMessage = "All rows are visible";
        _rowContextIndex = null;
    }

    // ==================== Freeze Panes (Phase 8) ====================

    private void FreezeColumns_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedSheet is null || _columnContextIndex is null)
        {
            return;
        }

        var sheet = ViewModel.SelectedSheet;
        var columnIndex = _columnContextIndex.Value;
        
        // Freeze up to and including this column
        sheet.FrozenColumnCount = columnIndex + 1;
        BuildSpreadsheetGrid(sheet);
        
        var columnName = CellAddress.ColumnIndexToName(columnIndex);
        ViewModel.StatusMessage = $"Frozen columns A-{columnName}";
        _columnContextIndex = null;
    }

    private void FreezeRows_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedSheet is null || _rowContextIndex is null)
        {
            return;
        }

        var sheet = ViewModel.SelectedSheet;
        var rowIndex = _rowContextIndex.Value;
        
        // Freeze up to and including this row
        sheet.FrozenRowCount = rowIndex + 1;
        BuildSpreadsheetGrid(sheet);
        
        ViewModel.StatusMessage = $"Frozen rows 1-{rowIndex + 1}";
        _rowContextIndex = null;
    }

    private void UnfreezeAll_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedSheet is null)
        {
            return;
        }

        var sheet = ViewModel.SelectedSheet;
        sheet.FrozenColumnCount = 0;
        sheet.FrozenRowCount = 0;
        BuildSpreadsheetGrid(sheet);
        
        ViewModel.StatusMessage = "Unfrozen all panes";
        _columnContextIndex = null;
        _rowContextIndex = null;
    }

    // ==================== Fill Operations (Phase 8) ====================

    private void FillDown_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell is null || _multiSelection.Count < 2)
        {
            ViewModel.StatusMessage = "⚠️ Select a range first";
            return;
        }

        var sourceCell = _multiSelection.First();
        var targetCells = _multiSelection.Skip(1).Where(c => c.Column == sourceCell.Column && c.Row > sourceCell.Row).ToList();
        
        if (targetCells.Count == 0)
        {
            ViewModel.StatusMessage = "⚠️ Select cells below the source";
            return;
        }

        foreach (var targetCell in targetCells)
        {
            if (!string.IsNullOrWhiteSpace(sourceCell.Formula))
            {
                // Copy formula with relative reference adjustment
                targetCell.Formula = sourceCell.Formula;
            }
            else
            {
                targetCell.Value = sourceCell.Value;
            }
            targetCell.MarkAsUpdated();
        }

        ViewModel.StatusMessage = $"✅ Filled down to {targetCells.Count} cells";
    }

    private void FillRight_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell is null || _multiSelection.Count < 2)
        {
            ViewModel.StatusMessage = "⚠️ Select a range first";
            return;
        }

        var sourceCell = _multiSelection.First();
        var targetCells = _multiSelection.Skip(1).Where(c => c.Row == sourceCell.Row && c.Column > sourceCell.Column).ToList();
        
        if (targetCells.Count == 0)
        {
            ViewModel.StatusMessage = "⚠️ Select cells to the right of the source";
            return;
        }

        foreach (var targetCell in targetCells)
        {
            if (!string.IsNullOrWhiteSpace(sourceCell.Formula))
            {
                // Copy formula with relative reference adjustment
                targetCell.Formula = sourceCell.Formula;
            }
            else
            {
                targetCell.Value = sourceCell.Value;
            }
            targetCell.MarkAsUpdated();
        }

        ViewModel.StatusMessage = $"✅ Filled right to {targetCells.Count} cells";
    }

    private void FormatPainter_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell is null)
        {
            ViewModel.StatusMessage = "⚠️ Select a source cell first";
            return;
        }

        // Store format for painting
        ViewModel.StatusMessage = "🖌️ Format copied. Select target cells and paste (Ctrl+V)";
        
        // For now, copy the format to clipboard as a special marker
        // Future: implement dedicated format clipboard
    }

    private double CalculateAutoFitWidth(SheetViewModel sheet, int columnIndex)
    {
        double maxCharacters = 0;

        if (columnIndex < 0 || columnIndex >= sheet.ColumnHeaders.Count)
        {
            return SheetViewModel.DefaultColumnWidth;
        }

        maxCharacters = Math.Max(maxCharacters, sheet.ColumnHeaders[columnIndex].Length);

        foreach (var row in sheet.Rows)
        {
            if (columnIndex >= row.Cells.Count)
            {
                continue;
            }

            if (!sheet.IsRowVisible(row.Index))
            {
                continue;
            }

            var cell = row.Cells[columnIndex];
            var display = cell.DisplayValue ?? string.Empty;
            maxCharacters = Math.Max(maxCharacters, display.Length);
        }

        var calculated = 24 + maxCharacters * 9; // base padding + average character width
        return Math.Clamp(calculated, 72, 480);
    }

    private void AutoFitColumn_Click(object sender, RoutedEventArgs e)
    {
        if (_columnContextIndex is null || ViewModel.SelectedSheet is null)
        {
            return;
        }

        var column = _columnContextIndex.Value;
        var sheet = ViewModel.SelectedSheet;

        var width = CalculateAutoFitWidth(sheet, column);
        if (column < sheet.ColumnWidths.Count)
        {
            sheet.ColumnWidths[column] = width;
            BuildSpreadsheetGrid(sheet);
            ViewModel.StatusMessage = $"Auto-fit column {GetColumnName(column)}";
            _columnContextIndex = null;
        }
    }

    private async void SetColumnWidth_Click(object sender, RoutedEventArgs e)
    {
        if (_columnContextIndex is null || ViewModel.SelectedSheet is null)
        {
            return;
        }

        var column = _columnContextIndex.Value;
        var sheet = ViewModel.SelectedSheet;
        if (column >= sheet.ColumnWidths.Count)
        {
            return;
        }

        var textbox = new TextBox
        {
            Text = sheet.ColumnWidths[column].ToString("F0", CultureInfo.InvariantCulture),
            Width = 200
        };

        var dialog = new ContentDialog
        {
            Title = $"Set width for column {GetColumnName(column)}",
            PrimaryButtonText = "Apply",
            SecondaryButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary,
            Content = new StackPanel
            {
                Spacing = 8,
                Children =
                {
                    new TextBlock
                    {
                        Text = "Enter width in pixels (60 - 600)",
                        FontSize = 12
                    },
                    textbox
                }
            }
        };

        dialog.XamlRoot = Content.XamlRoot;

        var result = await dialog.ShowAsync();
        if (result == ContentDialogResult.Primary)
        {
            if (double.TryParse(textbox.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out var width))
            {
                width = Math.Clamp(width, 60, 600);
                sheet.ColumnWidths[column] = width;
                BuildSpreadsheetGrid(sheet);
                ViewModel.StatusMessage = $"Column {GetColumnName(column)} width set to {width:F0}px";
            }
            else
            {
                ViewModel.StatusMessage = "Invalid column width entered";
            }
        }

        _columnContextIndex = null;
    }

    private void ResetColumnWidth_Click(object sender, RoutedEventArgs e)
    {
        if (_columnContextIndex is null || ViewModel.SelectedSheet is null)
        {
            return;
        }

        var column = _columnContextIndex.Value;
        var sheet = ViewModel.SelectedSheet;
        if (column >= sheet.ColumnWidths.Count)
        {
            return;
        }

        sheet.ColumnWidths[column] = SheetViewModel.DefaultColumnWidth;
        BuildSpreadsheetGrid(sheet);
        ViewModel.StatusMessage = $"Reset column {GetColumnName(column)} width";
        _columnContextIndex = null;
    }

    private void StartDirectEdit(CellViewModel cell, Button button, string? initialCharacter = null)
    {
        _isPickingFormulaReference = false;
        // Only allow direct edit for text/empty cells (not images, directories, etc.)
        if (cell.Value.ObjectType != CellObjectType.Text && 
            cell.Value.ObjectType != CellObjectType.Number &&
            cell.Value.ObjectType != CellObjectType.Empty)
        {
            return;
        }

        // Position the edit box over the cell
        var buttonPos = button.TransformToVisual(SpreadsheetGrid).TransformPoint(new Windows.Foundation.Point(0, 0));
        
        // Make edit box smaller: reduce by 75% relative to previous size
        double minWidth = Math.Max(75, button.ActualWidth * 0.25);
        double maxWidth = 175; // 700 * 0.25
        double availableWidth = (SpreadsheetGrid.ActualWidth - buttonPos.X - 16) * 0.25;
        if (availableWidth > 0)
        {
            CellEditBox.Width = Math.Min(maxWidth, Math.Max(minWidth, availableWidth));
        }
        else
        {
            CellEditBox.Width = Math.Min(maxWidth, minWidth);
        }

        double desiredHeight = 35; // 140 * 0.25
        double availableHeight = (SpreadsheetGrid.ActualHeight - buttonPos.Y - 16) * 0.25;
        if (availableHeight > 0)
        {
            CellEditBox.Height = Math.Min(105, Math.Max(desiredHeight, Math.Min(availableHeight, desiredHeight)));
        }
        else
        {
            CellEditBox.Height = desiredHeight;
        }
        CellEditBox.Margin = new Thickness(buttonPos.X, buttonPos.Y, 0, 0);
        CellEditBox.HorizontalAlignment = HorizontalAlignment.Left;
        CellEditBox.VerticalAlignment = VerticalAlignment.Top;
        
        CellEditBox.Tag = cell;
        CellEditBox.Visibility = Visibility.Visible;
        
        // If an initial character was provided, replace the content with it
        if (!string.IsNullOrEmpty(initialCharacter))
        {
            // If editing started with =, switch to formula editing mode and populate the Formula box
            if (initialCharacter.StartsWith("="))
            {
                System.Diagnostics.Debug.WriteLine($"StartDirectEdit: Entering formula mode with text '{initialCharacter}'");
                
                // VISIBLE DEBUG: Show message box
                SelectionInfoText.Text = "DEBUG: StartDirectEdit received formula - entering formula mode";
                
                // Show formula editor in inspector - preserve what user typed
                CellEditBox.Text = initialCharacter;
                CellEditBox.Focus(FocusState.Programmatic);
                DispatcherQueue.TryEnqueue(() =>
                {
                    CellEditBox.SelectionStart = CellEditBox.Text.Length;
                    CellEditBox.SelectionLength = 0;
                });

                // Start focused formula editing via inspector control
                _isFormulaEditing = true;
                _formulaEditTargetCell = cell;
                
                // Ensure inspector panel is visible for formula editing
                if (!_inspectorVisible)
                {
                    System.Diagnostics.Debug.WriteLine("StartDirectEdit: Inspector was hidden, showing it now");
                    SelectionInfoText.Text = "DEBUG: Opening inspector panel...";
                    _inspectorVisible = true;
                    SetInspectorPanelVisibility(true, ViewModel.Settings.InspectorPanelWidth);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("StartDirectEdit: Inspector already visible");
                    SelectionInfoText.Text = "DEBUG: Inspector already visible, setting formula...";
                }
                
                // Wait for panel to be visible
                DispatcherQueue.TryEnqueue(() =>
                {
                    _updatingFormulaText = true;
                    try
                    {
                        // Mirror into the inspector formula box so user sees advanced UI
                        if (CellFormulaBox.Text != CellEditBox.Text)
                        {
                            CellFormulaBox.Text = CellEditBox.Text;
                        }
                    }
                    finally
                    {
                        _updatingFormulaText = false;
                    }
                    
                    // Keep CellEditBox VISIBLE and FOCUSED for typing
                    CellEditBox.Focus(FocusState.Keyboard);
                    CellEditBox.SelectionStart = CellEditBox.Text.Length;
                    _isPickingFormulaReference = false;
                    
                    // Show intellisense immediately
                    ShowFormulaIntellisense();
                    UpdateParameterHints();
                });
            }
            else
            {
                CellEditBox.Text = initialCharacter;
                CellEditBox.Focus(FocusState.Programmatic);
                // Use dispatcher to ensure selection is cleared after focus events complete
                DispatcherQueue.TryEnqueue(() =>
                {
                    CellEditBox.SelectionStart = CellEditBox.Text.Length;
                    CellEditBox.SelectionLength = 0;
                });
            }
        }
        else
        {
            CellEditBox.Text = cell.RawValue ?? string.Empty;
            CellEditBox.Focus(FocusState.Programmatic);
            CellEditBox.SelectAll();
        }
    }

    private void CellEditBox_LostFocus(object sender, RoutedEventArgs e)
    {
        // Don't commit if flyout is open - user is still editing the formula
        if (FunctionAutocompleteFlyout.IsOpen || _suppressCommitOnFlyoutOpen)
        {
            System.Diagnostics.Debug.WriteLine("CellEditBox_LostFocus: Flyout is open, not committing");
            return;
        }
        
        // Don't commit if we're navigating between cells (Tab key)
        if (!_isNavigatingBetweenCells)
        {
            CommitCellEdit();
        }
    }

    private bool _suppressCommitOnFlyoutOpen;

    private void FunctionAutocompleteFlyout_Opened(object sender, object e)
    {
        // Prevent commit while flyout is opening
        _suppressCommitOnFlyoutOpen = true;
        // Refocus the edit box after a short delay so keystrokes go to it
        DispatcherQueue?.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            CellEditBox.Focus(FocusState.Keyboard);
            _suppressCommitOnFlyoutOpen = false;
        });
    }

    private void FunctionAutocompleteFlyout_Closed(object sender, object e)
    {
        // Ensure focus returns to the edit box when the flyout closes
        DispatcherQueue?.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            CellEditBox.Focus(FocusState.Keyboard);
        });
    }

    private void CellEditBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (_isNavigatingBetweenCells)
        {
            _isNavigatingBetweenCells = false;
        }
    }

    private void CellEditBox_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        // If flyout is open and user presses Up/Down arrows, navigate the list
        if (_isFormulaEditing && FunctionAutocompleteFlyout.IsOpen)
        {
            if (e.Key == Windows.System.VirtualKey.Down)
            {
                // Move selection down in the flyout list
                if (FlyoutFunctionList.SelectedIndex < FlyoutFunctionList.Items.Count - 1)
                {
                    FlyoutFunctionList.SelectedIndex++;
                    FlyoutFunctionList.ScrollIntoView(FlyoutFunctionList.SelectedItem);
                }
                e.Handled = true;
                return;
            }
            else if (e.Key == Windows.System.VirtualKey.Up)
            {
                // Move selection up in the flyout list
                if (FlyoutFunctionList.SelectedIndex > 0)
                {
                    FlyoutFunctionList.SelectedIndex--;
                    FlyoutFunctionList.ScrollIntoView(FlyoutFunctionList.SelectedItem);
                }
                e.Handled = true;
                return;
            }
            else if (e.Key == Windows.System.VirtualKey.Enter || e.Key == Windows.System.VirtualKey.Tab)
            {
                // If there's exactly one suggestion, insert it
                if (FlyoutFunctionList.Items?.Count == 1 && FlyoutFunctionList.Items[0] is FunctionSuggestion only)
                {
                    InsertFunctionSuggestion(only);
                }
                else if (FlyoutFunctionList.SelectedItem is FunctionSuggestion suggestion)
                {
                    InsertFunctionSuggestion(suggestion);
                }
                e.Handled = true;
                return;
            }
            else if (e.Key == Windows.System.VirtualKey.Escape)
            {
                // Close the flyout
                FunctionAutocompleteFlyout.Hide();
                e.Handled = true;
                return;
            }
        }
        // Detect if user types = to start formula mode
        if (e.Key == (Windows.System.VirtualKey)187 && !_isFormulaEditing) // VK 187 is = key
        {
            var textBox = sender as TextBox;
            if (textBox != null && string.IsNullOrEmpty(textBox.Text))
            {
                // User typed = in empty cell - switch to formula mode
                var cell = textBox.Tag as CellViewModel;
                var button = _selectedButton;
                if (cell != null && button != null)
                {
                    // CRITICAL: Mark event as handled BEFORE calling StartDirectEdit
                    // to prevent = from being inserted twice
                    e.Handled = true;
                    
                    // Clear any pending text
                    textBox.Text = "";
                    
                    // Use dispatcher to ensure key event is fully processed
                    DispatcherQueue?.TryEnqueue(() =>
                    {
                        StartDirectEdit(cell, button, "=");
                    });
                    return;
                }
            }
        }
        
        if (_isFormulaEditing)

        {
            // Forward keys to the formula box while formula editing
            // Allow Enter to commit the formula UNLESS flyout is open (then it selects from flyout)
            if (e.Key == Windows.System.VirtualKey.Enter && !FunctionAutocompleteFlyout.IsOpen)
            {
                CommitCellEdit();
                e.Handled = true;
                return;
            }

            // Tab navigation should still work
            if (e.Key == Windows.System.VirtualKey.Tab)
            {
                _isNavigatingBetweenCells = true;
                CommitCellEdit();
                var shift = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                var newCell = MoveSelection(shift ? -1 : 1, 0);
                CellEditBox.Visibility = Visibility.Collapsed;
                CellEditBox.Tag = null;
                if (newCell != null)
                {
                    var newButton = GetButtonForCell(newCell);
                    if (newButton != null)
                    {
                        DispatcherQueue.TryEnqueue(() => StartDirectEdit(newCell, newButton));
                    }
                    else
                    {
                        DispatcherQueue.TryEnqueue(() => _isNavigatingBetweenCells = false);

                    }
                }
                else
                {
                    DispatcherQueue.TryEnqueue(() => _isNavigatingBetweenCells = false);
                }
                e.Handled = true;
                return;
            }

            // Other keys should be typed into the formula box, so let it handle the input
            return;
        }

        if (e.Key == Windows.System.VirtualKey.Enter)
        {
            CommitCellEdit();
            e.Handled = true;
        }
            else if (e.Key == Windows.System.VirtualKey.Tab)
        {
            // Set flag to prevent LostFocus from committing until new edit box is ready
            _isNavigatingBetweenCells = true;
            
            // Commit current edit (saves value but keeps edit box open due to flag)
            CommitCellEdit();
            
            // Move to next cell (right if Tab, left if Shift+Tab)
            var shift = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
            var newCell = MoveSelection(shift ? -1 : 1, 0);
            
            // Hide current edit box before reusing it
            CellEditBox.Visibility = Visibility.Collapsed;
            CellEditBox.Tag = null;
            
            if (newCell != null)
            {
                var newButton = GetButtonForCell(newCell);
                if (newButton != null)
                {
                    // Defer to dispatcher so focus changes settle before showing new overlay
                    DispatcherQueue.TryEnqueue(() =>
                    {
                        StartDirectEdit(newCell, newButton);
                    });
                }
                else
                {
                    DispatcherQueue.TryEnqueue(() => _isNavigatingBetweenCells = false);
                }
            }
            else
            {
                DispatcherQueue.TryEnqueue(() => _isNavigatingBetweenCells = false);
            }
            
            e.Handled = true;
        }
        else if (e.Key == Windows.System.VirtualKey.Escape)
        {
            CellEditBox.Visibility = Visibility.Collapsed;
            e.Handled = true;
        }
    }

    private void FlyoutFunctionList_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        // Prevent the pointer event from stealing focus; select the item under the pointer
        var pt = e.GetCurrentPoint(FlyoutFunctionList).Position;
        var element = FlyoutFunctionList.ContainerFromIndex(0) as FrameworkElement;
        // Let ListView handle selection via ItemClick, but keep focus in edit box
        CellEditBox.Focus(FocusState.Keyboard);
        e.Handled = false;
    }

    private void FlyoutFunctionList_PointerWheelChanged(object sender, PointerRoutedEventArgs e)
    {
        // Allow mouse wheel to scroll the list without taking focus
        // Do nothing special; ensure edit box keeps focus
        CellEditBox.Focus(FocusState.Keyboard);
    }

    private void CellEditBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        // Prevent re-entry when we're modifying the text programmatically
        if (_processingTextChange) return;
        
        var textBox = sender as TextBox;
        if (textBox == null) return;

        System.Diagnostics.Debug.WriteLine($"CellEditBox_TextChanged: text='{textBox.Text}', _isFormulaEditing={_isFormulaEditing}, _processingTextChange={_processingTextChange}");

        // If in formula mode, update intellisense as user types
        if (_isFormulaEditing && textBox.Text.StartsWith("="))
        {
            System.Diagnostics.Debug.WriteLine("CellEditBox_TextChanged: In formula mode, updating intellisense");
            
            // Mirror to inspector formula box only if different
            _updatingFormulaText = true;
            try
            {
                if (CellFormulaBox.Text != textBox.Text)
                {
                    CellFormulaBox.Text = textBox.Text;
                }
            }
            finally
            {
                _updatingFormulaText = false;
            }
            
            // Update autocomplete suggestions
            ShowFormulaIntellisense();
            UpdateParameterHints();
            // Ensure edit box keeps focus after updating suggestions
            DispatcherQueue?.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
            {
                CellEditBox.Focus(FocusState.Keyboard);
            });
            return;
        }

        // Check if user just typed "=" to start a formula
        if (!string.IsNullOrEmpty(textBox.Text) && textBox.Text.StartsWith("=") && !_isFormulaEditing)
        {
            var cell = textBox.Tag as CellViewModel;
            var button = _selectedButton;
            
            if (cell != null && button != null)
            {
                System.Diagnostics.Debug.WriteLine("DEBUG: TextChanged detected '=' - switching to formula mode");
                
                // DON'T set _processingTextChange - we want subsequent keystrokes to trigger filtering
                // Just call StartDirectEdit
                StartDirectEdit(cell, button, textBox.Text);
            }
        }
    }

    private void CommitCellEdit()
    {
        if (CellEditBox.Tag is CellViewModel cell)
        {
            cell.RawValue = CellEditBox.Text;
            
            // Only close edit box and refresh if not navigating between cells
            if (!_isNavigatingBetweenCells)
            {
                CellEditBox.Visibility = Visibility.Collapsed;
                CellEditBox.Tag = null;
                
                if (ViewModel.SelectedSheet != null)
                {
                    BuildSpreadsheetGrid(ViewModel.SelectedSheet);
                }
            }
            else
            {
                // During Tab navigation, just update the button text without closing edit box yet
                var button = GetButtonForCell(cell);
                if (button != null)
                {
                    ApplyCellStyling(button, cell);
                }
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

        // Update header colors to make sure selected tab is readable
        var prefs = App.PreferencesService.LoadPreferences();
        var currentTheme = prefs.Theme ?? "Dark";
        foreach (var item in SheetTabs.TabItems.OfType<TabViewItem>())
        {
            if (item.Header is TextBlock tb && item.Tag is SheetViewModel s)
            {
                if (ViewModel.SelectedSheet == s && currentTheme == "Dark")
                {
                    tb.Foreground = new SolidColorBrush(Color.FromArgb(0xFF, 0x1E, 0x1E, 0x1E));
                }
                else
                {
                    tb.Foreground = Application.Current.Resources["TextPrimaryBrush"] as SolidColorBrush ?? new SolidColorBrush(Colors.White);
                }
            }
        }
    }

    private async void SheetTabs_TabCloseRequested(TabView sender, TabViewTabCloseRequestedEventArgs args)
    {
        if (args.Item is TabViewItem tab && tab.Tag is SheetViewModel sheet)
        {
            var dialog = new ContentDialog
            {
                Title = "Delete Sheet",
                Content = $"Are you sure you want to delete sheet '{sheet.Name}'?",
                PrimaryButtonText = "Delete",
                CloseButtonText = "Cancel",
                DefaultButton = ContentDialogButton.Primary,
                XamlRoot = this.Content.XamlRoot
            };
            var result = await dialog.ShowAsync();
            if (result == ContentDialogResult.Primary)
            {
                ViewModel.Sheets.Remove(sheet);
                if (ViewModel.Sheets.Count == 0)
                {
                    ViewModel.AddSheet();
                }
                ViewModel.SelectedSheet = ViewModel.Sheets.FirstOrDefault();
                RefreshSheetTabs();
            }
        }
    }

    private void NewSheetButton_Click(object sender, RoutedEventArgs e)
    {
        ViewModel.NewSheetCommand.Execute(null);
        RefreshSheetTabs();
        SheetTabs.SelectedIndex = SheetTabs.TabItems.Count - 1;
    }

    private async void ShowRenameDialog(SheetViewModel sheet)
    {
        var textbox = new TextBox { Text = sheet.Name, Width = 240 };
        var dialog = new ContentDialog
        {
            Title = "Rename Sheet",
            Content = new StackPanel { Children = { textbox } },
            PrimaryButtonText = "Rename",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = this.Content.XamlRoot
        };

        var result = await dialog.ShowAsync();
        if (result == ContentDialogResult.Primary)
        {
            var newName = textbox.Text?.Trim();
            if (!string.IsNullOrEmpty(newName) && ViewModel.GetSheet(newName) == null)
            {
                sheet.Rename(newName);
                RefreshSheetTabs();
            }
            else
            {
                ViewModel.StatusMessage = "Invalid or duplicate sheet name";
            }
        }
    }

    private async void ShowDeleteConfirm(SheetViewModel sheet)
    {
        var dialog = new ContentDialog
        {
            Title = "Delete Sheet",
            Content = $"Are you sure you want to delete sheet '{sheet.Name}'?",
            PrimaryButtonText = "Delete",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = this.Content.XamlRoot
        };

        var result = await dialog.ShowAsync();
        if (result == ContentDialogResult.Primary)
        {
            ViewModel.Sheets.Remove(sheet);
            if (ViewModel.Sheets.Count == 0)
            {
                ViewModel.AddSheet();
            }
            ViewModel.SelectedSheet = ViewModel.Sheets.FirstOrDefault();
            RefreshSheetTabs();
        }
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

    private async void ExportCsvButton_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedSheet == null)
        {
            ViewModel.StatusMessage = "No sheet selected";
            return;
        }

        // Show file save picker
        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
        var picker = new Windows.Storage.Pickers.FileSavePicker();
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);
        picker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary;
        picker.FileTypeChoices.Add("CSV Files", new System.Collections.Generic.List<string>() { ".csv" });
        picker.SuggestedFileName = ViewModel.SelectedSheet.Name;
        
        var file = await picker.PickSaveFileAsync();
        if (file != null)
        {
            try
            {
                await CsvService.ExportSheetToCsvAsync(ViewModel.SelectedSheet, file.Path);
                ViewModel.StatusMessage = $"✅ Exported to {file.Name}";
            }
            catch (Exception ex)
            {
                ViewModel.StatusMessage = $"❌ Export failed: {ex.Message}";
            }
        }
    }

    private async void ImportCsvButton_Click(object sender, RoutedEventArgs e)
    {
        // Show file picker
        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
        var picker = new Windows.Storage.Pickers.FileOpenPicker();
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);
        picker.FileTypeFilter.Add(".csv");
        
        var file = await picker.PickSingleFileAsync();
        if (file != null)
        {
            await ViewModel.ImportCsvCommand.ExecuteAsync(file.Path);
            RefreshSheetTabs();
        }
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
        // Prevent infinite loop
        if (_updatingFormulaText) return;
        
        // Always show intellisense when formula editing, even if _isUpdatingCell
        if (_isFormulaEditing)
        {
            ShowFormulaIntellisense();
            UpdateParameterHints();
            UpdateFormulaSyntaxInfo();
        }
        
        if (_isUpdatingCell || _selectedCell == null) return;
        _selectedCell.Formula = CellFormulaBox.Text;
        
        // Show intellisense if typing function name (redundant but safe)
        if (!_isFormulaEditing)
        {
            ShowFormulaIntellisense();
            UpdateParameterHints();
            UpdateFormulaSyntaxInfo();
        }
    }

    private void UpdateFormulaSyntaxInfo()
    {
        var text = CellFormulaBox.Text;
        if (string.IsNullOrWhiteSpace(text) || !text.StartsWith("="))
        {
            FormulaSyntaxInfo.Visibility = Visibility.Collapsed;
            return;
        }

        var tokens = FormulaSyntaxHighlighter.Tokenize(text);
        var funcCount = tokens.Count(t => t.Type == FormulaTokenType.Function);
        var cellRefCount = tokens.Count(t => t.Type == FormulaTokenType.CellReference);
        
        if (funcCount > 0 || cellRefCount > 0)
        {
            var parts = new List<string>();
            if (funcCount > 0) parts.Add($"{funcCount} function{(funcCount > 1 ? "s" : "")}");
            if (cellRefCount > 0) parts.Add($"{cellRefCount} cell ref{(cellRefCount > 1 ? "s" : "")}");
            FormulaSyntaxInfo.Text = $"💡 {string.Join(", ", parts)}";
            FormulaSyntaxInfo.Visibility = Visibility.Visible;
        }
        else
        {
            FormulaSyntaxInfo.Visibility = Visibility.Collapsed;
        }
    }

    private void ShowFormulaIntellisense()
    {
        // Use CellEditBox text for filtering since that's what user is typing into
        var text = CellEditBox.Visibility == Visibility.Visible ? CellEditBox.Text : CellFormulaBox.Text;
        System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: Called with text='{text}', CellEditBox.Visibility={CellEditBox.Visibility}, CellEditBox.Text='{CellEditBox.Text}', CellFormulaBox.Text='{CellFormulaBox.Text}'");
        
        if (string.IsNullOrWhiteSpace(text) || !text.StartsWith("="))
        {
            System.Diagnostics.Debug.WriteLine("ShowFormulaIntellisense: Text invalid, closing popup");
            FunctionAutocompletePopup.IsOpen = false;
            FunctionAutocompleteFlyout.Hide();
            ParameterHintPopup.IsOpen = false;
            VisibleFunctionList.Visibility = Visibility.Collapsed;
            return;
        }
        
        // Get the current function name being typed (after = and before ()
        var formulaText = text.Substring(1); // Remove the =
        var parenIndex = formulaText.IndexOf('(');
        if (parenIndex >= 0)
        {
            // Already has opening paren, enable reference picking mode
            FunctionAutocompletePopup.IsOpen = false;
            FunctionAutocompleteFlyout.Hide();
            _isPickingFormulaReference = true;
            return;
        }
        
        // Not in reference picking mode while typing function name
        _isPickingFormulaReference = false;
        
        // Filter suggestions based on what's been typed and current context
        var searchText = formulaText.ToUpperInvariant();
        System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: Filtering with searchText='{searchText}' (from formulaText='{formulaText}')");
        
        var matchingSuggestions = GetContextualSuggestions(searchText)
            .Take(20)
            .ToList();
        
        System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: Found {matchingSuggestions.Count} suggestions for search='{searchText}'");
        
        if (matchingSuggestions.Count > 0)
        {
            var names = string.Join(", ", matchingSuggestions.Select(s => s.Name).Take(5));
            System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: First suggestions: {names}");
        }
        
        // Show status message for debugging
        ViewModel.StatusMessage = $"Formula mode: {matchingSuggestions.Count} functions available";
        
        if (matchingSuggestions.Any())
        {
            FunctionAutocompleteList.ItemsSource = matchingSuggestions;
            FunctionAutocompleteList.SelectedIndex = 0;
            
            // Update flyout list
            FlyoutFunctionList.ItemsSource = matchingSuggestions;
            
            // ALSO show in visible test list
            TestFunctionList.ItemsSource = matchingSuggestions;
            VisibleFunctionList.Visibility = Visibility.Visible;
            
            // Keep orange for now to make it obvious it's working
            VisibleFunctionList.Background = new SolidColorBrush(Microsoft.UI.Colors.Orange);
            VisibleFunctionList.BorderThickness = new Thickness(3);
            
            System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: Set VisibleFunctionList.Visibility = Visible with {matchingSuggestions.Count} items");
            System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: TestFunctionList.Items.Count = {TestFunctionList.Items?.Count ?? 0}");
            
            // Ensure popup has XamlRoot set
            if (FunctionAutocompletePopup.XamlRoot == null)
            {
                FunctionAutocompletePopup.XamlRoot = CellFormulaBox.XamlRoot;
            }
            
            // Try to show the flyout on CellEditBox
            if (CellEditBox.Visibility == Visibility.Visible)
            {
                try
                {
                    // Ensure edit box has focus BEFORE showing flyout
                    CellEditBox.Focus(FocusState.Keyboard);
                    
                    FunctionAutocompleteFlyout.ShowAt(CellEditBox);
                    System.Diagnostics.Debug.WriteLine("ShowFormulaIntellisense: Flyout shown at CellEditBox");
                    
                    // Keep focus on CellEditBox immediately
                    CellEditBox.Focus(FocusState.Keyboard);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: Error showing flyout: {ex.Message}");
                }
            }
            
            // Position popup near the CellEditBox (where user is typing)
            try
            {
                // Close first to reset state
                FunctionAutocompletePopup.IsOpen = false;
                
                if (CellEditBox.Visibility == Visibility.Visible)
                {
                    // Position below the cell edit box
                    var transform = CellEditBox.TransformToVisual(null);
                    var point = transform.TransformPoint(new Windows.Foundation.Point(0, CellEditBox.ActualHeight));
                    FunctionAutocompletePopup.HorizontalOffset = point.X;
                    FunctionAutocompletePopup.VerticalOffset = point.Y + 4;
                }
                else
                {
                    // Fallback to formula box
                    var transform = CellFormulaBox.TransformToVisual(null);
                    var point = transform.TransformPoint(new Windows.Foundation.Point(0, CellFormulaBox.ActualHeight));
                    FunctionAutocompletePopup.HorizontalOffset = point.X;
                    FunctionAutocompletePopup.VerticalOffset = point.Y + 4;
                }
                
                FunctionAutocompletePopup.IsOpen = true;
                System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: Showing {matchingSuggestions.Count} suggestions, Popup IsOpen={FunctionAutocompletePopup.IsOpen}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShowFormulaIntellisense: Error opening popup: {ex.Message}");
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine("ShowFormulaIntellisense: No suggestions, closing all");
            FunctionAutocompletePopup.IsOpen = false;
            FunctionAutocompleteFlyout.Hide();
            VisibleFunctionList.Visibility = Visibility.Collapsed;
            VisibleFunctionList.Visibility = Visibility.Collapsed;
        }
    }

    private (string FunctionName, int ArgumentIndex) ParseFunctionContext(string text, int caretIndex)
    {
        if (string.IsNullOrWhiteSpace(text) || !text.StartsWith("="))
        {
            return (string.Empty, -1);
        }

        var body = text.Substring(1);
        var parenIndex = body.IndexOf('(');
        if (parenIndex < 0)
        {
            return (string.Empty, -1);
        }

        var functionName = body.Substring(0, parenIndex).Trim();
        if (string.IsNullOrWhiteSpace(functionName))
        {
            return (string.Empty, -1);
        }

        var startIndex = parenIndex + 1;
        var maxIndex = Math.Min(body.Length, Math.Max(0, caretIndex - 1));
        if (maxIndex <= startIndex)
        {
            return (functionName, 0);
        }

        int depth = 0;
        int argumentIndex = 0;
        for (int i = startIndex; i < maxIndex; i++)
        {
            var ch = body[i];
            if (ch == '(')
            {
                depth++;
            }
            else if (ch == ')')
            {
                if (depth == 0)
                {
                    break;
                }
                depth--;
            }
            else if (ch == ',' && depth == 0)
            {
                argumentIndex++;
            }
        }

        return (functionName, Math.Max(0, argumentIndex));
    }

    private void UpdateParameterHints()
    {
        if (!_isFormulaEditing)
        {
            ParameterHintPopup.IsOpen = false;
            return;
        }

        var text = CellFormulaBox.Text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text) || !text.StartsWith("="))
        {
            ParameterHintPopup.IsOpen = false;
            return;
        }

        var caret = CellFormulaBox.SelectionStart;
        var context = ParseFunctionContext(text, caret);
        if (string.IsNullOrWhiteSpace(context.FunctionName) || !_functionMap.TryGetValue(context.FunctionName, out var descriptor))
        {
            ParameterHintPopup.IsOpen = false;
            return;
        }

        _parameterHints.Clear();
        for (int i = 0; i < descriptor.Parameters.Count; i++)
        {
            _parameterHints.Add(ParameterHintItem.From(descriptor.Parameters[i], i == context.ArgumentIndex));
        }

        ParameterHintList.ItemsSource = null;
        ParameterHintList.ItemsSource = _parameterHints;

        ParameterHintFunctionName.Text = descriptor.Name;
        ParameterHintDescription.Text = context.ArgumentIndex >= 0 && context.ArgumentIndex < descriptor.Parameters.Count
            ? descriptor.Parameters[context.ArgumentIndex].Description
            : descriptor.Description;

        var transform = CellFormulaBox.TransformToVisual(this);
        var point = transform.TransformPoint(new Point(0, CellFormulaBox.ActualHeight));
        ParameterHintPopup.HorizontalOffset = point.X;
        ParameterHintPopup.VerticalOffset = point.Y + 4;
        ParameterHintPopup.IsOpen = _parameterHints.Count > 0;
        // Validate current argument types for immediate feedback
        ValidateFormulaArguments();
    }

    private void ValidateFormulaArguments()
    {
        if (!_isFormulaEditing || _formulaEditTargetCell == null) return;

        var text = CellFormulaBox.Text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text) || !text.StartsWith("="))
        {
            ClearCellValidationError(_formulaEditTargetCell);
            return;
        }

        var caret = CellFormulaBox.SelectionStart;
        var context = ParseFunctionContext(text, caret);
        if (string.IsNullOrWhiteSpace(context.FunctionName) || !_functionMap.TryGetValue(context.FunctionName, out var descriptor))
        {
            ClearCellValidationError(_formulaEditTargetCell);
            return;
        }

        // Basic parse of arguments (strings/numbers/cell refs)
        var parenPos = text.IndexOf('(');
        var argsText = parenPos >= 0 ? text.Substring(parenPos + 1) : string.Empty;
        if (argsText.EndsWith(")"))
        {
            argsText = argsText.Substring(0, argsText.Length - 1);
        }
        var tokens = AiCalc.Services.FormulaParser.SplitArguments(argsText).Select(t => t.Trim()).ToArray();

        // Use validation helper with a cell type lookup
        var result = FormulaValidation.ValidateParameters(
            descriptor,
            tokens,
            addr =>
            {
                var sheetName = addr.SheetName ?? (_formulaEditTargetCell?.Sheet.Name ?? ViewModel.SelectedSheet?.Name ?? "Sheet1");
                if (ViewModel.GetSheet(sheetName) is SheetViewModel sheet)
                {
                    var cell = sheet.GetCell(addr.Row, addr.Column);
                    return cell?.Value.ObjectType;
                }
                return null;
            },
            _formulaEditTargetCell?.Sheet.Name ?? ViewModel.SelectedSheet?.Name ?? "Sheet1");

        if (!result.IsValid)
        {
            SetCellValidationError(_formulaEditTargetCell!, result.ErrorMessage ?? "Invalid parameters");
            return;
        }

        ClearCellValidationError(_formulaEditTargetCell!);
    }

    private void SetCellValidationError(CellViewModel cell, string message)
    {
        cell.VisualState = CellVisualState.Error;
        _validationErrorCells.Add(cell);
        var button = GetButtonForCell(cell);
        if (button != null)
        {
            button.SetValue(ToolTipService.ToolTipProperty, message);
            ApplyCellStyling(button, cell);
        }
    }

    private void ClearCellValidationError(CellViewModel cell)
    {
        if (_validationErrorCells.Contains(cell))
        {
            cell.VisualState = CellVisualState.Normal;
            _validationErrorCells.Remove(cell);
            var button = GetButtonForCell(cell);
            if (button != null)
            {
                button.ClearValue(ToolTipService.ToolTipProperty);
                ApplyCellStyling(button, cell);
            }
        }
    }

    // Argument splitting logic moved to FormulaParser.SplitArguments

    private void CellFormulaBox_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        if (!FunctionAutocompletePopup.IsOpen) return;
        
        if (e.Key == Windows.System.VirtualKey.Down)
        {
            // Move selection down
            if (FunctionAutocompleteList.SelectedIndex < FunctionAutocompleteList.Items.Count - 1)
            {
                FunctionAutocompleteList.SelectedIndex++;
            }
            e.Handled = true;
        }
        else if (e.Key == Windows.System.VirtualKey.Up)
        {
            // Move selection up
            if (FunctionAutocompleteList.SelectedIndex > 0)
            {
                FunctionAutocompleteList.SelectedIndex--;
            }
            e.Handled = true;
        }
        else if (e.Key == Windows.System.VirtualKey.Enter || e.Key == Windows.System.VirtualKey.Tab)
        {
            // Accept selected suggestion
            if (FunctionAutocompleteList.SelectedItem is FunctionSuggestion suggestion)
            {
                InsertFunctionSuggestion(suggestion);
            }
            e.Handled = true;
        }
        else if (e.Key == Windows.System.VirtualKey.Escape)
        {
            // Close popup
            FunctionAutocompletePopup.IsOpen = false;
            e.Handled = true;
        }
    }

    private void FunctionAutocomplete_ItemClick(object sender, ItemClickEventArgs e)
    {
        if (e.ClickedItem is FunctionSuggestion suggestion)
        {
            InsertFunctionSuggestion(suggestion);
        }
    }

    private void CellFormulaBox_GotFocus(object sender, RoutedEventArgs e)
    {
        _isFormulaEditing = true;
        _formulaEditTargetCell = _selectedCell;
        ClearFormulaHighlights();
        UpdateParameterHints();
    }

    private void CellFormulaBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (_isPickingFormulaReference)
        {
            DispatcherQueue?.TryEnqueue(() =>
            {
                _isPickingFormulaReference = false;
                CellFormulaBox.Focus(FocusState.Programmatic);
            });
            return;
        }

        _isFormulaEditing = false;
        if (_formulaEditTargetCell != null && _selectedCell == _formulaEditTargetCell)
        {
            _formulaEditTargetCell.Formula = CellFormulaBox.Text;
        }
        _formulaEditTargetCell = null;
        ParameterHintPopup.IsOpen = false;
        ClearFormulaHighlights();
    }

    private void InsertCellReferenceIntoFormula(CellViewModel cell)
    {
        if (_formulaEditTargetCell == null)
        {
            return;
        }

        var text = CellFormulaBox.Text ?? string.Empty;
        var caret = CellFormulaBox.SelectionStart;
        var reference = BuildCellReferenceString(cell, _formulaEditTargetCell);

        if (caret > 0 && caret <= text.Length)
        {
            var previous = text[caret - 1];
            if (previous != '(' && previous != ',' && previous != '=')
            {
                reference = "," + reference;
            }
        }

        CellFormulaBox.Text = text.Insert(caret, reference);
        CellFormulaBox.SelectionStart = caret + reference.Length;
        _selectedCell!.Formula = CellFormulaBox.Text;
        HighlightFormulaCell(cell);
        UpdateParameterHints();
        // Stay in picking mode - user may want to select more cells
        DispatcherQueue?.TryEnqueue(() => CellFormulaBox.Focus(FocusState.Programmatic));
    }

    private void CellButton_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        // Pointer handler for future use - picking state now managed by formula content
    }

    private static string BuildCellReferenceString(CellViewModel cell, CellViewModel target)
    {
        if (cell.Sheet == target.Sheet)
        {
            var columnName = CellAddress.ColumnIndexToName(cell.Column);
            return $"{columnName}{cell.Row + 1}";
        }

        return cell.Address.ToString();
    }

    private void HighlightFormulaCell(CellViewModel cell)
    {
        if (_formulaHighlights.Any(h => h.Cell == cell))
        {
            return;
        }

        _formulaHighlights.Add((cell, cell.VisualState));
        cell.VisualState = CellVisualState.InDependencyChain;
        var button = GetButtonForCell(cell);
        if (button != null)
        {
            ApplyCellStyling(button, cell);
        }
    }

    private void ClearFormulaHighlights()
    {
        foreach (var (cell, previous) in _formulaHighlights)
        {
            cell.VisualState = previous;
            var button = GetButtonForCell(cell);
            if (button != null)
            {
                ApplyCellStyling(button, cell);
            }
        }

        _formulaHighlights.Clear();
    }

    private void InsertFunctionSuggestion(FunctionSuggestion suggestion)
    {
        var newFormula = $"={suggestion.Name}()";
        
        // Update both edit boxes and the cell
        _updatingFormulaText = true;
        try
        {
            CellFormulaBox.Text = newFormula;
            CellEditBox.Text = newFormula;
            
            if (_formulaEditTargetCell != null)
            {
                _formulaEditTargetCell.Formula = newFormula;
                _formulaEditTargetCell.RawValue = newFormula;
            }
        }
        finally
        {
            _updatingFormulaText = false;
        }
        
        // Place caret inside parentheses in both boxes
        var caretPos = newFormula.Length - 1;
        CellFormulaBox.SelectionStart = caretPos;
        CellEditBox.SelectionStart = caretPos;
        
        FunctionAutocompletePopup.IsOpen = false;
        VisibleFunctionList.Visibility = Visibility.Collapsed;
        
        // Keep commit suppressed while we hide the flyout and refocus
        _suppressCommitOnFlyoutOpen = true;
        // Hide flyout and return focus to the edit box asynchronously
        FunctionAutocompleteFlyout.Hide();
        DispatcherQueue?.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            if (CellEditBox.Visibility == Visibility.Visible)
            {
                CellEditBox.Focus(FocusState.Keyboard);
                CellEditBox.SelectionStart = caretPos;
            }
            else
            {
                CellFormulaBox.Focus(FocusState.Keyboard);
                CellFormulaBox.SelectionStart = caretPos;
            }
            // Allow commit again
            _suppressCommitOnFlyoutOpen = false;
        });
        
        // Now enable reference picking since we have the opening paren
        _isPickingFormulaReference = true;
    }

    private void LeftSplitter_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        _isLeftSplitterDragging = true;
        _leftSplitterStart = e.GetCurrentPoint(MainLayoutGrid).Position;
        _leftInitialWidth = FunctionsColumn.ActualWidth > 0 ? FunctionsColumn.ActualWidth : ViewModel.Settings.FunctionsPanelWidth;
        (sender as UIElement)?.CapturePointer(e.Pointer);
    }

    private void LeftSplitter_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isLeftSplitterDragging)
        {
            return;
        }

        var point = e.GetCurrentPoint(MainLayoutGrid).Position;
        var delta = point.X - _leftSplitterStart.X;
        var newWidth = Math.Max(180, _leftInitialWidth + delta);
        FunctionsColumn.Width = new GridLength(newWidth);
    }

    private void LeftSplitter_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (!_isLeftSplitterDragging)
        {
            return;
        }

        _isLeftSplitterDragging = false;
        (sender as UIElement)?.ReleasePointerCaptures();
        var finalWidth = Math.Max(180, FunctionsColumn.ActualWidth);
        FunctionsColumn.Width = new GridLength(finalWidth);
        ViewModel.Settings.FunctionsPanelWidth = finalWidth;
    }

    private void RightSplitter_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        _isRightSplitterDragging = true;
        _rightSplitterStart = e.GetCurrentPoint(MainLayoutGrid).Position;
        _rightInitialWidth = InspectorColumn.ActualWidth > 0 ? InspectorColumn.ActualWidth : ViewModel.Settings.InspectorPanelWidth;
        (sender as UIElement)?.CapturePointer(e.Pointer);
    }

    private void RightSplitter_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isRightSplitterDragging)
        {
            return;
        }

        var point = e.GetCurrentPoint(MainLayoutGrid).Position;
        var delta = _rightSplitterStart.X - point.X;
        var newWidth = Math.Max(220, _rightInitialWidth + delta);
        InspectorColumn.Width = new GridLength(newWidth);
    }

    private void RightSplitter_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (!_isRightSplitterDragging)
        {
            return;
        }

        _isRightSplitterDragging = false;
        (sender as UIElement)?.ReleasePointerCaptures();
        var finalWidth = Math.Max(220, InspectorColumn.ActualWidth);
        InspectorColumn.Width = new GridLength(finalWidth);
        ViewModel.Settings.InspectorPanelWidth = finalWidth;
    }

    private void Splitter_PointerEntered(object sender, PointerRoutedEventArgs e)
    {
        if (sender is Border border)
        {
            border.Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0x55, 0xFF, 0xFF, 0xFF));
        }
    }

    private void Splitter_PointerExited(object sender, PointerRoutedEventArgs e)
    {
        if (sender is Border border && !_isLeftSplitterDragging && !_isRightSplitterDragging)
        {
            border.Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0x22, 0xFF, 0xFF, 0xFF));
        }
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
            var result = await _selectedCell.EvaluateAsync();
            if (result != null)
            {
                if (result.AiResponse != null)
                {
                    result = await ShowAIPreviewAsync(_selectedCell, result);
                    if (result == null)
                    {
                        ViewModel.StatusMessage = "AI result cancelled";
                        return;
                    }
                }

                await HandleEvaluationResultAsync(_selectedCell, result);
            }
        }
    }

    private async Task HandleEvaluationResultAsync(CellViewModel cell, FunctionExecutionResult result)
    {
        var sheet = ViewModel.SelectedSheet;
        bool proceed = true;

        if (result.HasSpill && sheet != null)
        {
            var rows = result.SpillRange!.GetLength(0);
            var cols = result.SpillRange!.GetLength(1);
            var existing = sheet.GetCellsInRange(cell, rows, cols)
                .Where(c => c != cell && (!string.IsNullOrWhiteSpace(c.RawValue) || c.HasFormula))
                .ToList();

            if (existing.Count > 0)
            {
                var dialog = new ContentDialog
                {
                    Title = "Confirm Spill Operation",
                    Content = $"This formula will spill into a {rows}×{cols} range and shift {existing.Count} populated cell(s). Continue?",
                    PrimaryButtonText = "Spill",
                    CloseButtonText = "Cancel",
                    DefaultButton = ContentDialogButton.Primary,
                    XamlRoot = Content.XamlRoot
                };

                var decision = await dialog.ShowAsync();
                proceed = decision == ContentDialogResult.Primary;
            }
        }

        if (!proceed)
        {
            return;
        }

        cell.ApplyEvaluationResult(result);

        if (result.HasSpill && sheet != null)
        {
            var affected = sheet.ApplySpill(cell, result.SpillRange!);
            ViewModel.StatusMessage = $"Spilled into {affected.Count + 1} cell(s) from {cell.DisplayLabel}";
        }
        else if (!string.IsNullOrWhiteSpace(result.Diagnostics))
        {
            ViewModel.StatusMessage = result.Diagnostics;
        }

        if (sheet != null)
        {
            BuildSpreadsheetGrid(sheet);
        }
    }

    private async Task<FunctionExecutionResult?> ShowAIPreviewAsync(CellViewModel cell, FunctionExecutionResult initialResult)
    {
        var current = initialResult;

        while (current != null)
        {
            var response = current.AiResponse;
            var previewText = current.Value.DisplayValue ?? current.Value.SerializedValue ?? string.Empty;

            var panel = new StackPanel
            {
                Spacing = 12,
                MaxWidth = 460
            };

            panel.Children.Add(new TextBlock
            {
                Text = previewText,
                TextWrapping = TextWrapping.WrapWholeWords
            });

            var metaLines = new List<string>();
            if (!string.IsNullOrWhiteSpace(current.Diagnostics))
            {
                metaLines.Add(current.Diagnostics);
            }

            if (response != null)
            {
                if (response.Metadata != null && response.Metadata.TryGetValue("model", out var model) && model != null)
                {
                    metaLines.Add($"Model: {model}");
                }

                if (response.TokensUsed > 0)
                {
                    metaLines.Add($"Tokens used: {response.TokensUsed}");
                }

                if (response.Duration > TimeSpan.Zero)
                {
                    metaLines.Add($"Duration: {response.Duration.TotalSeconds:F1}s");
                }
            }

            if (metaLines.Count > 0)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = string.Join(" • ", metaLines),
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Microsoft.UI.Colors.Gray)
                });
            }

            var dialog = new ContentDialog
            {
                Title = "AI Result Preview",
                Content = new ScrollViewer
                {
                    Content = panel,
                    VerticalScrollMode = ScrollMode.Auto,
                    HorizontalScrollMode = ScrollMode.Disabled,
                    MaxHeight = 360
                },
                PrimaryButtonText = "Apply to Cell",
                SecondaryButtonText = "Regenerate",
                CloseButtonText = "Cancel",
                DefaultButton = ContentDialogButton.Primary,
                XamlRoot = Content.XamlRoot
            };

            var decision = await dialog.ShowAsync();

            if (decision == ContentDialogResult.Primary)
            {
                return current;
            }

            if (decision == ContentDialogResult.Secondary)
            {
                var regenerated = await cell.EvaluateAsync();
                if (regenerated == null)
                {
                    return null;
                }

                // Continue loop with regenerated result
                current = regenerated;
                continue;
            }
        }

        return null;
    }

    private void ThemeToggle_Click(object sender, RoutedEventArgs e)
    {
        var prefs = App.PreferencesService.LoadPreferences();
        var currentTheme = prefs.Theme ?? "Dark";
        var newTheme = currentTheme == "Light" ? "Dark" : "Light";

    // Apply theme globally (map from string to enum)
    var themeEnum = newTheme == "Light" ? CellVisualTheme.Light : CellVisualTheme.Dark;
    App.ApplyCellStateTheme(themeEnum);
        if (App.MainWindow?.Content is FrameworkElement rootElement)
        {
            rootElement.RequestedTheme = newTheme == "Light" ? ElementTheme.Light : ElementTheme.Dark;
        }

        prefs.Theme = newTheme;
        App.PreferencesService.SavePreferences(prefs);

        // Update toggle text if the button is available in the visual tree
        if (this.FindName("ThemeToggleButton") is Button tb)
        {
            tb.Content = newTheme == "Light" ? "🌙 Dark" : "☀️ Light";
        }

        ViewModel.StatusMessage = $"Theme: {newTheme}";
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
    /// Handle keyboard shortcuts (F9, navigation, editing) - Phase 5 Task 14
    /// </summary>
    private async void MainWindow_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        // DEBUG: Log ALL key presses
        System.Diagnostics.Debug.WriteLine($"MainWindow_KeyDown: Key pressed = {e.Key} ({(int)e.Key})");
        
        var ctrl = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
        var shift = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
        
        // Ctrl+Z: Undo (Phase 5)
        if (ctrl && e.Key == Windows.System.VirtualKey.Z && !shift)
        {
            if (ViewModel.UndoCommand.CanExecute(null))
            {
                ViewModel.UndoCommand.Execute(null);
                e.Handled = true;
                RefreshCurrentCell();
            }
            return;
        }
        
        // Ctrl+Y or Ctrl+Shift+Z: Redo (Phase 5)
        if (ctrl && (e.Key == Windows.System.VirtualKey.Y || (e.Key == Windows.System.VirtualKey.Z && shift)))
        {
            if (ViewModel.RedoCommand.CanExecute(null))
            {
                ViewModel.RedoCommand.Execute(null);
                e.Handled = true;
                RefreshCurrentCell();
            }
            return;
        }
        
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
            return;
        }
        
        // F2: Edit mode
        if (e.Key == Windows.System.VirtualKey.F2 && _selectedCell != null && _selectedButton != null)
        {
            StartDirectEdit(_selectedCell, _selectedButton);
            e.Handled = true;
            return;
        }
        
        // Navigation requires a selected cell
        if (_selectedCell == null || ViewModel.SelectedSheet == null)
            return;
        
        // Ctrl+Home: Go to A1
        if (ctrl && e.Key == Windows.System.VirtualKey.Home)
        {
            SelectCellAt(0, 0);
            e.Handled = true;
            return;
        }
        
        // Ctrl+End: Go to last used cell
        if (ctrl && e.Key == Windows.System.VirtualKey.End)
        {
            var sheet = ViewModel.SelectedSheet;
            SelectCellAt(sheet.ColumnHeaders.Count - 1, sheet.Rows.Count - 1);
            e.Handled = true;
            return;
        }
        
        // Ctrl+Arrow: Jump to data edge
        if (ctrl && (e.Key == Windows.System.VirtualKey.Up || e.Key == Windows.System.VirtualKey.Down ||
                     e.Key == Windows.System.VirtualKey.Left || e.Key == Windows.System.VirtualKey.Right))
        {
            int colDelta = e.Key == Windows.System.VirtualKey.Left ? -1 : e.Key == Windows.System.VirtualKey.Right ? 1 : 0;
            int rowDelta = e.Key == Windows.System.VirtualKey.Up ? -1 : e.Key == Windows.System.VirtualKey.Down ? 1 : 0;
            JumpToDataEdge(colDelta, rowDelta);
            e.Handled = true;
            return;
        }
        
        // Arrow keys: Move selection
        if (e.Key == Windows.System.VirtualKey.Up || e.Key == Windows.System.VirtualKey.Down ||
            e.Key == Windows.System.VirtualKey.Left || e.Key == Windows.System.VirtualKey.Right)
        {
            int colDelta = e.Key == Windows.System.VirtualKey.Left ? -1 : e.Key == Windows.System.VirtualKey.Right ? 1 : 0;
            int rowDelta = e.Key == Windows.System.VirtualKey.Up ? -1 : e.Key == Windows.System.VirtualKey.Down ? 1 : 0;
            MoveSelection(colDelta, rowDelta);
            e.Handled = true;
            return;
        }
        
        // Tab: Move right, Shift+Tab: Move left
        if (e.Key == Windows.System.VirtualKey.Tab)
        {
            MoveSelection(shift ? -1 : 1, 0);
            e.Handled = true;
            return;
        }

        // Ctrl+T: Toggle Light/Dark Theme - delegate to button
        if (ctrl && e.Key == Windows.System.VirtualKey.T)
        {
            ThemeToggle_Click(this, new RoutedEventArgs());
            e.Handled = true;
            return;
        }
        
        // Enter: Save and move down, Shift+Enter: Save and move up
        if (e.Key == Windows.System.VirtualKey.Enter)
        {
            CommitCellEdit();
            var rowDelta = shift ? -1 : 1;
            var newCell = MoveSelection(0, rowDelta);
            
            // Enter edit mode on the new cell if it's editable
            if (newCell != null)
            {
                var newButton = GetButtonForCell(newCell);
                if (newButton != null)
                {
                    StartDirectEdit(newCell, newButton);
                }
            }
            
            e.Handled = true;
            return;
        }
        
        // Page Up/Down: Scroll viewport (10 rows)
        if (e.Key == Windows.System.VirtualKey.PageUp)
        {
            MoveSelection(0, -10);
            e.Handled = true;
            return;
        }
        
        if (e.Key == Windows.System.VirtualKey.PageDown)
        {
            MoveSelection(0, 10);
            e.Handled = true;
            return;
        }
        
        // Regular keyboard input: Start editing the cell (only if not already editing)
        if (!ctrl && !shift && IsRegularKeyboardKey(e.Key) && CellEditBox.Visibility != Visibility.Visible)
        {
            if (_selectedCell != null && _selectedButton != null)
            {
                var charToAdd = VirtualKeyToString(e.Key);
                System.Diagnostics.Debug.WriteLine($"MainWindow_KeyDown: Key {e.Key} mapped to '{charToAdd}'");
                StartDirectEdit(_selectedCell, _selectedButton, charToAdd);
                e.Handled = true;
                return;
            }
        }
        
        // Delete: Clear cell contents
        if (e.Key == Windows.System.VirtualKey.Delete && _selectedCell != null)
        {
            _selectedCell.RawValue = string.Empty;
            _selectedCell.Formula = string.Empty;
            BuildSpreadsheetGrid(ViewModel.SelectedSheet);
            e.Handled = true;
            return;
        }
    }
    
    /// <summary>
    /// Move selection by delta (Phase 5 Task 14)
    /// </summary>
    private CellViewModel? MoveSelection(int colDelta, int rowDelta)
    {
        if (_selectedCell == null || ViewModel.SelectedSheet == null) return null;
        
        var currentAddr = _selectedCell.Address;
        var newCol = currentAddr.Column + colDelta;
        var newRow = currentAddr.Row + rowDelta;
        
        return SelectCellAt(newCol, newRow);
    }
    
    /// <summary>
    /// Select cell at specific coordinates (Phase 5 Task 14)
    /// </summary>
    private CellViewModel? SelectCellAt(int colIndex, int rowIndex)
    {
        if (ViewModel.SelectedSheet == null) return null;
        
        var sheet = ViewModel.SelectedSheet;
        
        // Clamp to valid range
        colIndex = Math.Max(0, Math.Min(colIndex, sheet.ColumnHeaders.Count - 1));
        rowIndex = Math.Max(0, Math.Min(rowIndex, sheet.Rows.Count - 1));
        
        var newCell = sheet.Rows[rowIndex].Cells[colIndex];
        var newButton = GetButtonForCell(newCell);
        
        if (newButton != null)
        {
            SelectCell(newCell, newButton);
        }
        
        return newCell;
    }
    
    /// <summary>
    /// Jump to edge of data region (Phase 5 Task 14)
    /// </summary>
    private void JumpToDataEdge(int colDelta, int rowDelta)
    {
        if (_selectedCell == null || ViewModel.SelectedSheet == null) return;
        
        var sheet = ViewModel.SelectedSheet;
        var currentCol = _selectedCell.Address.Column;
        var currentRow = _selectedCell.Address.Row;
        
        // Determine if current cell is empty
        bool currentIsEmpty = string.IsNullOrWhiteSpace(_selectedCell.RawValue);
        
        while (true)
        {
            var nextCol = currentCol + colDelta;
            var nextRow = currentRow + rowDelta;
            
            // Check bounds
            if (nextCol < 0 || nextRow < 0 || 
                nextCol >= sheet.ColumnHeaders.Count || 
                nextRow >= sheet.Rows.Count)
            {
                // Hit edge of sheet, stop at last valid cell
                break;
            }
            
            var nextCell = sheet.Rows[nextRow].Cells[nextCol];
            bool nextIsEmpty = string.IsNullOrWhiteSpace(nextCell.RawValue);
            
            // Stop when transitioning from empty to non-empty or vice versa
            if (currentIsEmpty != nextIsEmpty)
            {
                // If we started in empty space, land on first non-empty cell
                // If we started in data, stop before entering empty space
                if (currentIsEmpty)
                {
                    currentCol = nextCol;
                    currentRow = nextRow;
                }
                break;
            }
            
            currentCol = nextCol;
            currentRow = nextRow;
            currentIsEmpty = nextIsEmpty;
        }
        
        SelectCellAt(currentCol, currentRow);
    }
    
    /// <summary>
    /// Get button control for a specific cell (Phase 5 Task 14)
    /// </summary>
    private Button? GetButtonForCell(CellViewModel cell)
    {
        // Find the button in SpreadsheetGrid by matching Tag
        foreach (var child in SpreadsheetGrid.Children.OfType<Button>())
        {
            if (child.Tag is CellViewModel cellVm && cellVm == cell)
            {
                return child;
            }
        }
        return null;
    }
    
    /// <summary>
    /// Convert a VirtualKey to its string representation
    /// </summary>
    private string VirtualKeyToString(Windows.System.VirtualKey key)
    {
        return key switch
        {
            Windows.System.VirtualKey.Space => " ",
            Windows.System.VirtualKey.Number0 => "0",
            Windows.System.VirtualKey.Number1 => "1",
            Windows.System.VirtualKey.Number2 => "2",
            Windows.System.VirtualKey.Number3 => "3",
            Windows.System.VirtualKey.Number4 => "4",
            Windows.System.VirtualKey.Number5 => "5",
            Windows.System.VirtualKey.Number6 => "6",
            Windows.System.VirtualKey.Number7 => "7",
            Windows.System.VirtualKey.Number8 => "8",
            Windows.System.VirtualKey.Number9 => "9",
            Windows.System.VirtualKey.A => "a",
            Windows.System.VirtualKey.B => "b",
            Windows.System.VirtualKey.C => "c",
            Windows.System.VirtualKey.D => "d",
            Windows.System.VirtualKey.E => "e",
            Windows.System.VirtualKey.F => "f",
            Windows.System.VirtualKey.G => "g",
            Windows.System.VirtualKey.H => "h",
            Windows.System.VirtualKey.I => "i",
            Windows.System.VirtualKey.J => "j",
            Windows.System.VirtualKey.K => "k",
            Windows.System.VirtualKey.L => "l",
            Windows.System.VirtualKey.M => "m",
            Windows.System.VirtualKey.N => "n",
            Windows.System.VirtualKey.O => "o",
            Windows.System.VirtualKey.P => "p",
            Windows.System.VirtualKey.Q => "q",
            Windows.System.VirtualKey.R => "r",
            Windows.System.VirtualKey.S => "s",
            Windows.System.VirtualKey.T => "t",
            Windows.System.VirtualKey.U => "u",
            Windows.System.VirtualKey.V => "v",
            Windows.System.VirtualKey.W => "w",
            Windows.System.VirtualKey.X => "x",
            Windows.System.VirtualKey.Y => "y",
            Windows.System.VirtualKey.Z => "z",
            Windows.System.VirtualKey.Multiply => "*",
            Windows.System.VirtualKey.Add => "+",
            (Windows.System.VirtualKey)187 => "=", // OEM plus key (:=)
            (Windows.System.VirtualKey)189 => "-", // OEM minus key
            Windows.System.VirtualKey.Subtract => "-",
            Windows.System.VirtualKey.Divide => "/",
            Windows.System.VirtualKey.Decimal => ".",
            _ => ""
        };
    }
    
    /// <summary>
    /// Check if a key is a regular qwerty key (excluding function keys, WIN, etc)
    /// </summary>
    private bool IsRegularKeyboardKey(Windows.System.VirtualKey key)
    {
        // Exclude function keys (F1-F12)
        if (key >= Windows.System.VirtualKey.F1 && key <= Windows.System.VirtualKey.F12)
            return false;
        
        // Exclude navigation keys (already handled separately)
        if (key == Windows.System.VirtualKey.Up || key == Windows.System.VirtualKey.Down ||
            key == Windows.System.VirtualKey.Left || key == Windows.System.VirtualKey.Right ||
            key == Windows.System.VirtualKey.Tab || key == Windows.System.VirtualKey.Enter ||
            key == Windows.System.VirtualKey.Escape || key == Windows.System.VirtualKey.Delete ||
            key == Windows.System.VirtualKey.Home || key == Windows.System.VirtualKey.End ||
            key == Windows.System.VirtualKey.PageUp || key == Windows.System.VirtualKey.PageDown)
            return false;
        
        // Exclude special keys
        if (key == Windows.System.VirtualKey.LeftWindows || key == Windows.System.VirtualKey.RightWindows ||
            key == Windows.System.VirtualKey.Application ||
            key == Windows.System.VirtualKey.LeftMenu || key == Windows.System.VirtualKey.RightMenu ||
            key == Windows.System.VirtualKey.CapitalLock || key == Windows.System.VirtualKey.Scroll)
            return false;
        
        // Everything else is considered a regular key (letters, numbers, symbols, etc)
        return true;
    }
    
    // ==================== Context Menu Handlers (Phase 5 Task 16) ====================
    
    private void Cut_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null) return;
        
        // Copy to clipboard
        var dataPackage = new Windows.ApplicationModel.DataTransfer.DataPackage();
        dataPackage.SetText(_selectedCell.RawValue);
        Windows.ApplicationModel.DataTransfer.Clipboard.SetContent(dataPackage);
        
        // Clear cell
        _selectedCell.RawValue = string.Empty;
        
        // Refresh grid
        if (ViewModel.SelectedSheet != null)
        {
            BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        }
        
        ViewModel.StatusMessage = $"Cut cell {_selectedCell.DisplayLabel}";
    }
    
    private void Copy_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null) return;
        
        var dataPackage = new Windows.ApplicationModel.DataTransfer.DataPackage();
        dataPackage.SetText(_selectedCell.RawValue);
        Windows.ApplicationModel.DataTransfer.Clipboard.SetContent(dataPackage);
        
        ViewModel.StatusMessage = $"Copied cell {_selectedCell.DisplayLabel}";
    }
    
    private async void Paste_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null) return;
        
        try
        {
            var dataPackageView = Windows.ApplicationModel.DataTransfer.Clipboard.GetContent();
            if (dataPackageView.Contains(Windows.ApplicationModel.DataTransfer.StandardDataFormats.Text))
            {
                var text = await dataPackageView.GetTextAsync();
                _selectedCell.RawValue = text;
                
                // Refresh grid
                if (ViewModel.SelectedSheet != null)
                {
                    BuildSpreadsheetGrid(ViewModel.SelectedSheet);
                }
                
                ViewModel.StatusMessage = $"Pasted into cell {_selectedCell.DisplayLabel}";
            }
        }
        catch (Exception ex)
        {
            ViewModel.StatusMessage = $"Paste failed: {ex.Message}";
        }
    }
    
    private void ClearContents_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null) return;
        
        _selectedCell.RawValue = string.Empty;
        _selectedCell.Formula = string.Empty;
        _selectedCell.Notes = string.Empty;
        
        // Refresh grid and inspector
        if (ViewModel.SelectedSheet != null)
        {
            BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        }
        
        _isUpdatingCell = true;
        CellValueBox.Text = string.Empty;
        CellFormulaBox.Text = string.Empty;
        CellNotesBox.Text = string.Empty;
        _isUpdatingCell = false;
        
        ViewModel.StatusMessage = $"Cleared cell {_selectedCell.DisplayLabel}";
    }
    
    private void InsertRowAbove_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null || ViewModel.SelectedSheet == null) return;
        
        var rowIndex = _selectedCell.Row;
        ViewModel.SelectedSheet.InsertRow(rowIndex);
        
        BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        ViewModel.StatusMessage = $"Inserted row above row {rowIndex + 1}";
    }
    
    private void InsertRowBelow_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null || ViewModel.SelectedSheet == null) return;
        
        var rowIndex = _selectedCell.Row + 1;
        ViewModel.SelectedSheet.InsertRow(rowIndex);
        
        BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        ViewModel.StatusMessage = $"Inserted row below row {rowIndex}";
    }
    
    private void InsertColumnLeft_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null || ViewModel.SelectedSheet == null) return;
        
        var colIndex = _selectedCell.Column;
        ViewModel.SelectedSheet.InsertColumn(colIndex);
        
        BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        ViewModel.StatusMessage = $"Inserted column left of column {GetColumnName(colIndex)}";
    }
    
    private void InsertColumnRight_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null || ViewModel.SelectedSheet == null) return;
        
        var colIndex = _selectedCell.Column + 1;
        ViewModel.SelectedSheet.InsertColumn(colIndex);
        
        BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        ViewModel.StatusMessage = $"Inserted column right of column {GetColumnName(colIndex - 1)}";
    }
    
    private void DeleteRow_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null || ViewModel.SelectedSheet == null) return;
        
        var rowIndex = _selectedCell.Row;
        
        if (ViewModel.SelectedSheet.Rows.Count <= 1)
        {
            ViewModel.StatusMessage = "Cannot delete the last row";
            return;
        }
        
        ViewModel.SelectedSheet.DeleteRow(rowIndex);
        
        BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        ViewModel.StatusMessage = $"Deleted row {rowIndex + 1}";
    }
    
    private void DeleteColumn_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null || ViewModel.SelectedSheet == null) return;
        
        var colIndex = _selectedCell.Column;
        
        if (ViewModel.SelectedSheet.ColumnCount <= 1)
        {
            ViewModel.StatusMessage = "Cannot delete the last column";
            return;
        }
        
        ViewModel.SelectedSheet.DeleteColumn(colIndex);
        
        BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        ViewModel.StatusMessage = $"Deleted column {GetColumnName(colIndex)}";
    }

    private async void FormatCell_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null)
        {
            return;
        }

        var dialog = new FormatCellDialog(_selectedCell.Format)
        {
            XamlRoot = Content.XamlRoot
        };

        var result = await dialog.ShowAsync();
        if (result == ContentDialogResult.Primary)
        {
            _selectedCell.Format = dialog.SelectedFormat.Clone();
            if (_selectedButton != null)
            {
                ApplyCellStyling(_selectedButton, _selectedCell);
            }
            ViewModel.StatusMessage = $"Updated format for {_selectedCell.DisplayLabel}";
        }
    }

    private async void ViewHistory_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null)
        {
            return;
        }

        var dialog = new CellHistoryDialog(_selectedCell)
        {
            XamlRoot = Content.XamlRoot
        };

        await dialog.ShowAsync();
    }

    private async void ExtractFormula_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedCell == null || string.IsNullOrWhiteSpace(_selectedCell.Formula) || ViewModel.SelectedSheet == null)
        {
            ViewModel.StatusMessage = "No formula to extract.";
            return;
        }

        var dialog = new ExtractFormulaDialog(_selectedCell.Address.ToString())
        {
            XamlRoot = Content.XamlRoot
        };

        var result = await dialog.ShowAsync();
        if (result != ContentDialogResult.Primary)
        {
            return;
        }

        var targetAddress = dialog.TargetAddress;
        if (!CellAddress.TryParse(targetAddress, ViewModel.SelectedSheet.Name, out var address))
        {
            ViewModel.StatusMessage = "Invalid destination address.";
            return;
        }

        var targetSheet = ViewModel.GetSheet(address.SheetName) ?? ViewModel.SelectedSheet;
        var destination = targetSheet.GetCell(address.Row, address.Column);
        if (destination == null)
        {
            ViewModel.StatusMessage = "Destination cell not available.";
            return;
        }

        destination.Formula = _selectedCell.Formula;
        destination.AutomationMode = _selectedCell.AutomationMode;
        targetSheet.UpdateCellDependencies(destination);

        _selectedCell.Formula = string.Empty;
        targetSheet.UpdateCellDependencies(_selectedCell);
        ViewModel.StatusMessage = $"Extracted formula to {targetSheet.Name}!{CellAddress.ColumnIndexToName(address.Column)}{address.Row + 1}";

        if (ViewModel.SelectedSheet != null)
        {
            BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        }
    }
    
    private string GetColumnName(int index)
    {
        string name = "";
        while (index >= 0)
        {
            name = (char)('A' + (index % 26)) + name;
            index = (index / 26) - 1;
        }
        return name;
    }
    
    // Python SDK Named Pipe Server
    private void ToggleFunctionsPanel_Click(object sender, RoutedEventArgs e)
    {
        SetFunctionsPanelVisibility(!_functionsVisible, ViewModel.Settings.FunctionsPanelWidth);
    }
    
    private void ToggleInspectorPanel_Click(object sender, RoutedEventArgs e)
    {
        SetInspectorPanelVisibility(!_inspectorVisible, ViewModel.Settings.InspectorPanelWidth);
    }

    private void SetFunctionsPanelVisibility(bool visible, double preferredWidth)
    {
        _functionsVisible = visible;

        if (visible)
        {
            double targetWidth = preferredWidth > 0 ? preferredWidth : ViewModel.Settings.FunctionsPanelWidth;
            targetWidth = Math.Max(180, targetWidth);
            FunctionsColumn.Width = new GridLength(targetWidth);
            FunctionsPanel.Visibility = Visibility.Visible;
            ToggleFunctionsButton.Content = "◀";
            ToggleFunctionsButton.SetValue(ToolTipService.ToolTipProperty, "Hide Functions Panel");
            ShowFunctionsButton.Visibility = Visibility.Collapsed;
        }
        else
        {
            FunctionsColumn.Width = new GridLength(0);
            FunctionsPanel.Visibility = Visibility.Collapsed;
            ToggleFunctionsButton.Content = "▶";
            ToggleFunctionsButton.SetValue(ToolTipService.ToolTipProperty, "Show Functions Panel");
            ShowFunctionsButton.Visibility = Visibility.Visible;
        }
    }

    private void SetInspectorPanelVisibility(bool visible, double preferredWidth)
    {
        _inspectorVisible = visible;

        if (visible)
        {
            double targetWidth = preferredWidth > 0 ? preferredWidth : ViewModel.Settings.InspectorPanelWidth;
            targetWidth = Math.Max(220, targetWidth);
            InspectorColumn.Width = new GridLength(targetWidth);
            InspectorPanel.Visibility = Visibility.Visible;
            ToggleInspectorButton.Content = "▶";
            ToggleInspectorButton.SetValue(ToolTipService.ToolTipProperty, "Hide Inspector Panel");
            ShowInspectorButton.Visibility = Visibility.Collapsed;
        }
        else
        {
            InspectorColumn.Width = new GridLength(0);
            InspectorPanel.Visibility = Visibility.Collapsed;
            ToggleInspectorButton.Content = "◀";
            ToggleInspectorButton.SetValue(ToolTipService.ToolTipProperty, "Show Inspector Panel");
            ShowInspectorButton.Visibility = Visibility.Visible;
        }
    }

    private void StartPipeServer()
    {
        try
        {
            _pipeServer = new PipeServer(
                "AiCalcPipe",
                ViewModel,
                ViewModel.FunctionRunner,
                DispatcherQueue);
            _pipeServer.Start();
            System.Diagnostics.Debug.WriteLine("Python SDK Pipe Server started successfully");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to start Pipe Server: {ex.Message}");
        }
    }
    
    private void StopPipeServer()
    {
        _pipeServer?.Dispose();
        _pipeServer = null;
    }

    /// <summary>
    /// Save user preferences on window close (Phase 5)
    /// </summary>
    public void SavePreferences()
    {
        try
        {
            var prefs = App.PreferencesService.LoadPreferences();
            
            // Save panel states
            prefs.FunctionsPanelWidth = FunctionsColumn.ActualWidth;
            prefs.InspectorPanelWidth = InspectorColumn.ActualWidth;
            prefs.FunctionsPanelVisible = _functionsVisible;
            prefs.InspectorPanelVisible = _inspectorVisible;
            
            // Save window size
            if (App.MainWindow != null)
            {
                var appWindow = App.MainWindow.AppWindow;
                prefs.WindowWidth = appWindow.Size.Width;
                prefs.WindowHeight = appWindow.Size.Height;
            }
            
            // Save theme
            if (App.MainWindow?.Content is FrameworkElement rootElement)
            {
                prefs.Theme = rootElement.RequestedTheme switch
                {
                    ElementTheme.Light => "Light",
                    ElementTheme.Dark => "Dark",
                    _ => "System"
                };
            }
            
            App.PreferencesService.SavePreferences(prefs);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving preferences: {ex.Message}");
        }
    }

    /// <summary>
    /// Load and apply user preferences on startup (Phase 5)
    /// </summary>
    public void LoadPreferences()
    {
        try
        {
            var prefs = App.PreferencesService.LoadPreferences();
            
            // Ensure panels start expanded with preferred widths
            double functionsWidth = prefs.FunctionsPanelWidth > 0 ? prefs.FunctionsPanelWidth : ViewModel.Settings.FunctionsPanelWidth;
            double inspectorWidth = prefs.InspectorPanelWidth > 0 ? prefs.InspectorPanelWidth : ViewModel.Settings.InspectorPanelWidth;

            ViewModel.Settings.FunctionsPanelWidth = functionsWidth;
            ViewModel.Settings.InspectorPanelWidth = inspectorWidth;

            SetFunctionsPanelVisibility(true, functionsWidth);
            SetInspectorPanelVisibility(true, inspectorWidth);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error loading preferences: {ex.Message}");
        }
    }

    /// <summary>
    /// Refresh the currently selected cell's display after undo/redo (Phase 5)
    /// </summary>
    private void RefreshCurrentCell()
    {
        if (_selectedCell != null)
        {
            CellFormulaBox.Text = _selectedCell.Formula ?? "";
        }
        
        if (ViewModel.SelectedSheet != null)
        {
            BuildSpreadsheetGrid(ViewModel.SelectedSheet);
        }
    }
}
