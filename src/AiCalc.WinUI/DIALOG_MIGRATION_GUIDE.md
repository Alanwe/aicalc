# Dialog Migration Guide

## Problem with ContentDialog
ContentDialog has internal constraints that make it difficult to control sizing reliably:
- MaxWidth constraints from styles override explicit Width settings
- Nested content can't force dialog expansion
- Fighting with WinUI's internal layout logic wastes time

## Solution: DialogWindowBase

Use `DialogWindowBase` for complex dialogs that need:
- Predictable, controllable sizing
- Wide layouts (>800px)
- Complex multi-column layouts
- No truncation issues

### Basic Usage

```csharp
public sealed partial class MyDialog : DialogWindowBase
{
    public MyDialog()
    {
        // Set dialog properties BEFORE calling constructor
        DialogTitle = "My Dialog";
        DialogWidth = 1100;
        DialogHeight = 600;
        PrimaryButtonText = "Save";
        CloseButtonText = "Cancel";
        
        InitializeComponent(); // This loads your XAML content
        
        // Set the content (your XAML root element)
        SetDialogContent(RootGrid); // RootGrid is the x:Name of your root element
    }
    
    protected override bool OnPrimaryButtonClick()
    {
        // Add validation logic here
        if (!IsValid())
        {
            return false; // Keep dialog open
        }
        
        // Save logic
        return true; // Close dialog
    }
}
```

### XAML Changes

Instead of `<ContentDialog>`, use a simple Grid or StackPanel:

**Before (ContentDialog):**
```xaml
<ContentDialog
    x:Class="AiCalc.MyDialog"
    Title="My Dialog"
    PrimaryButtonText="Save">
    
    <Grid>
        <!-- content -->
    </Grid>
</ContentDialog>
```

**After (DialogWindowBase):**
```xaml
<!-- MyDialog.xaml - Just the content, no window wrapper -->
<Grid x:Name="RootGrid"
      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      Padding="12">
    
    <!-- Your content here -->
    <!-- No need for ScrollViewer, DialogWindowBase provides it -->
    
</Grid>
```

### Code-Behind Changes

**Before:**
```csharp
public sealed partial class MyDialog : ContentDialog
{
    private void ContentDialog_PrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        // Save logic
    }
}
```

**After:**
```csharp
public sealed partial class MyDialog : DialogWindowBase
{
    public MyDialog()
    {
        DialogTitle = "My Dialog";
        DialogWidth = 1100;
        DialogHeight = 600;
        PrimaryButtonText = "Save";
        CloseButtonText = "Cancel";
        
        InitializeComponent();
        SetDialogContent(RootGrid);
    }
    
    protected override bool OnPrimaryButtonClick()
    {
        // Save logic
        return true; // or false to keep dialog open
    }
}
```

### Showing the Dialog

**Before:**
```csharp
var dialog = new MyDialog();
var result = await dialog.ShowAsync();
if (result == ContentDialogResult.Primary)
{
    // Handle save
}
```

**After:**
```csharp
var dialog = new MyDialog();
dialog.Closed += (sender, result) =>
{
    if (result == DialogWindowBase.DialogResult.Primary)
    {
        // Handle save
    }
};
dialog.Activate(); // Show the window
```

## When to Use What

### Use DialogWindowBase for:
- ✅ Wide dialogs (>800px)
- ✅ Complex multi-column layouts
- ✅ Dialogs where sizing must be precise
- ✅ Data entry forms with many fields

### Keep ContentDialog for:
- ✅ Simple confirmation dialogs
- ✅ Small, narrow dialogs (<600px)
- ✅ Quick yes/no prompts
- ✅ Simple message boxes

## Benefits

1. **Predictable Sizing**: Width and height work exactly as specified
2. **No Truncation**: Content fits properly without fighting layout constraints
3. **Reusable**: Set properties once, works consistently
4. **Flexible**: Full control over styling and behavior
5. **Fast Development**: Stop wasting time fighting with ContentDialog

## Migration Priority

Migrate in this order:
1. **DataSourcesDialog** - Wide, complex, currently problematic ✅ High Priority
2. **SettingsDialog** - Large but ContentDialog works OK (migrate if issues arise)
3. Keep simple dialogs as ContentDialog unless problems occur
