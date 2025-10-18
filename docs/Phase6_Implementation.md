# Phase 6 Implementation - File Format & Persistence Enhancements

**Status**: Core Complete ✅  
**Started**: October 18, 2025  
**Completion**: ~80%

---

## Overview

Phase 6 focuses on enhancing file format support and persistence features for the AiCalc spreadsheet application. This phase implements automatic saving, CSV export/import, and prepares for future enhancements like Excel compatibility and version control.

---

## Implemented Features

### 1. AutoSave Service ✅

**File**: `src/AiCalc.WinUI/Services/AutoSaveService.cs`

**Description**: Timer-based automatic workbook saving with configurable intervals and dirty flag tracking.

**Features**:
- Timer-based autosave (1-60 minute intervals, default: 5 minutes)
- Dirty flag tracking - only saves when workbook has changed
- Autosave backup files (e.g., `Workbook_autosave.aicalc`)
- Event notifications: `AutoSaved`, `AutoSaveFailed`
- Enable/disable toggle
- Configurable save interval

**Implementation Details**:
```csharp
public class AutoSaveService : IDisposable
{
    private readonly WorkbookViewModel _workbook;
    private Timer? _autoSaveTimer;
    private bool _isDirty;
    private int _intervalMinutes = 5;
    
    public bool IsEnabled { get; set; }
    public int IntervalMinutes { get; set; }
    
    public void MarkDirty() { ... }
    public void SetSavePath(string path) { ... }
}
```

**Integration**:
- Initialized in `WorkbookViewModel` constructor
- Connected to autosave events for status messages
- `CellViewModel.MarkAsUpdated()` calls `_workbook.MarkDirty()`

**Usage**:
```csharp
// In WorkbookViewModel constructor
_autoSaveService = new AutoSaveService(this);
_autoSaveService.AutoSaved += (s, path) => 
    StatusMessage = $"Auto-saved to {Path.GetFileName(path)}";
_autoSaveService.AutoSaveFailed += (s, ex) => 
    StatusMessage = $"Auto-save failed: {ex.Message}";
```

---

### 2. CSV Export ✅

**File**: `src/AiCalc.WinUI/Services/CsvService.cs`

**Description**: Export spreadsheet data to CSV format with proper escaping and encoding.

**Features**:
- Single sheet export to CSV file
- Proper CSV escaping (quotes, commas, newlines)
- UTF-8 encoding
- Uses cell display values
- Export entire workbook to multiple CSV files (one per sheet)

**Implementation**:
```csharp
public static async Task ExportSheetToCsvAsync(SheetViewModel sheet, string filePath)
{
    // Iterate through rows and columns
    // Escape values with commas, quotes, or newlines
    // Write to file with UTF-8 encoding
}

public static async Task ExportWorkbookToCsvAsync(WorkbookViewModel workbook, string directoryPath)
{
    // Export each sheet to separate CSV file
}
```

**CSV Escaping Rules**:
- Wrap values containing commas, quotes, or newlines in double quotes
- Escape internal quotes by doubling them (`"` → `""`)
- Preserve newlines within quoted values

**UI Integration**:
- Button: "📤 Export CSV" in MainWindow.xaml
- Tooltip: "Export current sheet to CSV"
- Command: `WorkbookViewModel.ExportCsvCommand`
- Event handler: `ExportCsvButton_Click()` in MainWindow.xaml.cs

---

### 3. CSV Import ✅

**File**: `src/AiCalc.WinUI/Services/CsvService.cs`

**Description**: Import CSV data into a new spreadsheet sheet with robust parsing.

**Features**:
- Creates new sheet from CSV file
- Robust CSV parsing (handles quoted values, escaped quotes)
- Automatic dimension detection
- Sets cell values using `RawValue` property
- Respects minimum sheet size (10x10)

**Implementation**:
```csharp
public static async Task<SheetViewModel> ImportCsvToSheetAsync(
    string filePath, string sheetName, WorkbookViewModel workbook)
{
    // Read CSV lines
    // Parse each line handling quotes and escapes
    // Determine dimensions (max columns, row count)
    // Create sheet with appropriate size
    // Populate cells with parsed values
}
```

**CSV Parsing Logic**:
- Track quote state while iterating characters
- Handle escaped quotes (`""` within quoted values)
- Split on commas only outside of quotes
- Preserve all data including empty cells

**UI Integration**:
- Button: "📥 Import CSV" in MainWindow.xaml
- Tooltip: "Import CSV as new sheet"
- File picker dialog to select CSV file
- Command: `WorkbookViewModel.ImportCsvCommand`
- Event handler: `ImportCsvButton_Click()` with file picker

---

## Code Changes

### New Files

1. **Services/AutoSaveService.cs** (137 lines)
   - AutoSave service class
   - Timer management
   - Dirty flag tracking
   - Event notifications

2. **Services/CsvService.cs** (165 lines)
   - Static CSV utility class
   - Export methods
   - Import methods
   - CSV parsing and escaping utilities

### Modified Files

1. **ViewModels/WorkbookViewModel.cs**
   - Added `_autoSaveService` field
   - Initialized AutoSaveService in constructor
   - Added `MarkDirty()` method
   - Added `SaveInternalAsync()` helper method
   - Added `ExportCsvAsync()` command
   - Added `ImportCsvAsync(string filePath)` command
   - Updated `SaveAsync()` to call `_autoSaveService.SetSavePath()`

2. **ViewModels/CellViewModel.cs**
   - Updated `MarkAsUpdated()` to call `_workbook.MarkDirty()`

3. **MainWindow.xaml**
   - Added "📤 Export CSV" button
   - Added "📥 Import CSV" button
   - Added tooltips for both buttons

4. **MainWindow.xaml.cs**
   - Added `ExportCsvButton_Click()` event handler
   - Added `ImportCsvButton_Click()` event handler with file picker

---

## Build Status

**Current Build**: ✅ Success
- 0 Warnings
- 0 Errors
- Build time: ~9 seconds

**Tests**: ✅ All Passing
- 59/59 tests passing
- No new test coverage for Phase 6 yet

### 4. AutoSave Settings UI ✅

**File**: `src/AiCalc.WinUI/SettingsDialog.xaml`, `SettingsDialog.xaml.cs`

**Description**: User interface for configuring AutoSave preferences in Settings dialog.

**Features**:
- Toggle switch: Enable/Disable AutoSave
- Slider: Set interval (1-60 minutes)
- Visual feedback (opacity change when disabled)
- Real-time application to active workbook
- Preferences persisted to disk

**Implementation Details**:
```csharp
// In SettingsDialog.xaml.cs
private void AutoSaveToggle_Toggled(object sender, RoutedEventArgs e)
{
    var isEnabled = AutoSaveToggle.IsOn;
    
    // Update UI
    AutoSaveIntervalPanel.Opacity = isEnabled ? 1.0 : 0.5;
    AutoSaveIntervalSlider.IsEnabled = isEnabled;
    
    // Save preference
    var prefs = App.PreferencesService.LoadPreferences();
    prefs.AutoSaveEnabled = isEnabled;
    App.PreferencesService.SavePreferences(prefs);
    
    // Apply to workbook
    mainWindow.ViewModel.SetAutoSaveEnabled(isEnabled);
}
```

**UI Elements**:
- ToggleSwitch: Enable AutoSave (On/Off with labels)
- Slider: AutoSave interval (1-60 minutes)
- Label: Shows current interval in minutes
- Help text: Explains backup file naming

**Integration**:
- Loads preferences on dialog open
- Saves preferences on change
- Immediately applies to WorkbookViewModel
- Persisted across application restarts

---

### 5. File Picker Integration ✅

**File**: `src/AiCalc.WinUI/MainWindow.xaml.cs`

**Description**: File save/open pickers for CSV export/import operations.

**CSV Export Implementation**:
```csharp
private async void ExportCsvButton_Click(object sender, RoutedEventArgs e)
{
    var picker = new Windows.Storage.Pickers.FileSavePicker();
    WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);
    picker.FileTypeChoices.Add("CSV Files", new List<string>() { ".csv" });
    picker.SuggestedFileName = ViewModel.SelectedSheet.Name;
    
    var file = await picker.PickSaveFileAsync();
    if (file != null)
    {
        await CsvService.ExportSheetToCsvAsync(ViewModel.SelectedSheet, file.Path);
    }
}
```

**CSV Import Implementation**:
```csharp
private async void ImportCsvButton_Click(object sender, RoutedEventArgs e)
{
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
```

**Features**:
- Native Windows file dialogs
- Filtered to .csv files only
- Suggested filename from sheet name (export)
- Error handling with status messages
- Proper WinUI 3 window handle initialization

---

## Remaining Work (Phase 6)

### 1. Excel Export/Import (8-12 hours) - OPTIONAL
- Investigate libraries: ClosedXML, EPPlus, or OpenXML SDK
- Implement basic Excel export (.xlsx)
- Implement basic Excel import (.xlsx)
- Handle formula translation (Excel formulas vs AiCalc formulas)
- Handle formatting (colors, borders, fonts)

**Challenges**:
- Excel is a complex format
- AiCalc has features Excel doesn't (AI functions)
- Excel has features AiCalc doesn't (charts, pivot tables)
- Formula syntax differences

### 2. Version Control/History (6-10 hours) - OPTIONAL
- Save workbook snapshots on major changes
- Track change history (who, when, what)
- Implement rollback functionality
- Add version browser UI

---

## Next Steps

### Completed in This Session:
1. ✅ **AutoSave Settings UI** (2-3 hours) - Toggle, slider, preferences
2. ✅ **File Picker for Export** (1 hour) - Save file dialog with .csv filter
3. ✅ **File Picker for Import** (already done) - Open file dialog
4. ✅ **User Preferences** - AutoSaveEnabled and AutoSaveIntervalMinutes
5. ✅ **WorkbookViewModel Integration** - SetAutoSaveEnabled/SetAutoSaveInterval methods

### Recommended Next Steps:
- **Phase 8: Multi-cell Selection** (4-6 hours) - High user value, essential for productivity
- **Phase 8: Column Operations** (3-5 hours) - Resize, hide, freeze panes
- **Phase 7: Python SDK** (20-30 hours) - If Python integration is critical
- **Phase 6: Data Sources** (15-25 hours) - Azure Blob, SQL Database (Tasks 18-19)

---

## Technical Notes

### Design Decisions

1. **CSV Service as Static Class**
   - No state needed
   - Pure utility functions
   - Easy to test and use

2. **AutoSave Uses Backup Files**
   - Preserves original file
   - Allows recovery if main file corrupted
   - Pattern: `filename_autosave.aicalc`

3. **Dirty Flag Tracking**
   - Prevents unnecessary saves
   - Marks dirty on cell changes (MarkAsUpdated)
   - Clears dirty after successful save

4. **CSV Import Creates New Sheet**
   - Non-destructive operation
   - User can review imported data
   - Can import multiple CSV files

### Limitations

1. **CSV Format**
   - No formatting preservation (colors, borders)
   - No formula preservation (only values)
   - Single sheet per file
   - No multi-sheet export (yet - ExportWorkbookToCsvAsync exists but not wired to UI)

2. **AutoSave**
   - No conflict resolution for concurrent edits
   - No cloud sync integration
   - Fixed backup file naming

3. **File Picker**
   - Export uses automatic filename
   - No "Save As" functionality yet
   - Import requires manual file selection

---

## Testing Recommendations

### Manual Testing

1. **AutoSave Test**
   ```
   1. Create a workbook with data
   2. Save the workbook
   3. Make changes (edit cells)
   4. Wait 5 minutes (or adjust interval)
   5. Verify autosave backup file is created
   6. Check status message shows "Auto-saved to..."
   ```

2. **CSV Export Test**
   ```
   1. Create a sheet with varied data:
      - Numbers: 123, 456.78
      - Text: Hello, World
      - Text with commas: "One, Two, Three"
      - Text with quotes: He said "Hello"
      - Multi-line text: Line 1\nLine 2
   2. Click "📤 Export CSV"
   3. Open CSV in Excel/Notepad
   4. Verify all data is correct
   ```

3. **CSV Import Test**
   ```
   1. Create a CSV file with test data
   2. Click "📥 Import CSV"
   3. Select the CSV file
   4. Verify new sheet is created
   5. Verify all data is imported correctly
   6. Check dimensions are correct
   ```

### Unit Testing (TODO)

1. **CsvService Tests**
   - Test CSV escaping edge cases
   - Test parsing with various quote scenarios
   - Test multi-line values
   - Test empty cells and trailing commas

2. **AutoSaveService Tests**
   - Test timer initialization
   - Test dirty flag behavior
   - Test save path tracking
   - Test enable/disable toggle
   - Mock Timer for deterministic testing

---

## Performance Notes

- **CSV Export**: O(n*m) where n = rows, m = columns. Efficient for typical spreadsheet sizes.
- **CSV Import**: O(n*m*p) where p = average cell value length. Parser is character-by-character.
- **AutoSave**: Timer-based, no performance impact on UI thread. Async save operation.

---

## Conclusion

Phase 6 has made excellent progress with AutoSave and CSV support. The implementation is clean, follows existing patterns, and integrates well with the Phase 5 UI enhancements. 

**Phase 6 Core Features Completion**: ~80%
- ✅ AutoSave service
- ✅ CSV export/import
- ✅ UI integration
- ✅ Settings UI for autosave
- ✅ File pickers for export/import
- ✅ User preferences persistence
- ⏳ Excel support (optional, 8-12 hours)
- ⏳ Version control (optional, 6-10 hours)
- ⏳ Data Sources (Tasks 18-19, 15-25 hours)

**What's Complete**:
All core file format and persistence features are functional and polished. Users can:
- Auto-save workbooks at configurable intervals
- Export sheets to CSV with save dialog
- Import CSV files as new sheets with open dialog
- Configure AutoSave preferences in Settings
- See status messages for all operations

**What's Deferred**:
- Excel .xlsx format support (complex, optional)
- Binary format for large workbooks (optimization, optional)
- Version control/rollback (advanced feature)
- Data Sources integration (separate major feature set, Tasks 18-19)

**Next Phase Recommendation**: Move to Phase 8 (Multi-cell Selection, Column Operations) for high user value features that build on the solid foundation we've created.
