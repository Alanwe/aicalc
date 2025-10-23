# Formula Range Selection Feature

## Overview
This feature allows users to select cell ranges directly from the spreadsheet while editing formulas, similar to Excel's behavior. When a user types a function like `=SUM()` and positions the cursor between the parentheses, they can click and drag on the spreadsheet to select a range, which is then automatically inserted into the formula.

## How It Works

### User Workflow
1. User types `=` to start a formula
2. User types a function name (e.g., `SUM`)
3. Autocomplete flyout appears - user selects the function
4. Function is inserted with cursor positioned inside parentheses: `=SUM(|)`
5. User clicks on a cell in the spreadsheet and drags to another cell
6. The selected range is highlighted visually during the drag
7. When the user releases the mouse, the range reference (e.g., `A1:B5`) is inserted into the formula
8. User can continue typing or select additional ranges

### Key Features
- **Visual Feedback**: Selected range is highlighted using the formula dependency chain color (blue)
- **Smart Reference Building**: 
  - Single cell: `A1`
  - Range: `A1:B5`
  - Cross-sheet: `Sheet2!A1:B5`
- **Automatic Comma Insertion**: If the cursor is not after `(`, `,`, or `=`, a comma is automatically added before the reference
- **Focus Preservation**: After range insertion, focus returns to the edit box with cursor positioned after the inserted range
- **Multi-range Support**: Users can insert multiple ranges by clicking and dragging again

## Implementation Details

### New Fields
```csharp
private bool _isSelectingRange = false;           // Currently dragging to select
private CellViewModel? _rangeSelectionStart;       // Starting cell of drag
private CellViewModel? _rangeSelectionEnd;         // Current cell under pointer
```

### Key Methods

#### `CellButton_PointerPressed`
- Detects when user clicks on a cell during formula editing
- Checks if `_isPickingFormulaReference` is true (cursor is in formula)
- Starts range selection mode
- Captures pointer for drag tracking
- Highlights the starting cell

#### `CellButton_PointerMoved`
- Tracks pointer movement during drag
- Updates `_rangeSelectionEnd` to current cell
- Calls `HighlightRangeSelection()` to update visual feedback

#### `CellButton_PointerReleased`
- Completes the range selection
- Releases pointer capture
- Calls `InsertRangeReferenceIntoFormula()` to insert the range
- Clears selection state and highlights

#### `InsertRangeReferenceIntoFormula`
- Builds range reference string:
  - Uses `BuildCellReferenceString()` for start and end cells
  - Formats as single cell or range (`A1` vs `A1:B5`)
  - Handles cross-sheet references automatically
- Determines active edit box (CellEditBox or CellFormulaBox)
- Inserts reference at current cursor position
- Adds comma separator if needed
- Updates both edit boxes synchronously
- Restores focus and cursor position

#### `HighlightRangeSelection`
- Calculates rectangular bounds from start to end cells
- Iterates through all cells in the range
- Applies `CellVisualState.InDependencyChain` to each cell
- Provides real-time visual feedback during drag

### Integration Points

#### Existing Picking Mode
The feature leverages the existing `_isPickingFormulaReference` flag:
- Set to `true` in `InsertFunctionSuggestion()` after a function is inserted
- Checked in `CellButton_PointerPressed` to enable range selection
- Set to `false` when editing is committed or cancelled

#### Pointer Event Chain
Cell buttons now handle three pointer events:
```csharp
button.PointerPressed += CellButton_PointerPressed;
button.PointerMoved += CellButton_PointerMoved;
button.PointerReleased += CellButton_PointerReleased;
```

#### Edit Box Synchronization
Both `CellEditBox` and `CellFormulaBox` are updated when a range is inserted:
- Prevents desync between the two input controls
- Ensures formula is properly stored in the cell's Formula property
- Maintains undo/redo compatibility

## Edge Cases Handled

1. **Single Cell Selection**: If start and end are the same, only a single cell reference is inserted
2. **Cross-Sheet References**: When selecting cells from a different sheet, the sheet name is included in the reference
3. **Comma Insertion**: Smart detection of what character precedes the cursor to determine if a comma is needed
4. **Focus Management**: Async focus restoration ensures the cursor returns to the correct position
5. **Highlight Cleanup**: Previous highlights are cleared before new ones are applied

## Testing Scenarios

### Basic Range Selection
1. Type `=SUM(` in a cell
2. Click and drag from A1 to B5
3. Verify `=SUM(A1:B5)` appears in the formula

### Multiple Parameters
1. Type `=IF(A1>5,`
2. Click and drag to select B1:B10
3. Verify `=IF(A1>5,B1:B10)` appears
4. Type `,`
5. Select another range C1:C10
6. Verify `=IF(A1>5,B1:B10,C1:C10)` appears

### Cross-Sheet Selection
1. Type `=SUM(` in Sheet1
2. Switch to Sheet2 and select range A1:A10
3. Verify `=SUM(Sheet2!A1:A10)` appears

### Single Cell
1. Type `=SQRT(`
2. Click on cell A1 (don't drag)
3. Verify `=SQRT(A1)` appears

## Known Limitations
- Range selection only works when `_isPickingFormulaReference` is true (after a function has been inserted)
- No keyboard-based range selection (Shift+Arrow keys) yet
- Cannot select non-contiguous ranges (e.g., A1:A5, C1:C5)

## Future Enhancements
- Add keyboard support for range selection
- Support for non-contiguous range selection with Ctrl+Click
- Named range selection
- Formula parameter hints showing expected argument types
- Auto-complete for named ranges
