# AiCalc Python SDK

Python SDK for interacting with AiCalc application via Named Pipes IPC.

## Status
🟢 **Alpha** - IPC bridge implemented, basic operations functional.

## Installation

```bash
cd python-sdk
pip install -e .
```

This will install the SDK in development mode along with required dependencies (pywin32 on Windows).

## Quick Start

Make sure AiCalc is running before using the SDK.

```python
from aicalc_sdk import connect

# Connect to AiCalc
with connect() as client:
    # Get cell value
    value = client.get_value("A1")
    print(f"A1 = {value}")
    
    # Set cell value
    client.set_value("A1", 42)
    
    # Set value with sheet reference
    client.set_value("Sheet1!B2", "Hello")
    
    # Get range of values
    values = client.get_range("A1:B10")
    
    # Execute AiCalc function
    result = client.run_function("SUM", 1, 2, 3, 4, 5)
    
    # Get list of sheets
    sheets = client.get_sheets()
    for sheet in sheets:
        print(f"Sheet: {sheet['name']}, Size: {sheet['row_count']}x{sheet['column_count']}")
```

## Testing

Run the test script to verify the SDK is working:

```bash
# Make sure AiCalc is running first!
python test_connection.py
```

## API Reference

### `connect(pipe_name='AiCalc_Bridge')`
Create and connect to AiCalc. Returns an `AiCalcClient` instance.

### `AiCalcClient`

#### `get_value(cell_ref: str) -> Any`
Get value from a cell.
- `cell_ref`: Cell reference (e.g., 'A1', 'Sheet1!B2')
- Returns: Cell value (display value)

#### `set_value(cell_ref: str, value: Any) -> bool`
Set value of a cell.
- `cell_ref`: Cell reference (e.g., 'A1', 'Sheet1!B2')
- `value`: Value to set
- Returns: True if successful

#### `get_range(range_ref: str) -> List[List[Any]]`
Get values from a range of cells.
- `range_ref`: Range reference (e.g., 'A1:B10', 'Sheet1!A1:C5')
- Returns: 2D list of cell values

#### `run_function(function_name: str, *args) -> Any`
Execute an AiCalc function.
- `function_name`: Name of the function to execute
- `*args`: Function arguments
- Returns: Function result

#### `get_sheets() -> List[Dict[str, Any]]`
Get list of sheets in the workbook.
- Returns: List of sheet information dictionaries

## Creating Custom Functions (Coming Soon)

```python
from aicalc_sdk import aicalc_function

@aicalc_function(name="DOUBLE")
def double(x: float) -> float:
    return x * 2
```

## Architecture

The SDK uses Named Pipes for IPC communication with AiCalc:
- **Pipe Name**: `\\.\pipe\AiCalc_Bridge`
- **Protocol**: JSON-based request/response
- **Transport**: Named Pipes (Windows)

## Requirements

- Python 3.8+
- Windows (Named Pipes)
- pywin32 (automatically installed)
