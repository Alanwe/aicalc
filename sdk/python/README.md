# AiCalc Python SDK

Python SDK for interacting with AiCalc - the AI-Native Spreadsheet Application.

## Installation

```bash
pip install aicalc-sdk
```

## Quick Start

```python
from aicalc import connect, CellAddress

# Connect to running AiCalc instance
workbook = connect()

# Get cell value
value = workbook.get_value("A1")
print(f"A1 = {value}")

# Set cell value
workbook.set_value("B1", "Hello from Python!")

# Set formula
workbook.set_formula("C1", "=SUM(A1:A10)")

# Get range as pandas DataFrame
df = workbook.get_range("A1:C10")
print(df)

# Run AiCalc function
result = workbook.run_function("TEXT_TO_IMAGE", "A beautiful sunset over mountains")
workbook.set_value("D1", result)
```

## Features

- **Cell Operations**: Read and write cell values and formulas
- **Range Operations**: Work with ranges of cells, export to pandas DataFrame
- **Function Execution**: Execute AiCalc functions from Python
- **Real-time Updates**: Subscribe to cell change events
- **Type Safety**: Full type hints for better IDE support

## Communication

The SDK uses Named Pipes (Windows) for IPC communication with the AiCalc application. The connection is established automatically when you call `connect()`.

## Requirements

- Python 3.8+
- AiCalc application running
- Windows OS (for named pipe support)

## API Reference

### Connection

```python
workbook = connect(
    pipe_name: str = "AiCalcPipe",  # Named pipe name
    timeout: int = 5000              # Connection timeout in ms
) -> Workbook
```

### Workbook Methods

```python
# Get cell value
value = workbook.get_value(cell_ref: str) -> Any

# Set cell value
workbook.set_value(cell_ref: str, value: Any) -> None

# Get/Set formula
formula = workbook.get_formula(cell_ref: str) -> str
workbook.set_formula(cell_ref: str, formula: str) -> None

# Get range as DataFrame
df = workbook.get_range(range_ref: str) -> pd.DataFrame

# Run function
result = workbook.run_function(func_name: str, *args) -> Any

# Subscribe to changes
workbook.on_cell_changed(cell_ref: str, callback: Callable) -> None
```

## Development

```bash
# Install development dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Format code
black src/

# Type checking
mypy src/
```

## License

MIT License - see LICENSE file for details
