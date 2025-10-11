# AiCalc Python SDK

Python SDK for interacting with AiCalc spreadsheet application.

## Status
⏳ **In Development** - Basic scaffolding complete, IPC implementation pending.

## Installation (Future)
```bash
pip install aicalc-sdk
```

## Quick Start
```python
import aicalc_sdk as aicalc

# Connect to AiCalc
workbook = aicalc.connect()

# Read/write cells
value = workbook.get_value('Sheet1!A1')
workbook.set_value('Sheet1!B1', 42)
```

## Creating Custom Functions
```python
from aicalc_sdk import aicalc_function

@aicalc_function(name="DOUBLE")
def double(x: float) -> float:
    return x * 2
```
