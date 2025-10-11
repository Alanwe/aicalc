"""
AiCalc Python SDK
Main module for connecting to and interacting with AiCalc application
"""

from .client import Workbook, connect
from .models import CellAddress, CellValue, CellType

__version__ = "0.1.0"
__all__ = [
    "Workbook",
    "connect",
    "CellAddress",
    "CellValue",
    "CellType",
]
