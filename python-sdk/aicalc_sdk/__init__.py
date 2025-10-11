"""
AiCalc Python SDK

This SDK provides a Python interface to interact with AiCalc spreadsheet application.
"""

__version__ = "0.1.0"
__author__ = "AiCalc Team"

from .client import connect, AiCalcClient
from .decorators import aicalc_function
from .types import CellValue, CellType, AutomationMode

__all__ = [
    'connect',
    'AiCalcClient',
    'aicalc_function',
    'CellValue',
    'CellType',
    'AutomationMode',
]
