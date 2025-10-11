"""Type definitions for AiCalc SDK"""

from enum import Enum
from dataclasses import dataclass
from typing import Optional

class CellType(Enum):
    """Cell object types supported by AiCalc"""
    EMPTY = "Empty"
    NUMBER = "Number"
    TEXT = "Text"
    IMAGE = "Image"
    TABLE = "Table"

class AutomationMode(Enum):
    """Cell automation modes"""
    MANUAL = "Manual"
    AUTO_ON_OPEN = "AutoOnOpen"
    AUTO_ON_DEPENDENCY_CHANGE = "AutoOnDependencyChange"

@dataclass
class CellValue:
    """Represents a cell value with type information"""
    object_type: CellType
    serialized_value: Optional[str] = None
    display_value: Optional[str] = None
    
    @classmethod
    def empty(cls) -> 'CellValue':
        return cls(CellType.EMPTY, None, "")
    
    def __str__(self) -> str:
        return self.display_value or ""
