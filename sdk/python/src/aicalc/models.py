"""
Data models for AiCalc SDK
"""

from enum import Enum
from dataclasses import dataclass
from typing import Any, Optional


class CellType(Enum):
    """Cell object types supported by AiCalc"""
    EMPTY = "Empty"
    TEXT = "Text"
    NUMBER = "Number"
    IMAGE = "Image"
    VIDEO = "Video"
    TABLE = "Table"
    DIRECTORY = "Directory"
    FILE = "File"
    PDF = "Pdf"
    PDF_PAGE = "PdfPage"
    MARKDOWN = "Markdown"
    JSON = "Json"
    XML = "Xml"
    CHART = "Chart"
    CODE_PYTHON = "Code-Python"
    CODE_CSHARP = "Code-CSharp"
    CODE_JAVASCRIPT = "Code-JavaScript"


@dataclass
class CellAddress:
    """Represents a cell address in the spreadsheet"""
    sheet: str
    row: int
    column: int
    
    @classmethod
    def parse(cls, ref: str) -> "CellAddress":
        """
        Parse cell reference string like "A1", "Sheet1!B2"
        
        Args:
            ref: Cell reference string
            
        Returns:
            CellAddress instance
            
        Examples:
            >>> CellAddress.parse("A1")
            CellAddress(sheet="Sheet1", row=0, column=0)
            >>> CellAddress.parse("Sheet2!C5")
            CellAddress(sheet="Sheet2", row=4, column=2)
        """
        if "!" in ref:
            sheet, cell = ref.split("!", 1)
        else:
            sheet = "Sheet1"
            cell = ref
        
        # Parse column (letters) and row (numbers)
        col_str = ""
        row_str = ""
        for char in cell:
            if char.isalpha():
                col_str += char
            elif char.isdigit():
                row_str += char
        
        # Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, ...)
        col = 0
        for char in col_str.upper():
            col = col * 26 + (ord(char) - ord('A') + 1)
        col -= 1  # Convert to 0-based index
        
        row = int(row_str) - 1  # Convert to 0-based index
        
        return cls(sheet=sheet, row=row, column=col)
    
    def to_string(self) -> str:
        """
        Convert to cell reference string
        
        Returns:
            Cell reference like "A1" or "Sheet1!B2"
        """
        # Convert column index to letters
        col = self.column + 1
        col_str = ""
        while col > 0:
            col -= 1
            col_str = chr(ord('A') + (col % 26)) + col_str
            col //= 26
        
        row_str = str(self.row + 1)
        
        if self.sheet == "Sheet1":
            return f"{col_str}{row_str}"
        else:
            return f"{self.sheet}!{col_str}{row_str}"


@dataclass
class CellValue:
    """Represents a cell value with its type"""
    value: Any
    cell_type: CellType
    raw_value: Optional[str] = None
    formula: Optional[str] = None
    notes: Optional[str] = None
    
    def __repr__(self) -> str:
        return f"CellValue(type={self.cell_type.value}, value={self.value!r})"
