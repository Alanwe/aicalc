"""Main client for interacting with AiCalc"""

from typing import Optional, List, Any
from .types import CellValue

class AiCalcClient:
    """Client for interacting with AiCalc application."""
    
    def __init__(self, connection_name: str = "aicalc_default"):
        self.connection_name = connection_name
        self._connected = False
        
    def connect(self, timeout: int = 5000) -> bool:
        """Connect to AiCalc application."""
        self._connected = True
        return True
    
    def disconnect(self) -> None:
        """Disconnect from AiCalc application"""
        self._connected = False
    
    def is_connected(self) -> bool:
        """Check if connected to AiCalc"""
        return self._connected
    
    def get_value(self, cell_ref: str) -> Any:
        """Get value from a cell."""
        if not self._connected:
            raise ConnectionError("Not connected to AiCalc")
        return None
    
    def set_value(self, cell_ref: str, value: Any) -> bool:
        """Set value of a cell."""
        if not self._connected:
            raise ConnectionError("Not connected to AiCalc")
        return True
    
    def __enter__(self):
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()

def connect(connection_name: str = "aicalc_default") -> AiCalcClient:
    """Connect to AiCalc application."""
    client = AiCalcClient(connection_name)
    client.connect()
    return client
