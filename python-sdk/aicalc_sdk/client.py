"""Main client for interacting with AiCalc"""

import json
import win32pipe
import win32file
import pywintypes
from typing import Optional, List, Any, Dict
from .types import CellValue, CellType

class AiCalcClient:
    """Client for interacting with AiCalc application via Named Pipes."""
    
    def __init__(self, pipe_name: str = "AiCalc_Bridge"):
        self.pipe_name = f"\\\\.\\pipe\\{pipe_name}"
        self._pipe_handle = None
        self._connected = False
        
    def connect(self, timeout: int = 5000) -> bool:
        """Connect to AiCalc application."""
        try:
            self._pipe_handle = win32file.CreateFile(
                self.pipe_name,
                win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                0,
                None,
                win32file.OPEN_EXISTING,
                0,
                None
            )
            
            # Server already set pipe to byte mode, don't change it
            
            # Test connection with ping
            response = self._send_command({"command": "ping"})
            if response.get("success") and response.get("data") == "pong":
                self._connected = True
                return True
            return False
        except pywintypes.error as e:
            raise ConnectionError(f"Failed to connect to AiCalc: {e}\nMake sure AiCalc is running and the Python bridge service has started.")
    
    def disconnect(self) -> None:
        """Disconnect from AiCalc application"""
        if self._pipe_handle:
            win32file.CloseHandle(self._pipe_handle)
            self._pipe_handle = None
        self._connected = False
    
    def is_connected(self) -> bool:
        """Check if connected to AiCalc"""
        return self._connected
    
    def _send_command(self, command: Dict[str, Any]) -> Dict[str, Any]:
        """Send command to AiCalc and receive response."""
        if not self._pipe_handle:
            raise ConnectionError("Not connected to AiCalc")
        
        # Send request
        request_json = json.dumps(command) + "\n"
        win32file.WriteFile(self._pipe_handle, request_json.encode('utf-8'))
        
        # Read response (up to 4KB)
        result, response_data = win32file.ReadFile(self._pipe_handle, 4096, None)
        response_str = response_data.decode('utf-8').strip()
        
        if not response_str:
            raise ConnectionError("No response from server")
        
        return json.loads(response_str)
    
    def get_value(self, cell_ref: str) -> Any:
        """Get value from a cell.
        
        Args:
            cell_ref: Cell reference (e.g., 'A1', 'Sheet1!B2')
            
        Returns:
            Cell value (display value)
        """
        if not self._connected:
            raise ConnectionError("Not connected to AiCalc")
        
        response = self._send_command({
            "command": "get_value",
            "cellRef": cell_ref
        })
        
        if not response.get("success"):
            raise ValueError(response.get("error", "Unknown error"))
        
        data = response.get("data", {})
        return data.get("value")
    
    def set_value(self, cell_ref: str, value: Any) -> bool:
        """Set value of a cell.
        
        Args:
            cell_ref: Cell reference (e.g., 'A1', 'Sheet1!B2')
            value: Value to set
            
        Returns:
            True if successful
        """
        if not self._connected:
            raise ConnectionError("Not connected to AiCalc")
        
        response = self._send_command({
            "command": "set_value",
            "cellRef": cell_ref,
            "value": value
        })
        
        if not response.get("success"):
            raise ValueError(response.get("error", "Unknown error"))
        
        return True
    
    def get_range(self, range_ref: str) -> List[List[Any]]:
        """Get values from a range of cells.
        
        Args:
            range_ref: Range reference (e.g., 'A1:B10', 'Sheet1!A1:C5')
            
        Returns:
            2D list of cell values
        """
        if not self._connected:
            raise ConnectionError("Not connected to AiCalc")
        
        response = self._send_command({
            "command": "get_range",
            "rangeRef": range_ref
        })
        
        if not response.get("success"):
            raise ValueError(response.get("error", "Unknown error"))
        
        return response.get("data", {}).get("values", [])
    
    def run_function(self, function_name: str, *args) -> Any:
        """Execute an AiCalc function.
        
        Args:
            function_name: Name of the function to execute
            *args: Function arguments
            
        Returns:
            Function result
        """
        if not self._connected:
            raise ConnectionError("Not connected to AiCalc")
        
        response = self._send_command({
            "command": "run_function",
            "functionName": function_name,
            "args": list(args)
        })
        
        if not response.get("success"):
            raise ValueError(response.get("error", "Unknown error"))
        
        data = response.get("data", {})
        return data.get("result")
    
    def get_sheets(self) -> List[Dict[str, Any]]:
        """Get list of sheets in the workbook.
        
        Returns:
            List of sheet information dictionaries
        """
        if not self._connected:
            raise ConnectionError("Not connected to AiCalc")
        
        response = self._send_command({"command": "get_sheets"})
        
        if not response.get("success"):
            raise ValueError(response.get("error", "Unknown error"))
        
        return response.get("data", {}).get("sheets", [])
    
    def __enter__(self):
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()

def connect(pipe_name: str = "AiCalc_Bridge") -> AiCalcClient:
    """Connect to AiCalc application.
    
    Args:
        pipe_name: Name of the named pipe (default: AiCalc_Bridge)
        
    Returns:
        Connected AiCalcClient instance
    """
    client = AiCalcClient(pipe_name)
    client.connect()
    return client
