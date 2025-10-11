"""
AiCalc Client - IPC communication with AiCalc application
Uses Named Pipes on Windows for inter-process communication
"""

import json
import struct
import time
from typing import Any, Callable, Dict, List, Optional
from dataclasses import dataclass

try:
    import win32pipe
    import win32file
    import pywintypes
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

from .models import CellAddress, CellValue, CellType


@dataclass
class IPCMessage:
    """IPC message structure"""
    command: str
    params: Dict[str, Any]
    request_id: int
    
    def to_bytes(self) -> bytes:
        """Serialize message to bytes"""
        data = {
            "command": self.command,
            "params": self.params,
            "request_id": self.request_id
        }
        json_str = json.dumps(data)
        json_bytes = json_str.encode('utf-8')
        # Prepend 4-byte length header
        length = struct.pack('I', len(json_bytes))
        return length + json_bytes
    
    @classmethod
    def from_bytes(cls, data: bytes) -> "IPCMessage":
        """Deserialize message from bytes"""
        # Remove 4-byte length header if present
        if len(data) > 4:
            data = data[4:]
        json_str = data.decode('utf-8')
        obj = json.loads(json_str)
        return cls(
            command=obj["command"],
            params=obj["params"],
            request_id=obj["request_id"]
        )


class NamedPipeClient:
    """Named pipe client for Windows IPC"""
    
    def __init__(self, pipe_name: str = "AiCalcPipe"):
        if not HAS_WIN32:
            raise RuntimeError("pywin32 is required for named pipe communication. Install with: pip install pywin32")
        
        self.pipe_name = f"\\\\.\\pipe\\{pipe_name}"
        self.handle = None
        self._request_counter = 0
    
    def connect(self, timeout: int = 5000) -> None:
        """Connect to named pipe server"""
        start_time = time.time()
        while time.time() - start_time < timeout / 1000:
            try:
                self.handle = win32file.CreateFile(
                    self.pipe_name,
                    win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                    0,
                    None,
                    win32file.OPEN_EXISTING,
                    0,
                    None
                )
                # Set pipe to message mode
                win32pipe.SetNamedPipeHandleState(
                    self.handle,
                    win32pipe.PIPE_READMODE_MESSAGE,
                    None,
                    None
                )
                return
            except pywintypes.error as e:
                if e.args[0] == 2:  # ERROR_FILE_NOT_FOUND
                    time.sleep(0.1)
                else:
                    raise
        
        raise TimeoutError(f"Could not connect to AiCalc pipe '{self.pipe_name}' within {timeout}ms")
    
    def send_message(self, message: IPCMessage) -> None:
        """Send message through pipe"""
        if not self.handle:
            raise RuntimeError("Not connected. Call connect() first.")
        
        data = message.to_bytes()
        win32file.WriteFile(self.handle, data)
    
    def receive_message(self) -> IPCMessage:
        """Receive message from pipe"""
        if not self.handle:
            raise RuntimeError("Not connected. Call connect() first.")
        
        # Read 4-byte length header
        _, length_bytes = win32file.ReadFile(self.handle, 4)
        length = struct.unpack('I', length_bytes)[0]
        
        # Read message body
        _, data = win32file.ReadFile(self.handle, length)
        return IPCMessage.from_bytes(data)
    
    def send_and_receive(self, message: IPCMessage) -> IPCMessage:
        """Send message and wait for response"""
        self.send_message(message)
        return self.receive_message()
    
    def close(self) -> None:
        """Close pipe connection"""
        if self.handle:
            win32file.CloseHandle(self.handle)
            self.handle = None
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


class Workbook:
    """
    Workbook client for interacting with AiCalc
    
    Example:
        >>> wb = Workbook()
        >>> wb.connect()
        >>> wb.set_value("A1", "Hello")
        >>> print(wb.get_value("A1"))
        Hello
    """
    
    def __init__(self, pipe_name: str = "AiCalcPipe"):
        self.client = NamedPipeClient(pipe_name)
        self._connected = False
    
    def connect(self, timeout: int = 5000) -> None:
        """Connect to AiCalc application"""
        self.client.connect(timeout)
        self._connected = True
    
    def disconnect(self) -> None:
        """Disconnect from AiCalc"""
        self.client.close()
        self._connected = False
    
    def _ensure_connected(self) -> None:
        """Ensure we're connected before operations"""
        if not self._connected:
            raise RuntimeError("Not connected. Call connect() first.")
    
    def _next_request_id(self) -> int:
        """Get next request ID"""
        self.client._request_counter += 1
        return self.client._request_counter
    
    def get_value(self, cell_ref: str) -> Any:
        """
        Get cell value
        
        Args:
            cell_ref: Cell reference like "A1" or "Sheet1!B2"
            
        Returns:
            Cell value (type depends on cell type)
        """
        self._ensure_connected()
        
        addr = CellAddress.parse(cell_ref)
        message = IPCMessage(
            command="GetValue",
            params={
                "sheet": addr.sheet,
                "row": addr.row,
                "column": addr.column
            },
            request_id=self._next_request_id()
        )
        
        response = self.client.send_and_receive(message)
        return response.params.get("value")
    
    def set_value(self, cell_ref: str, value: Any) -> None:
        """
        Set cell value
        
        Args:
            cell_ref: Cell reference like "A1"
            value: Value to set
        """
        self._ensure_connected()
        
        addr = CellAddress.parse(cell_ref)
        message = IPCMessage(
            command="SetValue",
            params={
                "sheet": addr.sheet,
                "row": addr.row,
                "column": addr.column,
                "value": value
            },
            request_id=self._next_request_id()
        )
        
        response = self.client.send_and_receive(message)
        if response.params.get("status") != "success":
            raise RuntimeError(f"Failed to set value: {response.params.get('error')}")
    
    def get_formula(self, cell_ref: str) -> Optional[str]:
        """Get cell formula"""
        self._ensure_connected()
        
        addr = CellAddress.parse(cell_ref)
        message = IPCMessage(
            command="GetFormula",
            params={
                "sheet": addr.sheet,
                "row": addr.row,
                "column": addr.column
            },
            request_id=self._next_request_id()
        )
        
        response = self.client.send_and_receive(message)
        return response.params.get("formula")
    
    def set_formula(self, cell_ref: str, formula: str) -> None:
        """Set cell formula"""
        self._ensure_connected()
        
        addr = CellAddress.parse(cell_ref)
        message = IPCMessage(
            command="SetFormula",
            params={
                "sheet": addr.sheet,
                "row": addr.row,
                "column": addr.column,
                "formula": formula
            },
            request_id=self._next_request_id()
        )
        
        response = self.client.send_and_receive(message)
        if response.params.get("status") != "success":
            raise RuntimeError(f"Failed to set formula: {response.params.get('error')}")
    
    def get_range(self, range_ref: str) -> List[List[Any]]:
        """
        Get range of cells as 2D array
        
        Args:
            range_ref: Range reference like "A1:C10"
            
        Returns:
            2D array of cell values
        """
        self._ensure_connected()
        
        message = IPCMessage(
            command="GetRange",
            params={"range": range_ref},
            request_id=self._next_request_id()
        )
        
        response = self.client.send_and_receive(message)
        return response.params.get("values", [])
    
    def run_function(self, function_name: str, *args) -> Any:
        """
        Execute AiCalc function
        
        Args:
            function_name: Name of the function (e.g., "SUM", "TEXT_TO_IMAGE")
            *args: Function arguments
            
        Returns:
            Function result
        """
        self._ensure_connected()
        
        message = IPCMessage(
            command="RunFunction",
            params={
                "function": function_name,
                "arguments": list(args)
            },
            request_id=self._next_request_id()
        )
        
        response = self.client.send_and_receive(message)
        if response.params.get("status") == "error":
            raise RuntimeError(f"Function execution failed: {response.params.get('error')}")
        return response.params.get("result")
    
    def evaluate_cell(self, cell_ref: str) -> None:
        """Trigger cell evaluation"""
        self._ensure_connected()
        
        addr = CellAddress.parse(cell_ref)
        message = IPCMessage(
            command="EvaluateCell",
            params={
                "sheet": addr.sheet,
                "row": addr.row,
                "column": addr.column
            },
            request_id=self._next_request_id()
        )
        
        response = self.client.send_and_receive(message)
        if response.params.get("status") != "success":
            raise RuntimeError(f"Evaluation failed: {response.params.get('error')}")
    
    def __enter__(self):
        if not self._connected:
            self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()


def connect(pipe_name: str = "AiCalcPipe", timeout: int = 5000) -> Workbook:
    """
    Connect to running AiCalc instance
    
    Args:
        pipe_name: Named pipe name (default: "AiCalcPipe")
        timeout: Connection timeout in milliseconds
        
    Returns:
        Connected Workbook instance
        
    Example:
        >>> workbook = connect()
        >>> workbook.set_value("A1", "Hello from Python!")
    """
    workbook = Workbook(pipe_name)
    workbook.connect(timeout)
    return workbook
