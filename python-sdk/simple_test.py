"""Minimal pipe test"""
import win32file
import win32pipe
import json

pipe_name = r"\\.\pipe\AiCalc_Bridge"

print(f"Connecting to {pipe_name}...")
try:
    handle = win32file.CreateFile(
        pipe_name,
        win32file.GENERIC_READ | win32file.GENERIC_WRITE,
        0, None,
        win32file.OPEN_EXISTING,
        0, None
    )
    print("✓ Connected!")
    
    # Send ping
    request = json.dumps({"command": "ping"}) + "\n"
    print(f"Sending: {request.strip()}")
    win32file.WriteFile(handle, request.encode('utf-8'))
    print("✓ Request sent")
    
    # Read response with timeout
    print("Reading response...")
    import msvcrt
    import time
    
    # Try overlapped I/O with timeout
    try:
        result, data = win32file.ReadFile(handle, 4096, None)
        print(f"✓ Response: {data.decode('utf-8')}")
    except Exception as e:
        print(f"✗ Read error: {e}")
    
    win32file.CloseHandle(handle)
    print("✓ Disconnected")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()
