"""Simple diagnostic to check if Named Pipe exists"""

import win32pipe
import pywintypes
import time

pipe_name = r"\\.\pipe\AiCalc_Bridge"

print(f"Checking for pipe: {pipe_name}")
print("=" * 60)

for i in range(10):
    try:
        result = win32pipe.WaitNamedPipe(pipe_name, 100)
        if result:
            print(f"✓ Pipe found! (attempt {i+1})")
            break
        else:
            print(f"✗ Pipe not available (attempt {i+1})")
    except pywintypes.error as e:
        print(f"✗ Error checking pipe (attempt {i+1}): {e}")
    
    time.sleep(0.5)
else:
    print("\n❌ Pipe not found after 10 attempts (5 seconds)")
    print("\nPossible issues:")
    print("1. AiCalc is not running")
    print("2. Python bridge service failed to start")
    print("3. Pipe name mismatch")
    print("\nCheck Debug output in Visual Studio for Python bridge startup messages.")
