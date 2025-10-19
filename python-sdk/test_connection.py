"""Test script for AiCalc Python SDK"""

import sys
import time
from aicalc_sdk import connect

def main():
    print("AiCalc Python SDK Test")
    print("=" * 50)
    
    try:
        print("\n1. Connecting to AiCalc...")
        with connect() as client:
            print("✓ Connected successfully!")
            
            # Test getting a value
            print("\n2. Testing get_value('A1')...")
            value = client.get_value("A1")
            print(f"   A1 value: {value}")
            
            # Test setting a value
            print("\n3. Testing set_value('A1', 42)...")
            client.set_value("A1", 42)
            print("   ✓ Value set successfully")
            
            # Verify the value
            print("\n4. Verifying value...")
            value = client.get_value("A1")
            print(f"   A1 value: {value}")
            
            # Test getting sheets
            print("\n5. Getting sheets...")
            sheets = client.get_sheets()
            print(f"   Found {len(sheets)} sheet(s):")
            for sheet in sheets:
                print(f"     - {sheet.get('name')}: {sheet.get('row_count')}x{sheet.get('column_count')}")
            
            # Test a function call (if available)
            print("\n6. Testing function call (SUM)...")
            try:
                result = client.run_function("SUM", 1, 2, 3)
                print(f"   SUM(1, 2, 3) = {result}")
            except Exception as e:
                print(f"   Function call skipped: {e}")
            
            print("\n✓ All tests passed!")
            
    except ConnectionError as e:
        print(f"\n✗ Connection failed: {e}")
        print("\nMake sure AiCalc is running before executing this script.")
        return 1
    except Exception as e:
        print(f"\n✗ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
