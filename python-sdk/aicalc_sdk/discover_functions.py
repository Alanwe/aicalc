"""
AiCalc Function Discovery Script

This script is called by AiCalc to discover functions decorated with @aicalc_function
in Python files. It outputs JSON metadata that AiCalc can parse.

Usage:
    python discover_functions.py <python_file_path>

Output (JSON):
    {
        "success": true,
        "functions": [
            {
                "name": "CUSTOM_SUM",
                "category": "Math",
                "description": "Custom sum function",
                "file_path": "path/to/file.py",
                "function_name": "custom_sum",
                "parameters": [
                    {"name": "a", "type": "float", "required": true, "default": null},
                    {"name": "b", "type": "float", "required": true, "default": null}
                ],
                "return_type": "float",
                "examples": ["=CUSTOM_SUM(A1, A2)"]
            }
        ],
        "error": null
    }
"""

import sys
import json
import importlib.util
import os
from pathlib import Path

def discover_functions(file_path: str):
    """
    Discover all @aicalc_function decorated functions in a Python file.
    
    Args:
        file_path: Path to Python file to scan
        
    Returns:
        Dictionary with discovered functions and metadata
    """
    try:
        # Validate file exists
        if not os.path.exists(file_path):
            return {
                "success": False,
                "functions": [],
                "error": f"File not found: {file_path}"
            }
        
        # Load module from file
        module_name = Path(file_path).stem
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        if spec is None or spec.loader is None:
            return {
                "success": False,
                "functions": [],
                "error": f"Could not load module from {file_path}"
            }
        
        module = importlib.util.module_from_spec(spec)
        
        # Add parent directory to sys.path so relative imports work
        parent_dir = str(Path(file_path).parent)
        if parent_dir not in sys.path:
            sys.path.insert(0, parent_dir)
        
        # Execute module
        spec.loader.exec_module(module)
        
        # Discover decorated functions
        functions = []
        for name in dir(module):
            obj = getattr(module, name)
            
            # Check if it's a decorated function
            if (callable(obj) and 
                hasattr(obj, '_aicalc_function') and 
                obj._aicalc_function):
                
                func_metadata = {
                    "name": getattr(obj, '_aicalc_name', name.upper()),
                    "category": getattr(obj, '_aicalc_category', 'Python'),
                    "description": getattr(obj, '_aicalc_description', ''),
                    "file_path": file_path,
                    "function_name": name,
                    "parameters": getattr(obj, '_aicalc_parameters', []),
                    "return_type": getattr(obj, '_aicalc_return_type', 'any'),
                    "examples": getattr(obj, '_aicalc_examples', [])
                }
                functions.append(func_metadata)
        
        return {
            "success": True,
            "functions": functions,
            "error": None
        }
        
    except Exception as e:
        return {
            "success": False,
            "functions": [],
            "error": str(e)
        }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({
            "success": False,
            "functions": [],
            "error": "Usage: python discover_functions.py <python_file_path>"
        }))
        sys.exit(1)
    
    file_path = sys.argv[1]
    result = discover_functions(file_path)
    print(json.dumps(result, indent=2))
