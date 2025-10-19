"""Decorators for registering Python functions in AiCalc"""

import functools
import inspect
from typing import Callable, Optional, List, Dict, Any

def aicalc_function(
    name: Optional[str] = None,
    category: str = "Python",
    description: Optional[str] = None,
    examples: Optional[List[str]] = None
):
    """
    Decorator to register a Python function as an AiCalc function.
    
    Args:
        name: Display name for the function (default: function name in UPPER_CASE)
        category: Function category (default: "Python")
        description: Function description (default: function docstring)
        examples: List of usage examples (default: None)
    
    Example:
        @aicalc_function(
            name="CUSTOM_SUM",
            category="Math",
            description="Custom sum with multiplier",
            examples=["=CUSTOM_SUM(A1, A2, 2)"]
        )
        def custom_sum(a: float, b: float, multiplier: float = 1.0) -> float:
            '''Sum two numbers and multiply by a factor'''
            return (a + b) * multiplier
    """
    def decorator(func: Callable) -> Callable:
        # Extract function signature
        sig = inspect.signature(func)
        
        # Build parameter metadata
        parameters = []
        for param_name, param in sig.parameters.items():
            param_info = {
                "name": param_name,
                "type": param.annotation.__name__ if param.annotation != inspect.Parameter.empty else "any",
                "required": param.default == inspect.Parameter.empty,
                "default": param.default if param.default != inspect.Parameter.empty else None
            }
            parameters.append(param_info)
        
        # Attach metadata to function
        func._aicalc_function = True
        func._aicalc_name = name or func.__name__.upper()
        func._aicalc_category = category
        func._aicalc_description = description or func.__doc__ or ""
        func._aicalc_examples = examples or []
        func._aicalc_parameters = parameters
        func._aicalc_return_type = sig.return_annotation.__name__ if sig.return_annotation != inspect.Parameter.empty else "any"
        
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        
        # Copy metadata to wrapper
        wrapper._aicalc_function = func._aicalc_function
        wrapper._aicalc_name = func._aicalc_name
        wrapper._aicalc_category = func._aicalc_category
        wrapper._aicalc_description = func._aicalc_description
        wrapper._aicalc_examples = func._aicalc_examples
        wrapper._aicalc_parameters = func._aicalc_parameters
        wrapper._aicalc_return_type = func._aicalc_return_type
        
        return wrapper
    return decorator


def get_function_metadata(func: Callable) -> Optional[Dict[str, Any]]:
    """
    Extract AiCalc function metadata from a decorated function.
    
    Returns:
        Dictionary with function metadata, or None if not an AiCalc function
    """
    if not hasattr(func, '_aicalc_function') or not func._aicalc_function:
        return None
    
    return {
        "name": getattr(func, '_aicalc_name', func.__name__.upper()),
        "category": getattr(func, '_aicalc_category', 'Python'),
        "description": getattr(func, '_aicalc_description', ''),
        "examples": getattr(func, '_aicalc_examples', []),
        "parameters": getattr(func, '_aicalc_parameters', []),
        "return_type": getattr(func, '_aicalc_return_type', 'any')
    }
