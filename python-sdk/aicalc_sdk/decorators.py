"""Decorators for registering Python functions in AiCalc"""

import functools
import inspect
from typing import Callable, Optional

def aicalc_function(name: Optional[str] = None, description: Optional[str] = None):
    """Decorator to register a Python function as an AiCalc function."""
    def decorator(func: Callable) -> Callable:
        func._aicalc_function = True
        func._aicalc_name = name or func.__name__.upper()
        func._aicalc_description = description or func.__doc__ or ""
        
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        return wrapper
    return decorator
