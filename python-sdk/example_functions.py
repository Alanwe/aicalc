"""
Example AiCalc Python Functions

This file demonstrates how to create custom functions for AiCalc
using the @aicalc_function decorator.
"""

from aicalc_sdk import aicalc_function


@aicalc_function(
    name="PYTHON_DOUBLE",
    category="Math",
    description="Doubles a number",
    examples=["=PYTHON_DOUBLE(A1)", "=PYTHON_DOUBLE(42)"]
)
def double_number(x: float) -> float:
    """Returns the input number multiplied by 2"""
    return float(x) * 2


@aicalc_function(
    name="PYTHON_CONCAT",
    category="Text",
    description="Concatenates text with a separator",
    examples=["=PYTHON_CONCAT(A1, A2, ' - ')"]
)
def concat_text(text1: str, text2: str, separator: str = " ") -> str:
    """Joins two text strings with an optional separator"""
    return f"{text1}{separator}{text2}"


@aicalc_function(
    name="PYTHON_WORD_COUNT",
    category="Text",
    description="Counts words in text",
    examples=["=PYTHON_WORD_COUNT(A1)"]
)
def word_count(text: str) -> int:
    """Returns the number of words in the text"""
    return len(str(text).split())


@aicalc_function(name="PYTHON_TRIPLE", category="Math", description="Triples a number")
def triple_number(x: float) -> float:
    """Returns three times the input number"""
    return float(x) * 3


@aicalc_function(
    name="PYTHON_CELSIUS_TO_F",
    category="Math",
    description="Converts Celsius to Fahrenheit",
    examples=["=PYTHON_CELSIUS_TO_F(0)", "=PYTHON_CELSIUS_TO_F(A1)"]
)
def celsius_to_fahrenheit(celsius: float) -> float:
    """Converts temperature from Celsius to Fahrenheit"""
    return (float(celsius) * 9/5) + 32


@aicalc_function(
    name="PYTHON_REVERSE",
    category="Text",
    description="Reverses text",
    examples=["=PYTHON_REVERSE(A1)", "=PYTHON_REVERSE('Hello')"]
)
def reverse_text(text: str) -> str:
    """Returns the text reversed"""
    return str(text)[::-1]
