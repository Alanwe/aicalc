#!/usr/bin/env python3
"""
AiCalc Python SDK
Provides Python API to interact with AiCalc spreadsheet application
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="aicalc-sdk",
    version="0.1.0",
    author="AiCalc Team",
    author_email="support@aicalc.app",
    description="Python SDK for AiCalc - AI-Native Spreadsheet",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Alanwe/aicalc",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pywin32>=304",  # For named pipe communication on Windows
    ],
    extras_require={
        "dev": [
            "pytest>=7.0",
            "pytest-asyncio>=0.20",
            "black>=22.0",
            "mypy>=0.950",
        ],
    },
)
