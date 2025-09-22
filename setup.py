"""
Setup script for HARA Automation Tool
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding="utf-8")

setup(
    name="hara-automation-tool",
    version="1.0.0",
    author="Automotive Safety Systems",
    description="Professional HARA Automation Tool with Fuzzy Matching",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/your-org/hara-tool",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Scientific/Engineering",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pandas>=2.0.0",
        "openpyxl>=3.1.0",
        "PySide6>=6.5.0",
        "fuzzywuzzy>=0.18.0",
        "python-Levenshtein>=0.20.0",
    ],
    entry_points={
        "console_scripts": [
            "hara-tool=hara_tool.main:main",
        ],
    },
    include_package_data=True,
    package_data={
        "hara_tool": ["resources/*.ico"],
    },
)