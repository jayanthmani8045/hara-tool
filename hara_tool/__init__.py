"""
HARA Automation Tool
Automated ASIL determination from HARA Excel sheets
"""

__version__ = "1.0.0"
__author__ = "Automotive Safety Systems"
__description__ = "Professional HARA Automation Tool with Fuzzy Matching"

from .processor import ExcelProcessor
from .matcher import FuzzyMatcher
from .constants import ASIL_DETERMINATION_TABLE

__all__ = [
    'ExcelProcessor',
    'FuzzyMatcher',
    'ASIL_DETERMINATION_TABLE'
]