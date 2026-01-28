"""
Office Automation Core Module

This module contains the core functionality for Office document automation,
including Word, Excel, PowerPoint processing and WPS integration.
"""

from .word_processor import WordProcessor
from .excel_processor import ExcelProcessor
from .powerpoint_processor import PowerPointProcessor
from .format_converter import FormatConverter
from .wps_integration import WPSIntegration

__all__ = [
    'WordProcessor',
    'ExcelProcessor', 
    'PowerPointProcessor',
    'FormatConverter',
    'WPSIntegration',
]

__version__ = '0.1.0'