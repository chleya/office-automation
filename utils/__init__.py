"""
Office Automation Utilities Module

This module contains utility functions and helper classes for
Office document automation.
"""

from .error_handler import (
    OfficeAutomationError,
    DocumentCreationError,
    FormatConversionError,
    WPSNotAvailableError,
    TemplateNotFoundError,
    handle_office_error,
)

from .templates import TemplateManager
from .batch_processor import BatchProcessor

__all__ = [
    'OfficeAutomationError',
    'DocumentCreationError',
    'FormatConversionError', 
    'WPSNotAvailableError',
    'TemplateNotFoundError',
    'handle_office_error',
    'TemplateManager',
    'BatchProcessor',
]

__version__ = '0.1.0'