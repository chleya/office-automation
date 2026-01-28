"""
Error handling for Office Automation

This module defines custom exceptions and error handling utilities
for the Office Automation skill.
"""

import traceback
from typing import Optional, Dict, Any
from dataclasses import dataclass


@dataclass
class ErrorContext:
    """Context information for error reporting."""
    operation: str
    file_path: Optional[str] = None
    format: Optional[str] = None
    template: Optional[str] = None
    additional_info: Optional[Dict[str, Any]] = None


class OfficeAutomationError(Exception):
    """Base exception class for all Office Automation errors."""
    
    def __init__(self, message: str, context: Optional[ErrorContext] = None):
        self.message = message
        self.context = context
        super().__init__(self._format_message())
    
    def _format_message(self) -> str:
        """Format error message with context."""
        msg = f"Office Automation Error: {self.message}"
        if self.context:
            msg += f"\nContext: {self.context.operation}"
            if self.context.file_path:
                msg += f" | File: {self.context.file_path}"
            if self.context.format:
                msg += f" | Format: {self.context.format}"
            if self.context.template:
                msg += f" | Template: {self.context.template}"
        return msg


class DocumentCreationError(OfficeAutomationError):
    """Raised when document creation fails."""
    pass


class DocumentReadError(OfficeAutomationError):
    """Raised when reading a document fails."""
    pass


class DocumentSaveError(OfficeAutomationError):
    """Raised when saving a document fails."""
    pass


class FormatConversionError(OfficeAutomationError):
    """Raised when format conversion fails."""
    pass


class WPSNotAvailableError(OfficeAutomationError):
    """Raised when WPS Office is not available but required."""
    pass


class TemplateNotFoundError(OfficeAutomationError):
    """Raised when a template file is not found."""
    pass


class ConfigurationError(OfficeAutomationError):
    """Raised when there's a configuration error."""
    pass


class ValidationError(OfficeAutomationError):
    """Raised when input validation fails."""
    pass


def handle_office_error(
    error: Exception,
    context: Optional[ErrorContext] = None,
    raise_exception: bool = True
) -> Optional[str]:
    """
    Handle Office Automation errors with proper formatting.
    
    Args:
        error: The exception that was raised
        context: Additional context information
        raise_exception: Whether to re-raise the exception
        
    Returns:
        Formatted error message if not re-raising, None otherwise
    """
    # Convert generic exceptions to our custom exceptions
    if isinstance(error, OfficeAutomationError):
        formatted_error = error
    else:
        # Wrap generic exceptions
        error_message = str(error)
        if "permission" in error_message.lower():
            formatted_error = DocumentSaveError(
                f"Permission denied: {error_message}",
                context
            )
        elif "not found" in error_message.lower():
            formatted_error = TemplateNotFoundError(
                f"File not found: {error_message}",
                context
            )
        elif "format" in error_message.lower():
            formatted_error = FormatConversionError(
                f"Format error: {error_message}",
                context
            )
        else:
            formatted_error = OfficeAutomationError(
                f"Unexpected error: {error_message}",
                context
            )
    
    # Log the error (in a real implementation, this would use a proper logger)
    error_details = {
        "error_type": type(formatted_error).__name__,
        "error_message": str(formatted_error),
        "context": context.__dict__ if context else None,
        "traceback": traceback.format_exc(),
    }
    
    print(f"[ERROR] Office Automation Error: {error_details}")
    
    if raise_exception:
        raise formatted_error
    else:
        return str(formatted_error)


def create_error_context(
    operation: str,
    file_path: Optional[str] = None,
    format: Optional[str] = None,
    template: Optional[str] = None,
    **additional_info
) -> ErrorContext:
    """
    Create an ErrorContext object with the given parameters.
    
    Args:
        operation: The operation being performed
        file_path: Path to the file being processed
        format: File format being used
        template: Template being used
        **additional_info: Additional context information
        
    Returns:
        ErrorContext object
    """
    return ErrorContext(
        operation=operation,
        file_path=file_path,
        format=format,
        template=template,
        additional_info=additional_info if additional_info else None,
    )


def validate_file_path(file_path: str, operation: str) -> None:
    """
    Validate a file path for Office operations.
    
    Args:
        file_path: Path to validate
        operation: Operation being performed
        
    Raises:
        ValidationError: If the file path is invalid
    """
    import os
    
    if not file_path:
        raise ValidationError(
            "File path cannot be empty",
            create_error_context(operation, file_path=file_path)
        )
    
    # Check for invalid characters (basic check)
    invalid_chars = ['<', '>', ':', '"', '|', '?', '*']
    for char in invalid_chars:
        if char in file_path:
            raise ValidationError(
                f"File path contains invalid character: {char}",
                create_error_context(operation, file_path=file_path)
            )
    
    # For save operations, check if directory exists
    if operation in ["save", "export", "convert"]:
        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory):
            try:
                os.makedirs(directory, exist_ok=True)
            except Exception as e:
                raise ValidationError(
                    f"Cannot create directory: {e}",
                    create_error_context(operation, file_path=file_path)
                )


def validate_format(format: str, allowed_formats: list, operation: str) -> None:
    """
    Validate a file format.
    
    Args:
        format: Format to validate
        allowed_formats: List of allowed formats
        operation: Operation being performed
        
    Raises:
        ValidationError: If the format is invalid
    """
    if format not in allowed_formats:
        raise ValidationError(
            f"Invalid format: {format}. Allowed formats: {allowed_formats}",
            create_error_context(operation, format=format)
        )