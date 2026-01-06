"""Centralized error handling for ExceLLM MCP server.

Provides ToolError exception and standardized error response builders.
"""

from typing import Any, Dict, Optional


class ToolError(Exception):
    """Exception raised for tool-related errors.

    Attributes:
        message: Error message
        code: Optional error code for categorization
        details: Optional additional error details
    """

    def __init__(
        self,
        message: str,
        code: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(message)
        self.message = message
        self.code = code or "TOOL_ERROR"
        self.details = details or {}

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for error response."""
        result = {
            "code": self.code,
            "message": self.message,
        }
        if self.details:
            result["details"] = self.details
        return result


def create_error_response(
    error: Exception,
    tool_name: str = "unknown",
) -> Dict[str, Any]:
    """Create a standardized error response dictionary.

    Args:
        error: The exception that occurred
        tool_name: Name of the tool that failed

    Returns:
        Dictionary with success=False and error details
    """
    if isinstance(error, ToolError):
        return {
            "success": False,
            "error": error.to_dict(),
            "tool": tool_name,
        }

    return {
        "success": False,
        "error": {
            "code": "UNEXPECTED_ERROR",
            "message": str(error),
        },
        "tool": tool_name,
    }


# Common error codes
class ErrorCodes:
    """Standard error codes for consistent error handling."""

    EXCEL_NOT_RUNNING = "EXCEL_NOT_RUNNING"
    NO_WORKBOOK_OPEN = "NO_WORKBOOK_OPEN"
    WORKBOOK_NOT_FOUND = "WORKBOOK_NOT_FOUND"
    SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
    INVALID_REFERENCE = "INVALID_REFERENCE"
    PROTECTED_SHEET = "PROTECTED_SHEET"
    READ_ONLY_WORKBOOK = "READ_ONLY_WORKBOOK"
    VALIDATION_ERROR = "VALIDATION_ERROR"
    COM_ERROR = "COM_ERROR"
    VBA_DISABLED = "VBA_DISABLED"
    WRITE_FAILED = "WRITE_FAILED"
