"""Core module for ExceLLM MCP server.

Provides shared utilities, connection management, and error handling.
"""

from .connection import batch_read_values, get_excel_app, with_excel_context
from .errors import ToolError, create_error_response
from .utils import (
    build_range_address,
    column_to_number,
    compute_density,
    is_cell_empty,
    normalize_address,
    number_to_column,
    parse_range_bounds,
    safe_bool,
    safe_int,
)

__all__ = [
    # Connection
    "get_excel_app",
    "with_excel_context",
    "batch_read_values",
    # Utils
    "normalize_address",
    "column_to_number",
    "number_to_column",
    "parse_range_bounds",
    "build_range_address",
    "is_cell_empty",
    "safe_int",
    "safe_bool",
    "compute_density",
    # Errors
    "ToolError",
    "create_error_response",
]
