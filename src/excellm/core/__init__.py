"""Core module for ExceLLM MCP server.

Provides shared utilities, connection management, and error handling.
"""

from .connection import get_excel_app, with_excel_context, batch_read_values
from .utils import (
    normalize_address,
    column_to_number,
    number_to_column,
    parse_range_bounds,
    build_range_address,
    is_cell_empty,
    safe_int,
    safe_bool,
    compute_density,
)
from .errors import ToolError, create_error_response

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
