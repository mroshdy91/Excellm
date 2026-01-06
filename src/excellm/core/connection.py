"""COM connection management for ExceLLM MCP server.

Provides thread-local connection pooling and batch read operations
for improved performance with Excel COM automation.
"""

import functools
import threading
from typing import Any, Callable, Dict, List, Tuple, TypeVar

import pythoncom
import win32com.client as win32

from .errors import ErrorCodes, ToolError
from .utils import number_to_column

# Thread-local storage for COM connections
_thread_local = threading.local()

# Type variable for decorator
F = TypeVar('F', bound=Callable[..., Any])


def _init_com() -> None:
    """Initialize COM for the current thread."""
    if not getattr(_thread_local, '_com_initialized', False):
        pythoncom.CoInitialize()
        _thread_local._com_initialized = True


def _uninit_com() -> None:
    """Uninitialize COM for the current thread.

    Call this to properly release COM resources when done.
    This is especially important for worker threads.
    """
    if getattr(_thread_local, '_com_initialized', False):
        try:
            # Clear cached Excel app reference
            _thread_local.excel_app = None
            pythoncom.CoUninitialize()
        except Exception:
            pass  # COM might already be uninitialized
        _thread_local._com_initialized = False


class COMContext:
    """Context manager for deterministic COM lifecycle management.

    Usage:
        with COMContext():
            app = get_excel_app()
            # ... do COM operations
        # COM is properly cleaned up here (for worker threads)

    Main thread keeps COM alive for performance; worker threads clean up.
    """

    def __init__(self, force_cleanup: bool = False):
        """Initialize context.

        Args:
            force_cleanup: If True, always uninitialize COM on exit.
                          If False (default), only uninit for worker threads.
        """
        self._force_cleanup = force_cleanup
        self._was_initialized = False

    def __enter__(self):
        self._was_initialized = getattr(_thread_local, '_com_initialized', False)
        _init_com()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Only uninit if:
        # 1. force_cleanup is True, OR
        # 2. This is a worker thread (not main) AND we initialized COM
        is_main = threading.current_thread() is threading.main_thread()

        if self._force_cleanup or (not is_main and not self._was_initialized):
            _uninit_com()

        return False  # Don't suppress exceptions


def get_excel_app():
    """Get or create Excel Application COM object for current thread.

    Uses thread-local storage to cache the connection, avoiding
    repeated COM initialization overhead.

    Returns:
        Excel.Application COM object

    Raises:
        ToolError: If Excel is not running
    """
    _init_com()

    # Check if we have a cached connection
    app = getattr(_thread_local, 'excel_app', None)

    if app is not None:
        # Verify connection is still valid
        try:
            _ = app.Workbooks.Count
            return app
        except Exception:
            # Connection is stale, need to reconnect
            _thread_local.excel_app = None

    # Create new connection
    try:
        app = win32.GetActiveObject("Excel.Application")
        _thread_local.excel_app = app
        return app
    except Exception as e:
        raise ToolError(
            f"Excel is not running or not accessible: {e}",
            code=ErrorCodes.EXCEL_NOT_RUNNING
        )


def get_workbook(app, workbook_name: str):
    """Get a workbook by name from the Excel application.

    Args:
        app: Excel Application COM object
        workbook_name: Name of the workbook

    Returns:
        Workbook COM object

    Raises:
        ToolError: If workbook not found
    """
    try:
        return app.Workbooks(workbook_name)
    except Exception:
        raise ToolError(
            f"Workbook '{workbook_name}' not found. Is it open in Excel?",
            code=ErrorCodes.WORKBOOK_NOT_FOUND
        )


def get_worksheet(workbook, sheet_name: str):
    """Get a worksheet by name from a workbook.

    Args:
        workbook: Workbook COM object
        sheet_name: Name of the worksheet

    Returns:
        Worksheet COM object

    Raises:
        ToolError: If worksheet not found
    """
    try:
        return workbook.Worksheets(sheet_name)
    except Exception:
        raise ToolError(
            f"Worksheet '{sheet_name}' not found in workbook.",
            code=ErrorCodes.SHEET_NOT_FOUND
        )


def with_excel_context(func: F) -> F:
    """Decorator that provides Excel app, workbook, and worksheet to a function.

    The decorated function should accept workbook_name and sheet_name as the
    first two arguments. The decorator adds kwargs: app, workbook, worksheet.

    Usage:
        @with_excel_context
        def my_operation(workbook_name, sheet_name, *, app, workbook, worksheet):
            # app, workbook, worksheet are provided automatically
            return worksheet.Range("A1").Value
    """
    @functools.wraps(func)
    def wrapper(workbook_name: str, sheet_name: str, *args, **kwargs):
        _init_com()

        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)
        worksheet = get_worksheet(workbook, sheet_name)

        # Check for protected/read-only
        if worksheet.ProtectContents:
            raise ToolError(
                f"Worksheet '{sheet_name}' is protected.",
                code=ErrorCodes.PROTECTED_SHEET
            )

        return func(
            workbook_name,
            sheet_name,
            *args,
            app=app,
            workbook=workbook,
            worksheet=worksheet,
            **kwargs
        )

    return wrapper  # type: ignore


def batch_read_values(
    worksheet,
    range_str: str,
) -> List[List[Any]]:
    """Read a range of values in a single COM call.

    This is significantly faster than reading cells individually,
    especially for large ranges (O(1) vs O(n) COM calls).

    Args:
        worksheet: Worksheet COM object
        range_str: Excel range like "A1:Z100"

    Returns:
        2D list of values (list of rows, each row is list of values)
    """
    range_obj = worksheet.Range(range_str)
    values = range_obj.Value2

    if values is None:
        return [[]]

    # Single cell returns scalar
    if not isinstance(values, tuple):
        return [[values]]

    # Single row returns 1D tuple, single column returns tuple of 1-tuples
    # Multi-cell range returns tuple of tuples
    result = []
    for row in values:
        if isinstance(row, tuple):
            result.append(list(row))
        else:
            # Single column case - each "row" is a single value
            result.append([row])

    return result


def batch_read_with_positions(
    worksheet,
    positions: List[Tuple[int, int]],
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> Dict[Tuple[int, int], Any]:
    """Read multiple cell positions using a single batch read.

    Reads the bounding rectangle and extracts only the requested positions.
    Much faster than individual cell reads.

    Args:
        worksheet: Worksheet COM object
        positions: List of (row, col) tuples to read (1-based)
        start_row, start_col, end_row, end_col: Bounding rectangle

    Returns:
        Dictionary mapping (row, col) to value
    """
    # Build range address for bounding rectangle
    range_str = f"{number_to_column(start_col)}{start_row}:{number_to_column(end_col)}{end_row}"

    # Batch read the entire rectangle
    all_values = batch_read_values(worksheet, range_str)

    # Extract requested positions
    result = {}
    for row, col in positions:
        # Convert to 0-based indices relative to start
        row_idx = row - start_row
        col_idx = col - start_col

        if 0 <= row_idx < len(all_values):
            row_data = all_values[row_idx]
            if 0 <= col_idx < len(row_data):
                result[(row, col)] = row_data[col_idx]
            else:
                result[(row, col)] = None
        else:
            result[(row, col)] = None

    return result


def ensure_workbook_open() -> None:
    """Verify that at least one workbook is open in Excel.

    Raises:
        ToolError: If no workbook is open
    """
    app = get_excel_app()
    if app.Workbooks.Count == 0:
        raise ToolError(
            "No workbook is open in Excel.",
            code=ErrorCodes.NO_WORKBOOK_OPEN
        )


def get_active_workbook():
    """Get the active workbook in Excel.

    Returns:
        Active Workbook COM object

    Raises:
        ToolError: If no workbook is open
    """
    app = get_excel_app()
    ensure_workbook_open()

    wb = app.ActiveWorkbook
    if not wb:
        raise ToolError(
            "No active workbook found.",
            code=ErrorCodes.NO_WORKBOOK_OPEN
        )

    return wb


def get_active_sheet():
    """Get the active worksheet in Excel.

    Returns:
        Tuple of (Workbook, Worksheet) COM objects

    Raises:
        ToolError: If no workbook/sheet is active
    """
    workbook = get_active_workbook()
    app = get_excel_app()

    try:
        worksheet = app.ActiveSheet
        return workbook, worksheet
    except Exception:
        raise ToolError(
            "No active worksheet found.",
            code=ErrorCodes.SHEET_NOT_FOUND
        )
