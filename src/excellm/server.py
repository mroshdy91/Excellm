"""Main MCP server for Excel live operations.

This module provides the FastMCP server that exposes Excel COM
automation tools to LLM clients (Claude, ChatGPT, Cursor, etc.).

Refactored version with modular tool organization.
"""

import asyncio
import logging
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP

# Import from refactored modules
from .core.errors import ErrorCodes, ToolError

# For backward compatibility - keep ExcelSessionManager for tools that still need it
from .excel_session import ExcelSessionManager
from .inspection import explore_sync, inspect_workbook_sync
from .tools import (
    batch_read_sync,
    # Screen Capture (NEW)
    capture_sheet_sync,
    # Range Operations (NEW)
    copy_range_sync,
    # Chart Operations (NEW)
    create_chart_sync,
    create_parallel_sessions_sync,
    # Pivot Table Operations (NEW)
    create_pivot_table_sync,
    # Table Operations (NEW)
    create_table_sync,
    # Session (NEW - stateful processing)
    create_transform_session_sync,
    delete_sync,
    delete_table_sync,
    # VBA Execution (NEW)
    execute_vba_sync,
    find_replace_sync,
    # Formatters
    format_range_sync,
    get_current_selection_sync,
    get_format_sync,
    get_merged_cells_sync,
    get_session_status_sync,
    get_unique_values_sync,
    insert_sync,
    list_tables_sync,
    # Workbook
    list_workbooks_sync,
    # Sheet Management
    manage_sheet_sync,
    merge_cells_sync,
    process_chunk_sync,
    # Readers
    read_range_sync,
    search_sync,
    select_range_sync,
    sort_range_sync,
    unmerge_cells_sync,
    validate_cell_reference_result,
    # Writers
    write_range_sync,
)
from .tools.history import get_recent_changes_sync

# Configure logging (write to stderr, not stdout!)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger(__name__)


# Create FastMCP instance
mcp = FastMCP(
    name="ExceLLM",
    instructions="""
    ExceLLM: Excel Live MCP Server
    ================================

    This server provides tools for real-time Excel automation on Windows.
    Connects to already-open Excel files and allows LLMs to:

    - List open workbooks and their sheets (with hidden status)
    - Read from cells or ranges (auto-detects based on reference)
    - Search and filter data before returning to LLM (reduces token usage)
    - Write to cells or ranges (auto-detects based on reference)
    - Read/write multiple scattered cells/ranges with comma-separated references
    - Add/remove/hide/unhide/copy/move worksheets
    - Insert/delete rows/columns at specific positions
    - Copy ranges between sheets/workbooks
    - Sort data in ranges
    - Find and replace values
    - Format cells/ranges with predefined styles or custom properties

    Prerequisites:
    - Windows OS (required for COM automation)
    - Microsoft Excel must be running
    - Workbooks must be open in Excel

    Key Tools:
    - list_open_workbooks(): List all open workbooks
    - read(): Read cells/ranges (supports batch=[...] for multi-read)
    - write(): Write to cells/ranges
    - search(): Filter data by conditions
    - format(): Apply formatting
    - insert(): Insert rows/columns
    - delete(): Delete rows/columns (NEW)
    - copy_range(): Copy data between locations (NEW)
    - sort_range(): Sort data by columns (NEW)
    - find_replace(): Find and replace values (NEW)
    - explore(): Sheet-level inspection
    - inspect_workbook(): Workbook-level radar
    """,
)

# Global session manager for backward compatibility
_session_manager: Optional[ExcelSessionManager] = None


def get_session_manager() -> ExcelSessionManager:
    """Get or create Excel session manager (for backward compatibility)."""
    global _session_manager
    if _session_manager is None:
        _session_manager = ExcelSessionManager()
    return _session_manager


# ============================================================================
# MCP Tool Implementations - Using Refactored Modules
# ============================================================================


@mcp.tool()
async def list_open_workbooks() -> dict:
    """List all open Excel workbooks with sheet names.

    Returns:
        Dictionary containing:
        - success: True if operation succeeded
        - workbooks: List of workbooks with name and sheets
        - count: Number of open workbooks

    Raises:
        ToolError: If Excel is not running or connection fails

    Example:
        >>> result = await list_open_workbooks()
        >>> print(result)
        {
            "success": True,
            "workbooks": [
                {"name": "data.xlsx", "sheets": [{"name": "Sheet1", "hidden": False}]},
                {"name": "report.xlsx", "sheets": [{"name": "Summary", "hidden": False}]}
            ],
            "count": 2
        }
    """
    try:
        workbooks = await asyncio.to_thread(list_workbooks_sync)
        return {
            "success": True,
            "workbooks": workbooks,
            "count": len(workbooks),
        }
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to list workbooks: {str(e)}") from e


@mcp.tool()
async def select_range(
    workbook_name: str, sheet_name: str, reference: str
) -> dict:
    """Activate a workbook, worksheet, and select a range.

    This moves the focus in the Excel UI to the specified location.

    Args:
        workbook_name: Name of open workbook (e.g., "data.xlsx")
        sheet_name: Name of worksheet (e.g., "Sheet1")
        reference: Cell or range reference (e.g., "A1", "B2:D10")

    Returns:
        Dictionary with success status and message
    """
    try:
        result = await asyncio.to_thread(
            select_range_sync, workbook_name, sheet_name, reference
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to select range: {str(e)}") from e


@mcp.tool()
async def read(
    workbook_name: str,
    sheet_name: str = None,
    reference: str = None,
    batch: List[dict] = None,
    max_rows: int = None,
) -> dict:
    """Read from a cell, range, or multiple ranges in Excel.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet (required if batch is None)
        reference: Cell (A1), range (A1:C5). Defaults to UsedRange.
        batch: Optional list of read requests: [{"sheet": "S1", "range": "A1"}, ...]
               If provided, sheet_name and reference are ignored.
        max_rows: Maximum rows to return (prevents token explosion). None = unlimited.

    Returns:
        Dictionary with cell/range data and metadata
    """
    try:
        if batch:
            return await asyncio.to_thread(batch_read_sync, workbook_name, batch)

        if not sheet_name:
            raise ToolError("sheet_name is required when batch is not used")

        # Use read_range_sync with max_rows support
        result = await asyncio.to_thread(
            read_range_sync, workbook_name, sheet_name, reference, max_rows
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to read: {str(e)}") from e



@mcp.tool()
async def write(
    workbook_name: str,
    sheet_name: str,
    reference: str,
    data: Any,
    force_overwrite: bool = False,
    activate: bool = True,
    dry_run: bool = False,
    strict_alignment: bool = False,
    # NEW SAFETY PARAMETERS:
    max_cells: int = None,
    verify_source: dict = None,
    abort_threshold: float = 0.0,
) -> dict:
    """Write to a cell or range in Excel.

    **Best Practice:** For large datasets (>100 cells), use a subagent with
    a focused context containing only the current chunk's source data.
    This prevents hallucination from context window pressure.
    Consider using create_transform_session for stateful, chunk-based processing.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        reference: Cell (A1), range (A1:C5), or comma-separated references
        data: Value or 2D array to write
        force_overwrite: If True, bypass caution mode
        activate: If True, activate the range after writing
        dry_run: If True, validate without writing
        strict_alignment: If True, require exact dimension match
        max_cells: Maximum cells allowed (rows × cols, default: 100). Prevents LLM hallucination.
        verify_source: Source verification config:
            {"column": "A", "key_index": 0, "match_mode": "contains"}
        abort_threshold: Max allowed mismatch rate (default: 0.0 = 0%)
            If verification fails above this rate, write is aborted.

    Returns:
        Dictionary with operation result
    """
    try:
        # Check if single cell or range
        if isinstance(data, list) and isinstance(data[0], list):
            # Range write - use write_range_sync directly with new safety features
            result = await asyncio.to_thread(
                write_range_sync,
                workbook_name, sheet_name, reference, data,
                force_overwrite=force_overwrite,
                activate=activate,
                dry_run=dry_run,
                strict_alignment=strict_alignment,
                max_cells=max_cells,
                verify_source=verify_source,
                abort_threshold=abort_threshold,
            )

        else:
            # Single cell or simple list - use session manager
            session = get_session_manager()
            result = await session.write(
                workbook_name, sheet_name, reference, data,
                force_overwrite=force_overwrite,
                activate=activate,
                dry_run=dry_run,
                strict_alignment=strict_alignment,
            )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to write: {str(e)}") from e


@mcp.tool()
async def search(
    workbook_name: str,
    filters: Any,
    sheet_name: str = None,
    range: str = None,
    has_header: bool = True,
    all_sheets: bool = False,
    max_rows: int = None,
) -> dict:
    """Search and filter Excel data before returning to LLM.

    Args:
        workbook_name: Name of open workbook
        filters: Filter dict or simple search string
        sheet_name: Name of worksheet (required if all_sheets is False)
        range: Excel range (defaults to UsedRange)
        has_header: Whether first row contains headers
        all_sheets: If True, search all sheets
        max_rows: Maximum rows to return (prevents token explosion). None = unlimited.

    Returns:
        Dictionary with filtered data and metadata
    """
    try:
        result = await asyncio.to_thread(
            search_sync, workbook_name, filters, sheet_name, range, has_header, all_sheets, max_rows
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to search: {str(e)}") from e


@mcp.tool()
async def manage_sheet(
    workbook_name: str,
    sheet_name: str = None,
    sheet_names: list = None,
    action: str = None,
    force_delete: bool = False,
    target_workbook: str = None,
    target_name: str = None,
    position: str = None,
    reference_sheet: str = None,
) -> dict:
    """Manage worksheets (add, remove, hide, unhide, copy, move, rename).

    Args:
        workbook_name: Name of open workbook
        sheet_name: Single sheet name for operations
        sheet_names: Multiple sheet names for bulk operations
        action: Operation (add, remove, hide, unhide, copy, move, rename)
        force_delete: If True, bypass empty check
        target_workbook: Target workbook for copy/move
        target_name: New name for copy/move/rename
        position: "before" or "after" for positioning
        reference_sheet: Sheet to position before/after

    Returns:
        Dictionary with operation result
    """
    try:
        result = await asyncio.to_thread(
            manage_sheet_sync,
            workbook_name, sheet_name, sheet_names, action,
            force_delete, target_workbook, target_name, position, reference_sheet
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to manage sheet: {str(e)}") from e


@mcp.tool()
async def validate_cell_reference(cell: str) -> dict:
    """Validate Excel cell reference format.

    Args:
        cell: Cell reference to validate

    Returns:
        Dictionary with validation result
    """
    return validate_cell_reference_result(cell)


@mcp.tool()
async def get_current_selection() -> dict:
    """Get information about the current selection in Excel.

    Returns:
        Dictionary with selection details (address, type, workbook, sheet, etc.)
    """
    try:
        result = await asyncio.to_thread(get_current_selection_sync)
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to get selection: {str(e)}") from e


@mcp.tool()
async def insert(
    workbook_name: str,
    sheet_name: str,
    insert_type: str,
    position: str,
    count: int = 1,
) -> dict:
    """Insert rows or columns at a specific position.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        insert_type: "row" or "column"
        position: Row number or column letter
        count: Number to insert (default: 1)

    Returns:
        Dictionary with operation result
    """
    try:
        result = await asyncio.to_thread(
            insert_sync, workbook_name, sheet_name, insert_type, position, count
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to insert: {str(e)}") from e


@mcp.tool()
async def delete(
    workbook_name: str,
    sheet_name: str,
    delete_type: str,
    position: str,
    count: int = 1,
) -> dict:
    """Delete rows or columns at a specific position.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        delete_type: "row" or "column"
        position: Row number (e.g., "5", "5:10") or column (e.g., "C", "C:E")
        count: Number to delete (default: 1, ignored if range specified)

    Returns:
        Dictionary with operation result

    Example:
        >>> delete("data.xlsx", "Sheet1", "row", "5", 3)
        {"success": True, "action": "rows_deleted", "count": 3, "at": "5"}
    """
    try:
        result = await asyncio.to_thread(
            delete_sync, workbook_name, sheet_name, delete_type, position, count
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to delete: {str(e)}") from e


@mcp.tool()
async def format(
    workbook_name: str,
    sheet_name: str,
    reference: str,
    style: str = None,
    format: dict = None,
    conditional_format: dict = None,
    activate: bool = True,
) -> dict:
    """Apply formatting to cells/ranges in Excel.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        reference: Cell/range reference
        style: Predefined style (header, currency, percent, warning, success, border, center, wrap)
        format: Custom format properties
        conditional_format: Conditional formatting rules with type:
            - colorScale: {type: "colorScale", min_color: "FF0000", max_color: "00FF00"}
            - dataBar: {type: "dataBar", bar_color: "638EC6"}
            - iconSet: {type: "iconSet", icon_style: "3trafficlights"}
            - cellIs: {type: "cellIs", operator: "greaterThan", value: 100, fill_color: "FFEB9C"}
        activate: If True, activate after formatting

    Returns:
        Dictionary with operation result
    """
    try:
        result = await asyncio.to_thread(
            format_range_sync, workbook_name, sheet_name, reference, style, format, conditional_format, activate
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to format: {str(e)}") from e


@mcp.tool()
async def get_format(
    workbook_name: str,
    sheet_name: str,
    reference: str,
) -> dict:
    """Get formatting properties from cells/ranges.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        reference: Cell/range reference

    Returns:
        Dictionary with formatting properties
    """
    try:
        result = await asyncio.to_thread(
            get_format_sync, workbook_name, sheet_name, reference
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to get format: {str(e)}") from e


@mcp.tool()
async def get_unique_values(
    workbook_name: str,
    sheet_name: str,
    range: str = None,
) -> dict:
    """Get unique values and their frequencies from an Excel range.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range: Excel range (defaults to UsedRange)

    Returns:
        Dictionary with unique values and counts
    """
    try:
        result = await asyncio.to_thread(
            get_unique_values_sync, workbook_name, sheet_name, range
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to get unique values: {str(e)}") from e


# ============================================================================
# NEW TOOLS - Added in Refactoring
# ============================================================================


@mcp.tool()
async def copy_range(
    source_workbook: str,
    source_sheet: str,
    source_range: str,
    target_workbook: str = None,
    target_sheet: str = None,
    target_cell: str = "A1",
    include_formatting: bool = True,
) -> dict:
    """Copy range to another location with optional formatting.

    Args:
        source_workbook: Name of source workbook
        source_sheet: Name of source worksheet
        source_range: Range to copy (e.g., "A1:D10")
        target_workbook: Target workbook (defaults to source)
        target_sheet: Target worksheet (defaults to source)
        target_cell: Top-left cell of paste destination
        include_formatting: If True, copy formatting

    Returns:
        Dictionary with operation result

    Example:
        >>> copy_range("data.xlsx", "Sheet1", "A1:D10", target_sheet="Sheet2", target_cell="E1")
        {"success": True, "cells_copied": 40, ...}
    """
    try:
        result = await asyncio.to_thread(
            copy_range_sync, source_workbook, source_sheet, source_range,
            target_workbook, target_sheet, target_cell, include_formatting
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to copy range: {str(e)}") from e


@mcp.tool()
async def sort_range(
    workbook_name: str,
    sheet_name: str,
    range: str,
    sort_by: list,
    has_header: bool = True,
) -> dict:
    """Sort data in a range by one or more columns.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range: Range to sort (e.g., "A1:D100")
        sort_by: List of sort specs: [{"column": "B", "order": "asc"}, ...]
        has_header: If True, first row is header

    Returns:
        Dictionary with operation result

    Example:
        >>> sort_range("data.xlsx", "Sheet1", "A1:D100",
        ...            [{"column": "B", "order": "asc"}, {"column": "C", "order": "desc"}])
        {"success": True, "rows_sorted": 99, ...}
    """
    try:
        result = await asyncio.to_thread(
            sort_range_sync, workbook_name, sheet_name, range, sort_by, has_header
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to sort range: {str(e)}") from e


@mcp.tool()
async def find_replace(
    workbook_name: str,
    find_value: str,
    replace_value: str,
    sheet_name: str = None,
    match_case: bool = False,
    match_entire_cell: bool = False,
    range: str = None,
    preview_only: bool = False,
) -> dict:
    """Find and replace values in a sheet or workbook.

    Args:
        workbook_name: Name of open workbook
        find_value: Value to find
        replace_value: Value to replace with
        sheet_name: Worksheet (None = all sheets)
        match_case: If True, match case exactly
        match_entire_cell: If True, match entire cell
        range: Specific range (defaults to UsedRange)
        preview_only: If True, count matches without replacing

    Returns:
        Dictionary with operation result

    Example:
        >>> find_replace("data.xlsx", "old", "new", sheet_name="Sheet1")
        {"success": True, "total_matches": 15, "total_replacements": 15, ...}
    """
    try:
        result = await asyncio.to_thread(
            find_replace_sync, workbook_name, sheet_name, find_value, replace_value,
            match_case, match_entire_cell, range, preview_only
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to find/replace: {str(e)}") from e


@mcp.tool()
async def get_recent_changes(limit: int = 10) -> Dict[str, Any]:
    """
    Get the history of recent user actions from the Undo and Redo stacks.

    This tool allows the LLM to "see" what the user has recently done in Excel,
    or what they have undone. It retrieves descriptive action strings (e.g., "Typing 's' in A1").

    Args:
        limit: Max number of items to retrieve per stack (default: 10)

    Returns:
        Dictionary with:
        - success: True
        - undo_history: List of past actions
        - redo_history: List of undone actions available for redo
        - count_undo: Number of undo items
        - count_redo: Number of redo items
    """
    try:
        # Run syncing logic in thread pool to avoid blocking async loop
        result = await asyncio.to_thread(get_recent_changes_sync, limit)

        return {
            "success": True,
            "undo_history": result["undo"],
            "redo_history": result["redo"],
            "count_undo": len(result["undo"]),
            "count_redo": len(result["redo"])
        }
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to get history: {str(e)}") from e


# ============================================================================
# Inspection Tools
# ============================================================================


@mcp.tool()
async def inspect_workbook() -> dict:
    """Fast workbook-level radar across ALL sheets, without reading cell values.

    Provides enough information to select which sheet to explore deeply.

    Returns:
        Dictionary with workbook info, sheet index, and recommendations

    Example:
        >>> result = await inspect_workbook()
        >>> print(result["recommendations"]["nextExploreSheet"])
        "Sheet1"
    """
    try:
        result = await asyncio.to_thread(inspect_workbook_sync)
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to inspect workbook: {str(e)}") from e


@mcp.tool()
async def explore(
    scope: dict,
    mode: str = "quick",
) -> dict:
    """Sheet radar that mimics human first glance: structure/layout without reading full data.

    Args:
        scope: Target sheet - {"sheet": "ACTIVE"} or {"sheet": "SheetName"}
        mode: "quick" (fast sampling) or "deep" (thorough analysis)

    Returns:
        Dictionary with dataFootprint, regions, outliers, layout, flags, readHints, recommendations

    Example:
        >>> result = await explore({"sheet": "ACTIVE"}, mode="quick")
        >>> if result["recommendations"]["shouldRunDeep"]:
        ...     result = await explore({"sheet": "Sheet1"}, mode="deep")
    """
    # Validate scope
    if not scope or "sheet" not in scope:
        raise ToolError("scope must contain 'sheet' key (e.g., {'sheet': 'ACTIVE'})")

    sheet_name = scope.get("sheet", "ACTIVE")

    # Validate mode
    if mode not in ("quick", "deep"):
        raise ToolError(f"Invalid mode '{mode}'. Must be 'quick' or 'deep'")

    try:
        result = await asyncio.to_thread(explore_sync, sheet_name, mode)
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to explore sheet: {str(e)}") from e


# ============================================================================
# Session Tools (NEW - for stateful chunk-based processing)
# ============================================================================


@mcp.tool()
async def create_transform_session(
    workbook_name: str,
    sheet_name: str,
    source_column: str,
    output_columns: str,
    start_row: int = 2,
    end_row: int = None,
    chunk_size: int = 25,
    verify_key_index: int = 0,
) -> dict:
    """Create a stateful transformation session for processing data in chunks.

    **Best Practice:** Use this for datasets > 25 rows. The server tracks progress
    so LLMs don't need to maintain state. Ideal for subagent-based processing
    where each chunk is handled by a focused subagent.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        source_column: Column containing source data (e.g., "A")
        output_columns: Target columns for output (e.g., "B:E")
        start_row: First data row (default: 2, assumes row 1 is header)
        end_row: Last data row (None = auto-detect from UsedRange)
        chunk_size: Rows per chunk (default: 25)
        verify_key_index: Index of output column to verify against source

    Returns:
        session_id, total_rows, first_chunk source data, processing instructions
    """
    try:
        result = await asyncio.to_thread(
            create_transform_session_sync,
            workbook_name, sheet_name, source_column, output_columns,
            start_row, end_row, chunk_size, verify_key_index
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to create session: {str(e)}") from e


@mcp.tool()
async def create_parallel_sessions(
    workbook_name: str,
    sheet_name: str,
    source_column: str,
    output_columns: str,
    start_row: int = 2,
    end_row: int = None,
    num_sessions: int = 2,
    chunk_size: int = 25,
    verify_key_index: int = 0,
) -> dict:
    """Create multiple parallel sessions for faster processing of large datasets.

    Use when:
    - Dataset is large (>500 rows)
    - You want to spawn multiple subagents to process in parallel
    - Each subagent handles a different row range independently

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        source_column: Column containing source data (e.g., "A")
        output_columns: Target columns for output (e.g., "B:E")
        start_row: First data row (default: 2)
        end_row: Last data row (None = auto-detect)
        num_sessions: Number of parallel sessions to create (default: 2)
        chunk_size: Rows per chunk within each session (default: 25)
        verify_key_index: Index of output column to verify against source

    Returns:
        List of sessions with session_id, row range, and first_chunk data
    """
    try:
        result = await asyncio.to_thread(
            create_parallel_sessions_sync,
            workbook_name, sheet_name, source_column, output_columns,
            start_row, end_row, num_sessions, chunk_size, verify_key_index
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to create parallel sessions: {str(e)}") from e


@mcp.tool()
async def process_chunk(
    session_id: str,
    data: list,
) -> dict:
    """Process a chunk of transformed data in an active session.

    Server knows which rows to write based on session state.
    Automatically verifies data against source column (zero tolerance).
    Returns next chunk's source data for continued processing.

    Args:
        session_id: Session ID from create_transform_session
        data: LLM's transformed data for current chunk (2D array)

    Returns:
        rows_written, verification results, remaining rows, next_chunk source data
    """
    try:
        result = await asyncio.to_thread(
            process_chunk_sync, session_id, data
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to process chunk: {str(e)}") from e


@mcp.tool()
async def get_session_status(session_id: str) -> dict:
    """Get the status of an existing transformation session.

    Use to check progress, resume after interruption, or review errors.

    Args:
        session_id: Session ID to check

    Returns:
        processed_rows, remaining_rows, errors, next_chunk source data if active
    """
    try:
        result = await asyncio.to_thread(
            get_session_status_sync, session_id
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to get session status: {str(e)}") from e


# ============================================================================
# NEW TOOLS - VBA, Screen Capture, Table Operations
# ============================================================================


@mcp.tool()
async def execute_vba(
    workbook_name: str,
    vba_code: str,
    module_name: str = None,
    procedure_name: str = "Main",
    sheet_name: str = None,
    timeout: int = 30,
) -> dict:
    """⚙️ ADVANCED OPERATION: Execute custom VBA code in Excel.

    ⚠️ USE WITH CAUTION: VBA code runs directly in Excel and can modify workbook state.
    Only use when standard tools are insufficient.

    🔒 SECURITY: This tool is DISABLED by default.
    Set environment variable EXCELLM_ENABLE_VBA=true to enable.

    Creates a temporary module, executes your VBA code, and cleans up automatically.
    MsgBox statements are removed to prevent blocking popups.

    💡 RECOMMENDED WORKFLOW:
    1. Try standard tools first (read, write, format, etc.)
    2. If operation is too complex, use this as fallback

    Args:
        workbook_name: Name of open workbook
        vba_code: VBA code to execute (can be raw statements or full Sub)
        module_name: Optional custom module name (auto-generated if None)
        procedure_name: Name for the procedure (default: "Main")
        sheet_name: Optional sheet to activate before execution
        timeout: Execution timeout in seconds (default: 30)

    Returns:
        Dictionary with execution result:
        {
            "success": True,
            "message": "VBA executed successfully",
            "procedure_name": "Main",
            "module_name": "TempModule1234"
        }

    Example:
        >>> execute_vba("data.xlsx", '''
        ... Dim ws As Worksheet
        ... Set ws = ActiveSheet
        ... ws.Range("A1").Font.Bold = True
        ... ws.Range("A1").Interior.Color = RGB(255, 255, 0)
        ... ''')
    """
    # Security gate: VBA execution must be explicitly enabled
    from .config import VBA_ENABLED
    if not VBA_ENABLED:
        raise ToolError(
            "VBA execution is DISABLED by default for security. "
            "VBA code runs with full Excel/COM privileges. "
            "To enable, set environment variable: EXCELLM_ENABLE_VBA=true",
            code=ErrorCodes.VBA_DISABLED
        )

    try:
        result = await asyncio.to_thread(
            execute_vba_sync,
            workbook_name,
            vba_code,
            module_name,
            procedure_name,
            sheet_name,
            timeout,
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to execute VBA: {str(e)}") from e


@mcp.tool()
async def capture_sheet(
    workbook_name: str,
    sheet_name: str,
    range_ref: str = None,
    output_format: str = "base64",
    output_path: str = None,
) -> dict:
    """📸 Capture screenshot of Excel sheet or range.

    Captures the visual representation including formatting, colors, borders, etc.
    Useful for validation, documentation, and visual verification of changes.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range_ref: Specific range to capture (None = entire UsedRange)
        output_format: "base64" (default) or "file"
        output_path: File path if output_format is "file"

    Returns:
        Dictionary with:
        {
            "success": True,
            "format": "base64" or "file",
            "image_data": "..." (if base64),
            "file_path": "..." (if file),
            "mime_type": "image/png",
            "range_captured": "A1:H20",
            "sheet_name": "Sheet1"
        }

    Example:
        >>> # Capture entire sheet as base64
        >>> capture_sheet("data.xlsx", "Sheet1")

        >>> # Capture specific range to file
        >>> capture_sheet("data.xlsx", "Sheet1",
        ...               range_ref="A1:H20",
        ...               output_format="file",
        ...               output_path="C:/temp/screenshot.png")
    """
    try:
        result = await asyncio.to_thread(
            capture_sheet_sync,
            workbook_name,
            sheet_name,
            range_ref,
            output_format,
            output_path,
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to capture sheet: {str(e)}") from e


@mcp.tool()
async def create_table(
    workbook_name: str,
    sheet_name: str,
    range_ref: str,
    table_name: str,
    has_headers: bool = True,
    table_style: str = "medium2",
) -> dict:
    """📊 Create an Excel Table (ListObject) from a range.

    Excel Tables provide:
    - Automatic formatting and banding
    - Built-in filtering (dropdown arrows)
    - Structured references in formulas
    - Easy expansion when adding data

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range_ref: Range for the table (e.g., "A1:D100")
        table_name: Name for the table (must be unique in workbook)
        has_headers: Whether first row contains headers (default: True)
        table_style: Style name - shortcuts: "light1", "medium2", "dark1"
                    or full names like "TableStyleMedium9" (default: "medium2")

    Returns:
        Dictionary with:
        {
            "success": True,
            "table_name": "SalesData",
            "range": "A1:D100",
            "rows": 100,
            "columns": 4,
            "has_headers": True,
            "style": "TableStyleMedium2",
            "headers": ["Name", "Date", "Amount", "Status"]
        }

    Example:
        >>> create_table("data.xlsx", "Sheet1", "A1:D100",
        ...              "SalesData", table_style="medium9")
    """
    try:
        result = await asyncio.to_thread(
            create_table_sync,
            workbook_name,
            sheet_name,
            range_ref,
            table_name,
            has_headers,
            table_style,
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to create table: {str(e)}") from e


@mcp.tool()
async def list_tables(
    workbook_name: str,
    sheet_name: str = None,
) -> dict:
    """📋 List all Excel Tables in a workbook or sheet.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Optional - specific sheet (None = all sheets)

    Returns:
        Dictionary with list of tables and their properties
    """
    try:
        result = await asyncio.to_thread(
            list_tables_sync,
            workbook_name,
            sheet_name,
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to list tables: {str(e)}") from e


@mcp.tool()
async def delete_table(
    workbook_name: str,
    sheet_name: str,
    table_name: str,
    keep_data: bool = True,
) -> dict:
    """🗑️ Delete an Excel Table, optionally keeping the data.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        table_name: Name of table to delete
        keep_data: If True, convert to normal range; if False, delete data too

    Returns:
        Dictionary with operation result
    """
    try:
        result = await asyncio.to_thread(
            delete_table_sync,
            workbook_name,
            sheet_name,
            table_name,
            keep_data,
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to delete table: {str(e)}") from e


# ============================================================================
# NEW TOOLS - Charts, Pivot Tables, Cell Merging, Formula Validation
# ============================================================================


@mcp.tool()
async def create_chart(
    workbook_name: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis_title: str = "",
    y_axis_title: str = "",
    width: float = 400,
    height: float = 300,
    style: dict = None,
) -> dict:
    """📊 Create a chart in an Excel worksheet.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        data_range: Source data range (e.g., "A1:D10")
        chart_type: Type of chart (line, bar, pie, scatter, area)
        target_cell: Cell where chart will be placed
        title: Chart title
        x_axis_title: X-axis label
        y_axis_title: Y-axis label
        width: Chart width in points
        height: Chart height in points
        style: Optional style configuration

    Returns:
        Dictionary with chart details
    """
    try:
        result = await asyncio.to_thread(
            create_chart_sync,
            workbook_name, sheet_name, data_range, chart_type,
            target_cell, title, x_axis_title, y_axis_title, width, height, style
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to create chart: {str(e)}") from e


@mcp.tool()
async def create_pivot_table(
    workbook_name: str,
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    columns: List[str] = None,
    agg_func: str = "sum",
    target_sheet: str = None,
    target_cell: str = "A1",
    table_name: str = None,
) -> dict:
    """📊 Create a pivot table from source data.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet with source data
        data_range: Source data range (e.g., "A1:D100")
        rows: Fields for row labels (column headers)
        values: Fields for values to aggregate
        columns: Optional fields for column labels
        agg_func: Aggregation function (sum, count, average, max, min)
        target_sheet: Sheet for pivot table (default: creates new sheet)
        target_cell: Cell for pivot table location
        table_name: Optional name for the pivot table

    Returns:
        Dictionary with pivot table details
    """
    try:
        result = await asyncio.to_thread(
            create_pivot_table_sync,
            workbook_name, sheet_name, data_range, rows, values,
            columns, agg_func, target_sheet, target_cell, table_name
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to create pivot table: {str(e)}") from e


@mcp.tool()
async def merge_cells(
    workbook_name: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
) -> dict:
    """Merge a range of cells.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        start_cell: Starting cell (e.g., "A1")
        end_cell: Ending cell (e.g., "D1")

    Returns:
        Dictionary with operation result
    """
    try:
        result = await asyncio.to_thread(
            merge_cells_sync, workbook_name, sheet_name, start_cell, end_cell
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to merge cells: {str(e)}") from e


@mcp.tool()
async def unmerge_cells(
    workbook_name: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
) -> dict:
    """Unmerge a previously merged range of cells.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        start_cell: Starting cell (e.g., "A1")
        end_cell: Ending cell (e.g., "D1")

    Returns:
        Dictionary with operation result
    """
    try:
        result = await asyncio.to_thread(
            unmerge_cells_sync, workbook_name, sheet_name, start_cell, end_cell
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to unmerge cells: {str(e)}") from e


@mcp.tool()
async def get_merged_cells(
    workbook_name: str,
    sheet_name: str,
) -> dict:
    """Get all merged cell ranges in a worksheet.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet

    Returns:
        Dictionary with list of merged ranges
    """
    try:
        result = await asyncio.to_thread(
            get_merged_cells_sync, workbook_name, sheet_name
        )
        return result
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        raise ToolError(f"Failed to get merged cells: {str(e)}") from e


@mcp.tool()
async def validate_formula(
    formula: str,
) -> dict:
    """Validate Excel formula syntax without applying it.

    Checks for balanced parentheses, valid function names, and common errors.

    Args:
        formula: Formula string to validate (e.g., "=SUM(A1:A10)")

    Returns:
        Dictionary with:
        {
            "valid": bool,
            "error": str or None,
            "warnings": list of str,
            "functions_used": list of str
        }
    """
    from .validators import validate_formula_sync
    return validate_formula_sync(formula)


# ============================================================================
# Server Entry Point
# ============================================================================


def create_server() -> FastMCP:
    """Create and return the MCP server instance.

    Returns:
        FastMCP server instance
    """
    return mcp


if __name__ == "__main__":
    mcp.run(transport="stdio")
