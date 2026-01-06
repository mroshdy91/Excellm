"""Search operations for ExceLLM MCP server.

Contains the search/filter tool that uses the FilterEngine.
"""

import logging
from typing import Any, Dict, Optional, Union

from ..core.connection import (
    _init_com,
    get_excel_app,
    get_workbook,
    get_worksheet,
)
from ..core.errors import ToolError
from ..filters import FilterEngine

logger = logging.getLogger(__name__)


def search_sync(
    workbook_name: str,
    filters: Union[str, Dict[str, Any]],
    sheet_name: Optional[str] = None,
    range_str: Optional[str] = None,
    has_header: bool = True,
    all_sheets: bool = False,
    max_rows: Optional[int] = None,
) -> Dict[str, Any]:
    """Search and filter Excel data.

    Args:
        workbook_name: Name of open workbook
        filters: Filter specification (string for simple search, dict for complex)
        sheet_name: Name of worksheet (required if all_sheets is False)
        range_str: Range to search (defaults to UsedRange)
        has_header: If True, first row contains headers
        all_sheets: If True, search all sheets
        max_rows: Maximum rows to return (prevents token explosion)

    Returns:
        Dictionary with filtered data and metadata
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)

    if all_sheets:
        # Search all sheets
        all_results = []
        sheets_searched = 0
        total_rows_filtered = 0

        for i in range(1, workbook.Worksheets.Count + 1):
            ws = workbook.Worksheets(i)
            if ws.Visible != -1:  # Skip hidden sheets
                continue

            try:
                result = _search_sheet(
                    ws, filters, range_str, has_header, workbook_name, max_rows
                )
                if result.get("rows_filtered", 0) > 0:
                    all_results.append({
                        "sheet": ws.Name,
                        "rows_filtered": result["rows_filtered"],
                        "data": result["data"],
                        "cell_locations": result.get("cell_locations", []),
                    })
                    total_rows_filtered += result["rows_filtered"]
                sheets_searched += 1
            except Exception as e:
                logger.warning(f"Error searching sheet {ws.Name}: {e}")

        return {
            "success": True,
            "workbook": workbook_name,
            "all_sheets": True,
            "sheets_searched": sheets_searched,
            "total_rows_filtered": total_rows_filtered,
            "results": all_results,
        }

    else:
        if not sheet_name:
            raise ToolError("sheet_name is required when all_sheets is False")

        worksheet = get_worksheet(workbook, sheet_name)
        result = _search_sheet(
            worksheet, filters, range_str, has_header, workbook_name, max_rows
        )
        result["sheet"] = sheet_name
        return result


def _search_sheet(
    worksheet,
    filters: Union[str, Dict[str, Any]],
    range_str: Optional[str],
    has_header: bool,
    workbook_name: str,
    max_rows: Optional[int] = None,
) -> Dict[str, Any]:
    """Search a single worksheet."""

    # Get range
    if range_str:
        rng = worksheet.Range(range_str)
    else:
        rng = worksheet.UsedRange
        range_str = rng.Address.replace("$", "")

    # Read data
    values = rng.Value
    if not values:
        return {
            "success": True,
            "workbook": workbook_name,
            "range": range_str,
            "data": [],
            "rows_filtered": 0,
            "rows_removed": 0,
            "original_rows": 0,
        }

    # Convert to 2D list
    data = []
    if isinstance(values, (list, tuple)):
        if len(values) > 0 and isinstance(values[0], (list, tuple)):
            for row in values:
                data.append([cell for cell in row])
        else:
            data.append([cell for cell in values])
    else:
        data.append([values])

    original_rows = len(data)

    # Extract headers if present
    headers = None
    if has_header and data:
        headers = [str(h) if h else f"Column{i+1}" for i, h in enumerate(data[0])]
        data_rows = data[1:]
    else:
        data_rows = data

    # Initialize filter engine
    engine = FilterEngine()

    # Apply filters
    filtered_data, row_indices = engine.filter_data(data_rows, filters, headers)

    # Build result with headers
    result_data = []
    if has_header and headers:
        result_data.append(headers)
    result_data.extend(filtered_data)

    # Calculate cell locations (Excel row numbers)
    cell_locations = []
    start_row = rng.Row + (1 if has_header else 0)
    for idx in row_indices:
        excel_row = start_row + idx
        cell_locations.append(f"A{excel_row}")

    rows_filtered = len(filtered_data)
    rows_removed = len(data_rows) - rows_filtered

    # Pagination: apply max_rows limit
    truncated = False
    total_available = rows_filtered
    if max_rows and rows_filtered > max_rows:
        filtered_data = filtered_data[:max_rows]
        cell_locations = cell_locations[:max_rows]
        rows_filtered = max_rows
        truncated = True

    # Get warnings from engine
    warnings = engine.get_warnings() if hasattr(engine, 'get_warnings') else []

    result = {
        "success": True,
        "workbook": workbook_name,
        "range": range_str,
        "data": result_data,
        "rows_filtered": rows_filtered,
        "rows_removed": rows_removed,
        "original_rows": original_rows - (1 if has_header else 0),
        "columns": headers,
        "cell_locations": cell_locations,
    }

    if warnings:
        result["warnings"] = warnings

    # Add truncation info if data was limited
    if truncated:
        result["truncated"] = True
        result["total_available"] = total_available
        result["hint"] = f"Returned {max_rows} of {total_available} matching rows."

    # Human-readable filter description
    if isinstance(filters, str):
        result["filter_applied"] = f"contains '{filters}'"
    elif isinstance(filters, dict):
        result["filter_applied"] = "complex filter"

    return result
