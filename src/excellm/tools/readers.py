"""Read operations for ExceLLM MCP server.

Contains tools for reading cells, ranges, and getting unique values.
"""

import logging
from typing import Any, Dict, List, Optional

from ..core.connection import (
    _init_com,
    get_excel_app,
    get_workbook,
    get_worksheet,
)
from ..core.utils import (
    get_cell_type,
    is_cell_empty,
)
from ..validators import (
    get_excel_error_info,
)

logger = logging.getLogger(__name__)


def read_cell_sync(
    workbook_name: str,
    sheet_name: str,
    cell: str,
) -> Dict[str, Any]:
    """Read a single cell value.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        cell: Cell reference (e.g., "A1", "B5")

    Returns:
        Dictionary with cell value and metadata
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    rng = worksheet.Range(cell)
    value = rng.Value

    # Determine value type
    value_type = get_cell_type(value)

    result = {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "cell": cell,
        "value": str(value) if value is not None else "",
        "type": value_type,
    }

    # Check for formula
    try:
        formula = rng.Formula
        if isinstance(formula, str) and formula.startswith("="):
            result["formula"] = formula
        else:
            result["formula"] = None
    except Exception:
        result["formula"] = None

    # Check for Excel error codes
    error_info = get_excel_error_info(value)
    if error_info:
        result["error_code"] = error_info[0]
        result["error_message"] = error_info[1]

    return result


def read_range_sync(
    workbook_name: str,
    sheet_name: str,
    range_str: Optional[str] = None,
    max_rows: Optional[int] = None,
) -> Dict[str, Any]:
    """Read a range of cells.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range_str: Range reference (e.g., "A1:C5"). Defaults to UsedRange.
        max_rows: Maximum rows to return (prevents token explosion). None = unlimited.

    Returns:
        Dictionary with range data and metadata. If truncated, includes:
        - truncated: True
        - total_available: actual row count before truncation
        - hint: suggestion to narrow range
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    # Get range reference
    if range_str:
        rng = worksheet.Range(range_str)
    else:
        rng = worksheet.UsedRange
        range_str = rng.Address.replace("$", "")

    # Use batch read for performance
    values = rng.Value


    data = []
    error_codes = []
    error_messages = []

    if values is not None:
        # Check if values is a list/tuple (range values) or scalar (single value)
        if isinstance(values, (list, tuple)):
            # Check if values is 2D
            is_2d = len(values) > 0 and isinstance(values[0], (list, tuple))

            if is_2d:
                for row in values:
                    row_data = []
                    row_error_codes = []
                    row_error_msgs = []
                    for cell in row:
                        row_data.append(str(cell) if cell is not None else "")
                        error_info = get_excel_error_info(cell)
                        if error_info:
                            row_error_codes.append(error_info[0])
                            row_error_msgs.append(error_info[1])
                        else:
                            row_error_codes.append(None)
                            row_error_msgs.append(None)
                    data.append(row_data)
                    error_codes.append(row_error_codes)
                    error_messages.append(row_error_msgs)
            else:
                # Single row (1D list/tuple)
                # Note: read_range usually returns 2D tuple ((v,v),), but just in case
                row_data = []
                row_error_codes = []
                row_error_msgs = []
                for cell in values:
                    row_data.append(str(cell) if cell is not None else "")
                    error_info = get_excel_error_info(cell)
                    if error_info:
                        row_error_codes.append(error_info[0])
                        row_error_msgs.append(error_info[1])
                    else:
                        row_error_codes.append(None)
                        row_error_msgs.append(None)
                data.append(row_data)
                error_codes.append(row_error_codes)
                error_messages.append(row_error_msgs)
        else:
            # Scalar value (single cell content like 'Text' or 123)
            # Treat as 1x1 grid
            row_data = [str(values)]
            error_info = get_excel_error_info(values)
            if error_info:
                row_err_code = [error_info[0]]
                row_err_msg = [error_info[1]]
            else:
                row_err_code = [None]
                row_err_msg = [None]

            data.append(row_data)
            error_codes.append(row_err_code)
            error_messages.append(row_err_msg)

    # Calculate dimensions
    rows = len(data)
    cols = len(data[0]) if data else 0

    # Pagination: apply max_rows limit if specified
    truncated = False
    total_available = rows
    if max_rows and rows > max_rows:
        data = data[:max_rows]
        error_codes = error_codes[:max_rows]
        error_messages = error_messages[:max_rows]
        truncated = True
        rows = max_rows

    result = {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "range": range_str,
        "data": data,
        "rows": rows,
        "cols": cols,
        "start_row": rng.Row,
        "start_col": rng.Column,
        "end_row": rng.Row + rows - 1,
        "end_col": rng.Column + cols - 1,
    }

    # Add truncation info if data was limited
    if truncated:
        result["truncated"] = True
        result["total_available"] = total_available
        result["hint"] = f"Returned {max_rows} of {total_available} rows. Narrow range or use offset."

    # Get formulas
    try:
        formulas = rng.Formula
        formula_data = []
        if formulas:
            if isinstance(formulas, (list, tuple)):
                if len(formulas) > 0 and isinstance(formulas[0], (list, tuple)):
                    for row in formulas:
                        row_formulas = []
                        for f in row:
                            if isinstance(f, str) and f.startswith("="):
                                row_formulas.append(f)
                            else:
                                row_formulas.append(None)
                        formula_data.append(row_formulas)
                else:
                    row_formulas = []
                    for f in formulas:
                        if isinstance(f, str) and f.startswith("="):
                            row_formulas.append(f)
                        else:
                            row_formulas.append(None)
                    formula_data.append(row_formulas)
            else:
                if isinstance(formulas, str) and formulas.startswith("="):
                    formula_data.append([formulas])
                else:
                    formula_data.append([None])

        result["formulas"] = formula_data if formula_data else None
    except Exception:
        result["formulas"] = None

    # Only include error fields if there are errors
    has_errors = any(any(cell for cell in row) for row in error_codes)
    if has_errors:
        result["error_codes"] = error_codes
        result["error_messages"] = error_messages

    return result


def get_unique_values_sync(
    workbook_name: str,
    sheet_name: str,
    range_str: Optional[str] = None,
) -> Dict[str, Any]:
    """Get unique values and their frequencies from an Excel range.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range_str: Excel range. Defaults to UsedRange if omitted.

    Returns:
        Dictionary with unique values and counts
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    # Get range
    if range_str:
        rng = worksheet.Range(range_str)
    else:
        rng = worksheet.UsedRange
        range_str = rng.Address.replace("$", "")

    # Read all values
    values = rng.Value

    # Count occurrences
    value_counts = {}

    if values:
        if isinstance(values, (list, tuple)):
            if len(values) > 0 and isinstance(values[0], (list, tuple)):
                for row in values:
                    for cell in row:
                        if not is_cell_empty(cell):
                            val = str(cell)
                            value_counts[val] = value_counts.get(val, 0) + 1
            else:
                for cell in values:
                    if not is_cell_empty(cell):
                        val = str(cell)
                        value_counts[val] = value_counts.get(val, 0) + 1
        else:
            if not is_cell_empty(values):
                val = str(values)
                value_counts[val] = 1

    # Sort by count descending
    sorted_values = sorted(
        [{"value": k, "count": v} for k, v in value_counts.items()],
        key=lambda x: x["count"],
        reverse=True
    )

    return {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "range": range_str,
        "unique_count": len(sorted_values),
        "values": sorted_values
    }


def get_current_selection_sync() -> Dict[str, Any]:
    """Get information about the current selection in Excel.

    Returns:
        Dictionary with selection details
    """
    _init_com()

    app = get_excel_app()

    if app.Workbooks.Count == 0:
        return {
            "success": True,
            "has_selection": False,
            "message": "No workbook is open"
        }

    try:
        selection = app.Selection
        if not selection:
            return {
                "success": True,
                "has_selection": False,
            }

        address = selection.Address.replace("$", "")
        workbook = app.ActiveWorkbook.Name
        sheet = app.ActiveSheet.Name

        rows = selection.Rows.Count
        cols = selection.Columns.Count
        cell_count = selection.Cells.Count

        # Determine selection type
        if rows == 1 and cols == 1:
            selection_type = "single_cell"
        elif rows == 1048576:  # All rows
            selection_type = "entire_column"
        elif cols == 16384:  # All columns
            selection_type = "entire_row"
        else:
            selection_type = "range"

        result = {
            "success": True,
            "has_selection": True,
            "selection_type": selection_type,
            "address": address,
            "address_local": selection.AddressLocal,
            "rows": rows,
            "columns": cols,
            "cell_count": cell_count,
            "workbook": workbook,
            "sheet": sheet,
        }

        # Add value(s) for reasonable selections
        if cell_count == 1:
            result["value"] = selection.Value
            result["values"] = None
        elif cell_count <= 1000:
            values = selection.Value
            if isinstance(values, (list, tuple)):
                if isinstance(values[0], (list, tuple)):
                    result["values"] = [[str(c) if c is not None else "" for c in row] for row in values]
                else:
                    result["values"] = [[str(c) if c is not None else "" for c in values]]
                result["value_rows"] = len(result["values"])
                result["value_cols"] = len(result["values"][0]) if result["values"] else 0
            else:
                result["value"] = values
            result["value"] = None
        else:
            result["value"] = None
            result["values"] = None

        return result


    except Exception as e:
        return {
            "success": True,
            "has_selection": False,
            "message": f"Could not get selection: {str(e)}"
        }


def batch_read_sync(
    workbook_name: str,
    batch: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """Execute multiple read requests reusing the COM connection.

    Args:
        workbook_name: Default workbook name (can be overridden in batch items)
        batch: List of read requests, e.g.:
               [{"sheet": "Sheet1", "range": "A1"}, {"range": "B2"}]

    Returns:
        Dictionary with results list and success status
    """
    _init_com()
    app = get_excel_app()

    # Cache workbook reference to avoid repeated lookups
    cached_wb_name = None
    cached_wb = None

    results = []

    for i, req in enumerate(batch):
        try:
            # Determine target workbook
            wb_name = req.get("workbook", workbook_name)

            # Get workbook object (reuse if same)
            if wb_name != cached_wb_name:
                cached_wb = get_workbook(app, wb_name)
                cached_wb_name = wb_name

            # Determine sheet
            sheet_name = req.get("sheet")
            if not sheet_name:
                # Use active sheet if not specified
                sheet = cached_wb.ActiveSheet
                sheet_name = sheet.Name
            else:
                sheet = get_worksheet(cached_wb, sheet_name)

            # Determine range
            range_str = req.get("range") # None = UsedRange

            # Perform read
            # reuse logic from read_range_sync but we have objects already
            if range_str:
                rng = sheet.Range(range_str)
            else:
                rng = sheet.UsedRange
                range_str = rng.Address.replace("$", "")

            # Read value (fastest method)
            val = rng.Value

            # Process value similar to read_range_sync
            data = []
            if val is not None:
                if isinstance(val, (list, tuple)):
                    # 2D or 1D tuple
                     if len(val) > 0 and isinstance(val[0], (list, tuple)):
                         data = [[str(c) if c is not None else "" for c in r] for r in val]
                     else:
                         data = [[str(c) if c is not None else "" for c in val]]
                else:
                    # Scalar
                    data = [[str(val)]]

            rows = len(data)
            cols = len(data[0]) if data else 0

            results.append({
                "request_index": i,
                "success": True,
                "workbook": wb_name,
                "sheet": sheet_name,
                "range": range_str,
                "data": data,
                "rows": rows,
                "cols": cols
            })

        except Exception as e:
            results.append({
                "request_index": i,
                "success": False,
                "error": str(e)
            })

    return {
        "success": True,
        "results": results,
        "count": len(results)
    }

