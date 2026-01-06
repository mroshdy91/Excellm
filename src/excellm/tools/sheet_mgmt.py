"""Sheet management operations for ExceLLM MCP server.

Contains tools for managing sheets, inserting and deleting rows/columns.
"""

import logging
from typing import Any, Dict, List, Optional

from ..core.connection import (
    _init_com,
    get_excel_app,
    get_workbook,
    get_worksheet,
)
from ..core.errors import ToolError
from ..core.utils import (
    column_to_number,
    number_to_column,
)

logger = logging.getLogger(__name__)


def insert_sync(
    workbook_name: str,
    sheet_name: str,
    insert_type: str,
    position: str,
    count: int = 1,
) -> Dict[str, Any]:
    """Insert rows or columns at a specific position.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        insert_type: "row" or "column"
        position: Row number or column letter/number
        count: Number of rows/columns to insert

    Returns:
        Dictionary with operation result
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    if insert_type.lower() == "row":
        # Parse position (could be "5" or "5:10")
        if ":" in position:
            start_row, end_row = position.split(":")
            start_row = int(start_row)
            end_row = int(end_row)
            count = end_row - start_row + 1
        else:
            start_row = int(position)

        # Insert rows
        for _ in range(count):
            worksheet.Rows(start_row).Insert()

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "action": "rows_inserted",
            "count": count,
            "at": position,
            "message": f"Inserted {count} row(s) at row {position}"
        }

    elif insert_type.lower() == "column":
        # Parse position (could be "C" or "3" or "C:E")
        if ":" in position:
            start_col, end_col = position.split(":")
            if start_col.isdigit():
                start_col = number_to_column(int(start_col))
            if end_col.isdigit():
                end_col = number_to_column(int(end_col))
            count = column_to_number(end_col) - column_to_number(start_col) + 1
        else:
            if position.isdigit():
                start_col = number_to_column(int(position))
            else:
                start_col = position.upper()

        # Insert columns
        for _ in range(count):
            worksheet.Columns(start_col).Insert()

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "action": "columns_inserted",
            "count": count,
            "at": start_col,
            "message": f"Inserted {count} column(s) at column {start_col}"
        }

    else:
        raise ToolError(f"Invalid insert_type '{insert_type}'. Must be 'row' or 'column'.")


def delete_sync(
    workbook_name: str,
    sheet_name: str,
    delete_type: str,
    position: str,
    count: int = 1,
) -> Dict[str, Any]:
    """Delete rows or columns at a specific position.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        delete_type: "row" or "column"
        position: Row number or column letter/number (e.g., "5", "5:10", "C", "C:E")
        count: Number of rows/columns to delete (ignored if range specified)

    Returns:
        Dictionary with operation result
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    if delete_type.lower() == "row":
        # Parse position
        if ":" in position:
            start_row, end_row = position.split(":")
            start_row = int(start_row)
            end_row = int(end_row)
            count = end_row - start_row + 1
            # Delete range of rows
            worksheet.Rows(f"{start_row}:{end_row}").Delete()
        else:
            start_row = int(position)
            # Delete count rows starting at position
            if count > 1:
                worksheet.Rows(f"{start_row}:{start_row + count - 1}").Delete()
            else:
                worksheet.Rows(start_row).Delete()

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "action": "rows_deleted",
            "count": count,
            "at": position,
            "message": f"Deleted {count} row(s) at row {position}"
        }

    elif delete_type.lower() == "column":
        # Parse position
        if ":" in position:
            start_col, end_col = position.split(":")
            if start_col.isdigit():
                start_col = number_to_column(int(start_col))
            if end_col.isdigit():
                end_col = number_to_column(int(end_col))
            count = column_to_number(end_col) - column_to_number(start_col) + 1
            # Delete range of columns
            worksheet.Columns(f"{start_col}:{end_col}").Delete()
        else:
            if position.isdigit():
                start_col = number_to_column(int(position))
            else:
                start_col = position.upper()
            # Delete count columns
            if count > 1:
                end_col = number_to_column(column_to_number(start_col) + count - 1)
                worksheet.Columns(f"{start_col}:{end_col}").Delete()
            else:
                worksheet.Columns(start_col).Delete()

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "action": "columns_deleted",
            "count": count,
            "at": start_col if 'start_col' in dir() else position,
            "message": f"Deleted {count} column(s) at column {position}"
        }

    else:
        raise ToolError(f"Invalid delete_type '{delete_type}'. Must be 'row' or 'column'.")


def manage_sheet_sync(
    workbook_name: str,
    sheet_name: Optional[str] = None,
    sheet_names: Optional[List[str]] = None,
    action: Optional[str] = None,
    force_delete: bool = False,
    target_workbook: Optional[str] = None,
    target_name: Optional[str] = None,
    position: Optional[str] = None,
    reference_sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """Manage worksheets in an open Excel workbook.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Single sheet name for add/remove/copy/move/single operation
        sheet_names: Multiple sheet names for bulk hide/unhide operations
        action: Operation to perform (add, remove, hide, unhide, copy, move, rename)
        force_delete: If True, bypass empty check when removing
        target_workbook: For copy/move, target workbook name
        target_name: For copy/move/rename, new name for the sheet
        position: For copy/move, "before" or "after" for positioning
        reference_sheet: For copy/move, sheet to position before/after

    Returns:
        Dictionary with operation result
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)

    if action == "add":
        if not sheet_name:
            raise ToolError("sheet_name is required for 'add' action")

        new_sheet = workbook.Worksheets.Add()
        new_sheet.Name = sheet_name

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "added",
            "sheet": sheet_name,
            "message": f"Added new sheet '{sheet_name}'"
        }

    elif action == "remove":
        if not sheet_name:
            raise ToolError("sheet_name is required for 'remove' action")

        worksheet = get_worksheet(workbook, sheet_name)

        # Check if sheet is empty (unless force_delete)
        if not force_delete:
            used_range = worksheet.UsedRange
            cells_count = used_range.Cells.Count
            if cells_count > 1:
                raise ToolError(
                    f"Sheet '{sheet_name}' is not empty ({cells_count} cells in use). "
                    "Use force_delete=True to delete anyway."
                )

        # Disable alerts to prevent confirmation dialog
        app.DisplayAlerts = False
        try:
            worksheet.Delete()
        finally:
            app.DisplayAlerts = True

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "removed",
            "sheet": sheet_name,
            "message": f"Removed sheet '{sheet_name}'"
        }

    elif action == "hide":
        sheets_to_hide = sheet_names if sheet_names else [sheet_name] if sheet_name else []
        if not sheets_to_hide:
            raise ToolError("sheet_name or sheet_names is required for 'hide' action")

        results = []
        for name in sheets_to_hide:
            worksheet = get_worksheet(workbook, name)
            worksheet.Visible = 0  # xlSheetHidden
            results.append({"sheet": name, "action": "hidden"})

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "hidden",
            "results": results,
            "count": len(results)
        }

    elif action == "unhide":
        sheets_to_unhide = sheet_names if sheet_names else [sheet_name] if sheet_name else []
        if not sheets_to_unhide:
            raise ToolError("sheet_name or sheet_names is required for 'unhide' action")

        results = []
        for name in sheets_to_unhide:
            worksheet = get_worksheet(workbook, name)
            worksheet.Visible = -1  # xlSheetVisible
            results.append({"sheet": name, "action": "unhidden"})

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "unhidden",
            "results": results,
            "count": len(results)
        }

    elif action == "copy":
        if not sheet_name:
            raise ToolError("sheet_name is required for 'copy' action")

        worksheet = get_worksheet(workbook, sheet_name)

        # Determine target workbook
        target_wb = workbook
        if target_workbook:
            target_wb = get_workbook(app, target_workbook)

        # Copy sheet
        if reference_sheet and position:
            ref_sheet = get_worksheet(target_wb, reference_sheet)
            if position.lower() == "before":
                worksheet.Copy(Before=ref_sheet)
            else:
                worksheet.Copy(After=ref_sheet)
        else:
            # Copy to end of target workbook
            worksheet.Copy(After=target_wb.Worksheets(target_wb.Worksheets.Count))

        # Rename if target_name provided
        if target_name:
            target_wb.Worksheets(target_wb.Worksheets.Count).Name = target_name

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "copied",
            "source_sheet": sheet_name,
            "target_workbook": target_workbook or workbook_name,
            "target_sheet": target_name or f"{sheet_name} (2)",
            "message": f"Copied sheet '{sheet_name}'"
        }

    elif action == "move":
        if not sheet_name:
            raise ToolError("sheet_name is required for 'move' action")

        worksheet = get_worksheet(workbook, sheet_name)

        # Determine target workbook
        target_wb = workbook
        if target_workbook:
            target_wb = get_workbook(app, target_workbook)

        # Move sheet
        if reference_sheet and position:
            ref_sheet = get_worksheet(target_wb, reference_sheet)
            if position.lower() == "before":
                worksheet.Move(Before=ref_sheet)
            else:
                worksheet.Move(After=ref_sheet)
        else:
            # Move to end
            worksheet.Move(After=target_wb.Worksheets(target_wb.Worksheets.Count))

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "moved",
            "sheet": sheet_name,
            "target_workbook": target_workbook or workbook_name,
            "message": f"Moved sheet '{sheet_name}'"
        }

    elif action == "rename":
        if not sheet_name:
            raise ToolError("sheet_name is required for 'rename' action")
        if not target_name:
            raise ToolError("target_name is required for 'rename' action")

        worksheet = get_worksheet(workbook, sheet_name)
        old_name = worksheet.Name
        worksheet.Name = target_name

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "renamed",
            "old_name": old_name,
            "new_name": target_name,
            "message": f"Renamed sheet '{old_name}' to '{target_name}'"
        }

    else:
        raise ToolError(
            f"Invalid action '{action}'. Must be one of: "
            "add, remove, hide, unhide, copy, move, rename"
        )
