"""Workbook-level operations for ExceLLM MCP server.

Contains tools for listing workbooks, selecting ranges, and validation.
"""

import logging
from typing import Any, Dict, List

from ..core.connection import (
    _init_com,
    get_excel_app,
    get_workbook,
    get_worksheet,
)

logger = logging.getLogger(__name__)


def list_workbooks_sync() -> List[Dict[str, Any]]:
    """List all open workbooks with their sheet information.

    Returns:
        List of dictionaries containing workbook name and sheets
    """
    _init_com()

    app = get_excel_app()
    workbooks = []

    try:
        wb_count = app.Workbooks.Count
    except Exception:
        logger.warning("No open workbooks found")
        return workbooks

    for i in range(1, wb_count + 1):
        try:
            workbook = app.Workbooks(i)
            sheets = []

            for j in range(1, workbook.Worksheets.Count + 1):
                worksheet = workbook.Worksheets(j)
                sheet_name = worksheet.Name
                # Check visibility: -1 = visible, 0 = hidden, 2 = very hidden
                is_hidden = worksheet.Visible != -1
                sheets.append({
                    "name": sheet_name,
                    "hidden": is_hidden
                })

            workbooks.append({
                "name": workbook.Name,
                "sheets": sheets
            })
        except Exception as e:
            logger.warning(f"Could not list sheets for workbook {i}: {str(e)[:100]}")
            try:
                workbooks.append({
                    "name": workbook.Name,
                    "sheets": []
                })
            except Exception:
                workbooks.append({
                    "name": "Unknown",
                    "sheets": []
                })

    return workbooks


def select_range_sync(
    workbook_name: str,
    sheet_name: str,
    reference: str,
) -> Dict[str, Any]:
    """Activate a workbook, worksheet, and select a range.

    This moves the focus in the Excel UI to the specified location.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        reference: Cell or range reference

    Returns:
        Dictionary with success status and message
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    # Activate workbook window
    try:
        workbook.Activate()
    except Exception:
        pass  # Ignore if fails

    # Activate worksheet
    worksheet.Activate()

    # Select target range
    rng = worksheet.Range(reference)
    rng.Select()

    return {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "reference": reference,
        "message": f"Selected {reference} in {sheet_name}"
    }


def get_sheet_names_sync(workbook_name: str) -> List[str]:
    """Get sheet names for a workbook.

    Args:
        workbook_name: Name of the workbook

    Returns:
        List of sheet names
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    sheets = []

    for i in range(1, workbook.Worksheets.Count + 1):
        sheets.append(workbook.Worksheets(i).Name)

    return sheets


def validate_cell_reference_result(cell: str) -> Dict[str, Any]:
    """Validate an Excel cell reference format.

    Args:
        cell: Cell reference to validate

    Returns:
        Dictionary with validation result
    """
    import re

    # Pattern for valid cell references: 1-3 letters + 1-7 digits
    pattern = r'^[A-Za-z]{1,3}[1-9][0-9]{0,6}$'

    if re.match(pattern, cell):
        return {
            "valid": True,
            "cell": cell,
            "message": "Valid cell reference"
        }
    else:
        return {
            "valid": False,
            "cell": cell,
            "message": "Invalid cell reference. Expected format: A1, B5, Z100"
        }
