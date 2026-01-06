"""Range operations for ExceLLM MCP server.

Contains tools for copy_range, sort_range, and find_replace.
These are new tools added in the refactoring.
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
    number_to_column,
)

logger = logging.getLogger(__name__)


def copy_range_sync(
    source_workbook: str,
    source_sheet: str,
    source_range: str,
    target_workbook: Optional[str] = None,
    target_sheet: Optional[str] = None,
    target_cell: str = "A1",
    include_formatting: bool = True,
) -> Dict[str, Any]:
    """Copy range to another location with optional formatting.

    Args:
        source_workbook: Name of source workbook
        source_sheet: Name of source worksheet
        source_range: Range to copy (e.g., "A1:D10")
        target_workbook: Name of target workbook (defaults to source)
        target_sheet: Name of target worksheet (defaults to source)
        target_cell: Top-left cell of paste destination
        include_formatting: If True, copy formatting along with values

    Returns:
        Dictionary with operation result
    """
    _init_com()

    app = get_excel_app()

    # Get source
    src_workbook = get_workbook(app, source_workbook)
    src_worksheet = get_worksheet(src_workbook, source_sheet)
    src_range = src_worksheet.Range(source_range)

    # Get target
    tgt_workbook_name = target_workbook or source_workbook
    tgt_sheet_name = target_sheet or source_sheet

    tgt_workbook = get_workbook(app, tgt_workbook_name)
    tgt_worksheet = get_worksheet(tgt_workbook, tgt_sheet_name)
    tgt_range = tgt_worksheet.Range(target_cell)

    # Copy data
    src_range.Copy()

    if include_formatting:
        # Paste with all formatting
        tgt_range.PasteSpecial(Paste=-4104)  # xlPasteAll
    else:
        # Paste values only
        tgt_range.PasteSpecial(Paste=-4163)  # xlPasteValues

    # Clear clipboard
    try:
        app.CutCopyMode = False
    except Exception:
        pass

    # Calculate destination range
    src_rows = src_range.Rows.Count
    src_cols = src_range.Columns.Count
    tgt_start_row = tgt_range.Row
    tgt_start_col = tgt_range.Column
    tgt_end_row = tgt_start_row + src_rows - 1
    tgt_end_col = tgt_start_col + src_cols - 1

    target_range_addr = f"{number_to_column(tgt_start_col)}{tgt_start_row}:{number_to_column(tgt_end_col)}{tgt_end_row}"

    return {
        "success": True,
        "source": {
            "workbook": source_workbook,
            "sheet": source_sheet,
            "range": source_range,
        },
        "target": {
            "workbook": tgt_workbook_name,
            "sheet": tgt_sheet_name,
            "range": target_range_addr,
        },
        "cells_copied": src_rows * src_cols,
        "include_formatting": include_formatting,
        "message": f"Copied {source_range} to {target_range_addr}"
    }


def sort_range_sync(
    workbook_name: str,
    sheet_name: str,
    range_str: str,
    sort_by: List[Dict[str, Any]],
    has_header: bool = True,
) -> Dict[str, Any]:
    """Sort data in a range by one or more columns.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range_str: Range to sort (e.g., "A1:D100")
        sort_by: List of sort specifications:
            [{"column": "B", "order": "asc"}, {"column": "C", "order": "desc"}]
        has_header: If True, first row is treated as header

    Returns:
        Dictionary with operation result
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    rng = worksheet.Range(range_str)

    # Clear existing sort fields
    worksheet.Sort.SortFields.Clear()

    # Add sort fields
    for sort_spec in sort_by:
        col = sort_spec.get("column", "A")
        order = sort_spec.get("order", "asc").lower()

        # Convert column letter to range reference
        if col.isdigit():
            col = number_to_column(int(col))

        # Get the column range within the sort range
        start_row = rng.Row
        end_row = rng.Row + rng.Rows.Count - 1
        key_range = worksheet.Range(f"{col}{start_row}:{col}{end_row}")

        # Order: 1 = xlAscending, 2 = xlDescending
        sort_order = 1 if order == "asc" else 2

        worksheet.Sort.SortFields.Add(
            Key=key_range,
            SortOn=0,  # xlSortOnValues
            Order=sort_order,
        )

    # Apply sort
    # Header: 1 = xlYes, 2 = xlNo
    header = 1 if has_header else 2

    worksheet.Sort.SetRange(rng)
    worksheet.Sort.Header = header
    worksheet.Sort.MatchCase = False
    worksheet.Sort.Orientation = 1  # xlTopToBottom
    worksheet.Sort.Apply()

    return {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "range": range_str,
        "sort_by": sort_by,
        "has_header": has_header,
        "rows_sorted": rng.Rows.Count - (1 if has_header else 0),
        "message": f"Sorted {range_str} by {len(sort_by)} column(s)"
    }


def find_replace_sync(
    workbook_name: str,
    sheet_name: Optional[str] = None,
    find_value: str = "",
    replace_value: str = "",
    match_case: bool = False,
    match_entire_cell: bool = False,
    range_str: Optional[str] = None,
    preview_only: bool = False,
) -> Dict[str, Any]:
    """Find and replace values in a sheet or workbook.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet (None = all sheets)
        find_value: Value to find
        replace_value: Value to replace with
        match_case: If True, match case exactly
        match_entire_cell: If True, match entire cell content
        range_str: Specific range to search (defaults to UsedRange)
        preview_only: If True, only count matches without replacing

    Returns:
        Dictionary with operation result
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)

    # LookAt: 1 = xlWhole, 2 = xlPart
    look_at = 1 if match_entire_cell else 2

    results = []
    total_replacements = 0

    if sheet_name:
        # Single sheet
        sheets = [get_worksheet(workbook, sheet_name)]
    else:
        # All sheets
        sheets = [workbook.Worksheets(i) for i in range(1, workbook.Worksheets.Count + 1)]

    for ws in sheets:
        # Get search range
        if range_str:
            search_range = ws.Range(range_str)
        else:
            search_range = ws.UsedRange

        # Count matches first
        matches = 0
        find_result = search_range.Find(
            What=find_value,
            LookIn=-4163,  # xlValues
            LookAt=look_at,
            MatchCase=match_case,
        )

        if find_result:
            first_address = find_result.Address
            matches = 1

            while True:
                find_result = search_range.FindNext(find_result)
                if not find_result or find_result.Address == first_address:
                    break
                matches += 1

        # Perform replacement if not preview
        replaced = 0
        if not preview_only and matches > 0:
            replaced = search_range.Replace(
                What=find_value,
                Replacement=replace_value,
                LookAt=look_at,
                MatchCase=match_case,
            )
            # Replace returns True/False, so we use the count we found
            replaced = matches if replaced else 0

        if matches > 0:
            results.append({
                "sheet": ws.Name,
                "matches_found": matches,
                "replacements_made": replaced if not preview_only else 0,
            })
            total_replacements += replaced if not preview_only else 0

    total_matches = sum(r.get("matches_found", 0) for r in results)

    return {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "find_value": find_value,
        "replace_value": replace_value,
        "match_case": match_case,
        "match_entire_cell": match_entire_cell,
        "preview_only": preview_only,
        "total_matches": total_matches,
        "total_replacements": total_replacements,
        "results": results,
        "message": f"Found {total_matches} matches" + (f", replaced {total_replacements}" if not preview_only else " (preview mode)")
    }
