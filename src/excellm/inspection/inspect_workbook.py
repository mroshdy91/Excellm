"""inspect_workbook tool implementation.

Fast workbook-level inspection across all sheets without reading cell values.
"""

import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client as win32

from .types import (
    TOOL_VERSION,
    InspectWorkbookResult,
    LayoutFlags,
    Meta,
    Recommendations,
    SheetIndex,
    SheetScore,
    WorkbookInfo,
)
from .utils import (
    compute_density,
    normalize_address,
    parse_range_bounds,
    safe_bool,
    sheet_name_priority_boost,
)

# Excel constants
XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0
XL_SHEET_VERY_HIDDEN = 2

XL_CELL_TYPE_FORMULAS = -4123
XL_CELL_TYPE_CONSTANTS = 2


def inspect_workbook_sync() -> Dict[str, Any]:
    """Synchronous implementation of inspect_workbook.

    Connects to running Excel and inspects all sheets in the active workbook.

    Returns:
        Dictionary matching InspectWorkbookResult schema

    Raises:
        Exception if Excel not running or no workbook open
    """
    start_time = time.perf_counter()

    pythoncom.CoInitialize()

    try:
        app = win32.GetActiveObject("Excel.Application")
    except Exception as e:
        raise RuntimeError(f"Excel not running or not accessible: {e}")

    if app.Workbooks.Count == 0:
        raise RuntimeError("No workbook is open in Excel")

    wb = app.ActiveWorkbook
    if not wb:
        raise RuntimeError("No active workbook found")

    # Workbook info
    workbook_info = WorkbookInfo(
        name=wb.Name,
        path=wb.Path if wb.Path else None,
        readOnly=safe_bool(wb.ReadOnly),
        protected=safe_bool(wb.ProtectStructure),
    )

    # Active sheet
    active_sheet_name = ""
    try:
        active_sheet_name = app.ActiveSheet.Name
    except Exception:
        pass

    # Inspect each sheet
    sheets_index: List[SheetIndex] = []

    for i in range(1, wb.Worksheets.Count + 1):
        try:
            ws = wb.Worksheets(i)
            sheet_info = _inspect_sheet(ws, app)
            sheets_index.append(sheet_info)
        except Exception:
            # If sheet inspection fails, add minimal info
            try:
                ws = wb.Worksheets(i)
                sheets_index.append(SheetIndex(
                    name=ws.Name,
                    state="visible",
                    protected=False,
                    usedRangeReported="A1",
                    realDataBounds=None,
                    dataCellCount=None,
                    nonEmptyRows=None,
                    nonEmptyCols=None,
                    layoutFlags=LayoutFlags(),
                    flags=["INSPECTION_ERROR"],
                    score=SheetScore(priority=0.0, density=None),
                ))
            except Exception:
                pass

    # Compute recommendations
    recommendations = _compute_recommendations(sheets_index, active_sheet_name)

    duration_ms = int((time.perf_counter() - start_time) * 1000)

    result = InspectWorkbookResult(
        meta=Meta(
            tool="inspect_workbook",
            version=TOOL_VERSION,
            timestamp=datetime.now(timezone.utc).isoformat(),
            durationMs=duration_ms,
        ),
        workbook=workbook_info,
        activeSheet=active_sheet_name,
        sheetsIndex=sheets_index,
        recommendations=recommendations,
    )

    return result.model_dump()


def _inspect_sheet(ws, app) -> SheetIndex:
    """Inspect a single worksheet.

    Args:
        ws: Excel Worksheet COM object
        app: Excel Application COM object

    Returns:
        SheetIndex with sheet metadata
    """
    name = ws.Name

    # Sheet state
    visibility = ws.Visible
    if visibility == XL_SHEET_VISIBLE:
        state = "visible"
    elif visibility == XL_SHEET_VERY_HIDDEN:
        state = "veryHidden"
    else:
        state = "hidden"

    # Protected
    protected = safe_bool(ws.ProtectContents)

    # UsedRange
    used_range = ws.UsedRange
    used_range_addr = normalize_address(used_range.Address)

    # Parse bounds for analysis
    start_row, start_col, end_row, end_col = parse_range_bounds(used_range_addr)
    total_rows = end_row - start_row + 1
    total_cols = end_col - start_col + 1

    # Layout flags
    layout = _get_layout_flags(ws, used_range, app)

    # Quick data approximation (sampling-based)
    data_cell_count = None
    non_empty_rows = None
    non_empty_cols = None
    real_data_bounds = None

    # Use SpecialCells for quick count approximation
    try:
        # Count formulas
        formulas_count = 0
        try:
            formula_cells = used_range.SpecialCells(XL_CELL_TYPE_FORMULAS)
            formulas_count = formula_cells.Cells.Count
        except Exception:
            pass

        # Count constants
        constants_count = 0
        try:
            constant_cells = used_range.SpecialCells(XL_CELL_TYPE_CONSTANTS)
            constants_count = constant_cells.Cells.Count
        except Exception:
            pass

        data_cell_count = formulas_count + constants_count
        layout.formulasCount = formulas_count if formulas_count > 0 else None

    except Exception:
        pass

    # Compute flags
    flags = _compute_sheet_flags(
        state=state,
        total_rows=total_rows,
        total_cols=total_cols,
        data_cell_count=data_cell_count,
        layout=layout,
    )

    # Compute score
    total_cells = total_rows * total_cols
    density = compute_density(data_cell_count, total_cells) if data_cell_count else None

    priority = _compute_priority(
        state=state,
        name=name,
        data_cell_count=data_cell_count,
        layout=layout,
        flags=flags,
    )

    return SheetIndex(
        name=name,
        state=state,
        protected=protected,
        usedRangeReported=used_range_addr,
        realDataBounds=real_data_bounds,
        dataCellCount=data_cell_count,
        nonEmptyRows=non_empty_rows,
        nonEmptyCols=non_empty_cols,
        layoutFlags=layout,
        flags=flags,
        score=SheetScore(priority=priority, density=density),
    )


def _get_layout_flags(ws, used_range, app) -> LayoutFlags:
    """Get layout-related flags for a worksheet.

    Args:
        ws: Worksheet COM object
        used_range: UsedRange COM object
        app: Application COM object

    Returns:
        LayoutFlags with detected features
    """
    layout = LayoutFlags()

    # Tables (ListObjects)
    try:
        table_count = ws.ListObjects.Count
        layout.hasTableObjects = table_count > 0 if table_count is not None else None
    except Exception:
        layout.hasTableObjects = None

    # AutoFilter
    try:
        layout.hasAutoFilter = safe_bool(ws.AutoFilterMode)
    except Exception:
        layout.hasAutoFilter = None

    # Freeze panes - need to check window
    try:
        # Get the window for this workbook
        window = app.ActiveWindow
        if window and window.FreezePanes:
            layout.hasFreezePanes = True
        else:
            layout.hasFreezePanes = False
    except Exception:
        layout.hasFreezePanes = None

    # Merged cells - approximate count
    try:
        merge_areas = used_range.MergeAreas
        if merge_areas:
            layout.mergedCellsCount = merge_areas.Count
    except Exception:
        layout.mergedCellsCount = None

    # Comments
    try:
        comments = ws.Comments
        if comments:
            layout.commentsCount = comments.Count
    except Exception:
        layout.commentsCount = None

    # Hidden rows/cols - sample-based approximation
    try:
        hidden_rows = 0
        hidden_cols = 0

        # Sample every 50th row for hidden check
        row_count = used_range.Rows.Count
        for r in range(1, min(row_count + 1, 1001), 50):
            try:
                if used_range.Rows(r).Hidden:
                    hidden_rows += 1
            except Exception:
                pass

        # Scale up approximation
        if row_count > 1000:
            layout.hiddenRowsCount = hidden_rows * 50
        else:
            layout.hiddenRowsCount = hidden_rows * 50 if hidden_rows > 0 else 0

        # Sample columns
        col_count = used_range.Columns.Count
        for c in range(1, min(col_count + 1, 51), 5):
            try:
                if used_range.Columns(c).Hidden:
                    hidden_cols += 1
            except Exception:
                pass

        layout.hiddenColsCount = hidden_cols * 5 if hidden_cols > 0 else 0

    except Exception:
        layout.hiddenRowsCount = None
        layout.hiddenColsCount = None

    return layout


def _compute_sheet_flags(
    state: str,
    total_rows: int,
    total_cols: int,
    data_cell_count: Optional[int],
    layout: LayoutFlags,
) -> List[str]:
    """Compute flags for a sheet.

    Args:
        state: Sheet visibility state
        total_rows: Total rows in UsedRange
        total_cols: Total columns in UsedRange
        data_cell_count: Approximated data cell count
        layout: Layout flags

    Returns:
        List of flag strings
    """
    flags = []

    # Hidden sheet
    if state in ("hidden", "veryHidden"):
        flags.append("HIDDEN_SHEET")

    # Extreme used range
    if total_rows > 200000 or total_cols > 200:
        flags.append("EXTREME_USED_RANGE")

    # Empty or near empty
    if data_cell_count is not None and data_cell_count == 0:
        flags.append("EMPTY_OR_NEAR_EMPTY")
    elif data_cell_count is not None and data_cell_count < 10:
        flags.append("EMPTY_OR_NEAR_EMPTY")

    # Likely format only (huge used range but no data)
    total_cells = total_rows * total_cols
    if total_cells > 10000 and (data_cell_count or 0) < 100:
        if layout.formulasCount in (None, 0):
            flags.append("LIKELY_FORMAT_ONLY")

    # Used range inflated
    if data_cell_count is not None and total_cells > 0:
        ratio = data_cell_count / total_cells
        if ratio < 0.01 and total_cells > 1000:  # Less than 1% filled
            flags.append("USED_RANGE_INFLATED")

    # Layout features
    if layout.hasTableObjects:
        flags.append("HAS_TABLE_OBJECTS")

    if layout.hasAutoFilter:
        flags.append("HAS_FILTER")

    if layout.hasFreezePanes:
        flags.append("HAS_FREEZE_PANES")

    if layout.mergedCellsCount and layout.mergedCellsCount > 0:
        flags.append("MERGED_CELLS_PRESENT")

    if layout.commentsCount and layout.commentsCount > 0:
        flags.append("HAS_COMMENTS_NOTES")

    if layout.formulasCount and layout.formulasCount > 0:
        flags.append("HAS_FORMULAS")

    return flags


def _compute_priority(
    state: str,
    name: str,
    data_cell_count: Optional[int],
    layout: LayoutFlags,
    flags: List[str],
) -> float:
    """Compute priority score for a sheet.

    Args:
        state: Sheet visibility
        name: Sheet name
        data_cell_count: Data cell count
        layout: Layout flags
        flags: Computed flags

    Returns:
        Priority score 0..1
    """
    priority = 0.5  # Base

    # Visibility
    if state == "visible":
        priority += 0.2
    elif state == "hidden":
        priority -= 0.2
    elif state == "veryHidden":
        priority -= 0.4

    # Data presence
    if data_cell_count is not None:
        if data_cell_count > 1000:
            priority += 0.15
        elif data_cell_count > 100:
            priority += 0.1
        elif data_cell_count > 10:
            priority += 0.05
        elif data_cell_count == 0:
            priority -= 0.3

    # Tables boost
    if layout.hasTableObjects:
        priority += 0.1

    # Name heuristics
    priority += sheet_name_priority_boost(name)

    # Penalize problematic sheets
    if "LIKELY_FORMAT_ONLY" in flags:
        priority -= 0.3
    if "EXTREME_USED_RANGE" in flags and "EMPTY_OR_NEAR_EMPTY" in flags:
        priority -= 0.2

    # Clamp to 0..1
    return max(0.0, min(1.0, priority))


def _compute_recommendations(
    sheets: List[SheetIndex], active_sheet_name: str
) -> Recommendations:
    """Compute recommendations based on sheet analysis.

    Args:
        sheets: List of inspected sheets
        active_sheet_name: Currently active sheet name

    Returns:
        Recommendations object
    """
    # Sort by priority
    sorted_sheets = sorted(sheets, key=lambda s: s.score.priority, reverse=True)

    # Primary candidates - top sheets that are safe
    primary_candidates = []
    avoid_sheets = []

    for sheet in sorted_sheets:
        if "LIKELY_FORMAT_ONLY" in sheet.flags:
            avoid_sheets.append(sheet.name)
        elif "EXTREME_USED_RANGE" in sheet.flags and "EMPTY_OR_NEAR_EMPTY" in sheet.flags:
            avoid_sheets.append(sheet.name)
        elif sheet.state == "veryHidden":
            avoid_sheets.append(sheet.name)
        else:
            if len(primary_candidates) < 4:
                primary_candidates.append(sheet.name)

    # Next explore sheet
    next_explore = active_sheet_name

    # If active is in avoid list, pick top candidate
    if active_sheet_name in avoid_sheets and primary_candidates:
        next_explore = primary_candidates[0]
    elif not active_sheet_name and primary_candidates:
        next_explore = primary_candidates[0]

    return Recommendations(
        primaryCandidateSheets=primary_candidates,
        avoidSheets=avoid_sheets,
        nextExploreSheet=next_explore,
    )
