"""explore tool implementation.

Sheet-level radar with quick/deep modes for layout and structure detection.
"""

import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Tuple

import pythoncom
import win32com.client as win32

from .types import (
    TOOL_VERSION,
    CommentsNotes,
    ContentTypes,
    DataFootprint,
    ExploreRecommendations,
    ExploreResult,
    Layout,
    Meta,
    NextAction,
    Outlier,
    OutlierDistance,
    ReadHints,
    RecommendationReason,
    Region,
    WriteSafety,
)
from .utils import (
    build_range_address,
    compute_density,
    generate_sample_positions,
    is_cell_empty,
    normalize_address,
    number_to_column,
    parse_range_bounds,
    safe_bool,
)

# Excel constants
XL_CELL_TYPE_FORMULAS = -4123
XL_CELL_TYPE_CONSTANTS = 2


def explore_sync(sheet: str, mode: str = "quick") -> Dict[str, Any]:
    """Synchronous implementation of explore tool.

    Args:
        sheet: Sheet name or "ACTIVE" for active sheet
        mode: "quick" or "deep"

    Returns:
        Dictionary matching ExploreResult schema
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

    # Resolve sheet
    if sheet.upper() == "ACTIVE":
        try:
            ws = app.ActiveSheet
            sheet_name = ws.Name
        except Exception as e:
            raise RuntimeError(f"Cannot access active sheet: {e}")
    else:
        try:
            ws = wb.Worksheets(sheet)
            sheet_name = ws.Name
        except Exception as e:
            raise RuntimeError(f"Sheet '{sheet}' not found: {e}")

    # Get UsedRange
    used_range = ws.UsedRange
    used_range_addr = normalize_address(used_range.Address)

    # Parse bounds
    start_row, start_col, end_row, end_col = parse_range_bounds(used_range_addr)

    if mode == "deep":
        result = _explore_deep(ws, app, used_range, used_range_addr,
                                start_row, start_col, end_row, end_col, sheet_name)
    else:
        result = _explore_quick(ws, app, used_range, used_range_addr,
                                 start_row, start_col, end_row, end_col, sheet_name)

    duration_ms = int((time.perf_counter() - start_time) * 1000)

    # Build final result
    explore_result = ExploreResult(
        meta=Meta(
            tool="explore",
            version=TOOL_VERSION,
            timestamp=datetime.now(timezone.utc).isoformat(),
            durationMs=duration_ms,
            sheet=sheet_name,
            mode=mode,
        ),
        **result
    )

    return explore_result.model_dump()


def _explore_quick(
    ws, app, used_range, used_range_addr: str,
    start_row: int, start_col: int, end_row: int, end_col: int,
    sheet_name: str
) -> Dict[str, Any]:
    """Quick mode exploration - fast sampling-based.

    Args:
        ws: Worksheet COM object
        app: Application COM object
        used_range: UsedRange COM object
        used_range_addr: Normalized address string
        start_row, start_col, end_row, end_col: Range bounds
        sheet_name: Name of sheet

    Returns:
        Dictionary with explore results (without meta)
    """
    flags = []

    # Data footprint - quick approximation
    data_cell_count = None
    formulas_count = 0
    constants_count = 0

    try:
        try:
            formula_cells = used_range.SpecialCells(XL_CELL_TYPE_FORMULAS)
            formulas_count = formula_cells.Cells.Count
        except Exception:
            pass

        try:
            constant_cells = used_range.SpecialCells(XL_CELL_TYPE_CONSTANTS)
            constants_count = constant_cells.Cells.Count
        except Exception:
            pass

        data_cell_count = formulas_count + constants_count
    except Exception:
        pass

    total_rows = end_row - start_row + 1
    total_cols = end_col - start_col + 1
    total_cells = total_rows * total_cols

    # Quick sampling to find real data bounds approximation
    real_data_bounds = None
    non_empty_rows = None
    non_empty_cols = None

    # Sample positions
    sample_positions = generate_sample_positions(
        start_row, start_col, end_row, end_col,
        max_tiles=10, probes_per_tile=2
    )

    # Track occupied tiles for region detection
    min_data_row = end_row
    max_data_row = start_row
    min_data_col = end_col
    max_data_col = start_col
    non_empty_count = 0

    for row, col in sample_positions:
        try:
            cell = ws.Cells(row, col)
            value = cell.Value2
            if not is_cell_empty(value):
                non_empty_count += 1
                min_data_row = min(min_data_row, row)
                max_data_row = max(max_data_row, row)
                min_data_col = min(min_data_col, col)
                max_data_col = max(max_data_col, col)
        except Exception:
            pass

    if non_empty_count > 0:
        real_data_bounds = build_range_address(
            min_data_row, min_data_col, max_data_row, max_data_col
        )
        # Rough approximation of rows/cols
        non_empty_rows = max_data_row - min_data_row + 1
        non_empty_cols = max_data_col - min_data_col + 1

    data_footprint = DataFootprint(
        usedRangeReported=used_range_addr,
        realDataBounds=real_data_bounds,
        dataCellCount=data_cell_count,
        nonEmptyRows=non_empty_rows,
        nonEmptyCols=non_empty_cols,
    )

    # Single primary region in quick mode
    regions = []
    if real_data_bounds:
        regions.append(Region(
            id="R1",
            range=real_data_bounds,
            density=compute_density(non_empty_count, len(sample_positions)),
            headerCandidateRows=[min_data_row] if non_empty_count > 0 else [],
        ))

    # Outliers - check corners outside primary region
    outliers = _detect_outliers_quick(
        ws, start_row, start_col, end_row, end_col,
        min_data_row, min_data_col, max_data_row, max_data_col
    )

    if outliers:
        flags.append("OUTLIER_DATA_PRESENT")

    # Layout info
    layout = _get_layout_info(ws, used_range, app)

    # Content types
    content_types = ContentTypes(
        constants=constants_count if constants_count > 0 else None,
        formulas=formulas_count if formulas_count > 0 else None,
    )

    # Comments/Notes
    comments_notes = _get_comments_info(ws)

    # Flags
    if data_cell_count is not None and data_cell_count == 0:
        flags.append("EMPTY_OR_NEAR_EMPTY")

    if total_rows > 200000 or total_cols > 200:
        flags.append("EXTREME_USED_RANGE")

    # Check for inflated used range
    if data_cell_count is not None and total_cells > 0:
        ratio = data_cell_count / total_cells
        if ratio < 0.01 and total_cells > 1000:
            flags.append("USED_RANGE_INFLATED")

    if layout.tables["count"] and layout.tables["count"] > 0:
        flags.append("HAS_TABLE_OBJECTS")

    if layout.autoFilter:
        flags.append("HAS_FILTER")

    if layout.freezePanesAt:
        flags.append("HAS_FREEZE_PANES")

    if layout.mergedCellsCount and layout.mergedCellsCount > 0:
        flags.append("MERGED_CELLS_PRESENT")

    if comments_notes.count and comments_notes.count > 0:
        flags.append("HAS_COMMENTS_NOTES")

    if formulas_count > 0:
        flags.append("HAS_FORMULAS")

    # Read hints
    read_hints = _build_read_hints(
        regions, outliers, min_data_row, min_data_col, max_data_row, max_data_col
    )

    # Recommendations for deep
    should_run_deep = False
    reasons = []

    if "OUTLIER_DATA_PRESENT" in flags:
        should_run_deep = True
        reasons.append(RecommendationReason(
            flag="OUTLIER_DATA_PRESENT",
            severity="medium",
            why="Outlier data blocks detected outside primary region",
        ))

    if "USED_RANGE_INFLATED" in flags or "EXTREME_USED_RANGE" in flags:
        should_run_deep = True
        reasons.append(RecommendationReason(
            flag="USED_RANGE_INFLATED" if "USED_RANGE_INFLATED" in flags else "EXTREME_USED_RANGE",
            severity="medium",
            why="Used range may not reflect actual data bounds",
        ))

    if "MERGED_CELLS_PRESENT" in flags:
        reasons.append(RecommendationReason(
            flag="MERGED_CELLS_PRESENT",
            severity="low",
            why="Merged cells may affect header detection",
        ))

    next_actions = []
    if should_run_deep:
        next_actions.append(NextAction(
            tool="explore",
            scope={"sheet": sheet_name},
            mode="deep",
        ))

    recommendations = ExploreRecommendations(
        shouldRunDeep=should_run_deep,
        reasons=reasons,
        nextActions=next_actions,
    )

    return {
        "dataFootprint": data_footprint,
        "regions": regions,
        "outliers": outliers,
        "layout": layout,
        "contentTypes": content_types,
        "commentsNotes": comments_notes,
        "flags": flags,
        "readHints": read_hints,
        "recommendations": recommendations,
    }


def _explore_deep(
    ws, app, used_range, used_range_addr: str,
    start_row: int, start_col: int, end_row: int, end_col: int,
    sheet_name: str
) -> Dict[str, Any]:
    """Deep mode exploration - more thorough analysis.

    Args:
        ws: Worksheet COM object
        app: Application COM object
        used_range: UsedRange COM object
        used_range_addr: Normalized address string
        start_row, start_col, end_row, end_col: Range bounds
        sheet_name: Name of sheet

    Returns:
        Dictionary with explore results (without meta)
    """
    flags = []

    # Data footprint - get actual counts
    data_cell_count = 0
    formulas_count = 0
    constants_count = 0

    try:
        try:
            formula_cells = used_range.SpecialCells(XL_CELL_TYPE_FORMULAS)
            formulas_count = formula_cells.Cells.Count
        except Exception:
            pass

        try:
            constant_cells = used_range.SpecialCells(XL_CELL_TYPE_CONSTANTS)
            constants_count = constant_cells.Cells.Count
        except Exception:
            pass

        data_cell_count = formulas_count + constants_count
    except Exception:
        pass

    total_rows = end_row - start_row + 1
    total_cols = end_col - start_col + 1
    total_cells = total_rows * total_cols

    # Find real data bounds via backward scanning
    real_start_row, real_start_col, real_end_row, real_end_col = _find_real_data_bounds(
        ws, start_row, start_col, end_row, end_col
    )

    real_data_bounds = None
    non_empty_rows = None
    non_empty_cols = None

    if real_start_row <= real_end_row and real_start_col <= real_end_col:
        real_data_bounds = build_range_address(
            real_start_row, real_start_col, real_end_row, real_end_col
        )
        non_empty_rows = real_end_row - real_start_row + 1
        non_empty_cols = real_end_col - real_start_col + 1

    data_footprint = DataFootprint(
        usedRangeReported=used_range_addr,
        realDataBounds=real_data_bounds,
        dataCellCount=data_cell_count if data_cell_count > 0 else None,
        nonEmptyRows=non_empty_rows,
        nonEmptyCols=non_empty_cols,
    )

    # Detect regions via blank row runs
    regions = _detect_regions_deep(
        ws, real_start_row, real_start_col, real_end_row, real_end_col
    )

    if len(regions) > 1:
        flags.append("MULTI_REGION_SHEET")

    # Detect outliers with distance calculation
    outliers = _detect_outliers_deep(
        ws, start_row, start_col, end_row, end_col,
        real_start_row, real_start_col, real_end_row, real_end_col
    )

    if outliers:
        flags.append("OUTLIER_DATA_PRESENT")

    # Layout info
    layout = _get_layout_info(ws, used_range, app)

    # Content types
    content_types = ContentTypes(
        constants=constants_count if constants_count > 0 else None,
        formulas=formulas_count if formulas_count > 0 else None,
    )

    # Comments/Notes - more thorough
    comments_notes = _get_comments_info(ws, detailed=True)

    # Flags
    if data_cell_count == 0:
        flags.append("EMPTY_OR_NEAR_EMPTY")

    if total_rows > 200000 or total_cols > 200:
        flags.append("EXTREME_USED_RANGE")

    if data_cell_count > 0 and total_cells > 0:
        ratio = data_cell_count / total_cells
        if ratio < 0.01 and total_cells > 1000:
            flags.append("USED_RANGE_INFLATED")

    if layout.tables["count"] and layout.tables["count"] > 0:
        flags.append("HAS_TABLE_OBJECTS")

    if layout.autoFilter:
        flags.append("HAS_FILTER")

    if layout.freezePanesAt:
        flags.append("HAS_FREEZE_PANES")

    if layout.mergedCellsCount and layout.mergedCellsCount > 0:
        flags.append("MERGED_CELLS_PRESENT")

    if comments_notes.count and comments_notes.count > 0:
        flags.append("HAS_COMMENTS_NOTES")

    if formulas_count > 0:
        flags.append("HAS_FORMULAS")

    # Read hints with region header scans
    read_hints = _build_read_hints_deep(
        regions, outliers, real_start_row, real_start_col, real_end_row, real_end_col
    )

    # Deep mode - no need for further deep exploration
    recommendations = ExploreRecommendations(
        shouldRunDeep=False,
        reasons=[],
        nextActions=[],
    )

    return {
        "dataFootprint": data_footprint,
        "regions": regions,
        "outliers": outliers,
        "layout": layout,
        "contentTypes": content_types,
        "commentsNotes": comments_notes,
        "flags": flags,
        "readHints": read_hints,
        "recommendations": recommendations,
    }


def _find_real_data_bounds(
    ws, start_row: int, start_col: int, end_row: int, end_col: int
) -> Tuple[int, int, int, int]:
    """Find actual data bounds by scanning backwards.

    Args:
        ws: Worksheet COM object
        start_row, start_col, end_row, end_col: UsedRange bounds

    Returns:
        Tuple of (real_start_row, real_start_col, real_end_row, real_end_col)
    """
    real_start_row = start_row
    real_start_col = start_col
    real_end_row = start_row
    real_end_col = start_col

    found_data = False

    # Find last row with data (scan backward with steps)
    step = max(1, (end_row - start_row) // 20)

    # First pass: find approximate last row
    approx_last_row = start_row
    for row in range(end_row, start_row - 1, -step):
        row_empty = True
        for col in range(start_col, min(start_col + 10, end_col + 1)):
            try:
                if not is_cell_empty(ws.Cells(row, col).Value2):
                    row_empty = False
                    break
            except Exception:
                pass

        if not row_empty:
            approx_last_row = row
            break

    # Refine: scan from approx to find exact
    for row in range(min(approx_last_row + step, end_row), approx_last_row - 1, -1):
        row_empty = True
        for col in range(start_col, min(start_col + 20, end_col + 1)):
            try:
                if not is_cell_empty(ws.Cells(row, col).Value2):
                    row_empty = False
                    real_end_row = row
                    found_data = True
                    break
            except Exception:
                pass
        if not row_empty:
            break

    # Find last column with data
    step = max(1, (end_col - start_col) // 20)
    approx_last_col = start_col

    for col in range(end_col, start_col - 1, -step):
        col_empty = True
        for row in range(start_row, min(start_row + 10, end_row + 1)):
            try:
                if not is_cell_empty(ws.Cells(row, col).Value2):
                    col_empty = False
                    break
            except Exception:
                pass

        if not col_empty:
            approx_last_col = col
            break

    # Refine column
    for col in range(min(approx_last_col + step, end_col), approx_last_col - 1, -1):
        col_empty = True
        for row in range(start_row, min(start_row + 20, end_row + 1)):
            try:
                if not is_cell_empty(ws.Cells(row, col).Value2):
                    col_empty = False
                    real_end_col = col
                    found_data = True
                    break
            except Exception:
                pass
        if not col_empty:
            break

    if not found_data:
        return (start_row, start_col, start_row, start_col)

    return (real_start_row, real_start_col, real_end_row, real_end_col)


def _detect_regions_deep(
    ws, start_row: int, start_col: int, end_row: int, end_col: int
) -> List[Region]:
    """Detect multiple regions via blank row runs.

    Args:
        ws: Worksheet COM object
        start_row, start_col, end_row, end_col: Data bounds

    Returns:
        List of Region objects
    """
    regions = []

    if start_row > end_row or start_col > end_col:
        return regions

    # Track row emptiness
    consecutive_empty = 0
    region_start_row = start_row
    region_id = 1

    # Check sample columns for each row
    check_cols = list(range(start_col, min(start_col + 5, end_col + 1)))

    for row in range(start_row, end_row + 1):
        row_empty = True
        for col in check_cols:
            try:
                if not is_cell_empty(ws.Cells(row, col).Value2):
                    row_empty = False
                    break
            except Exception:
                pass

        if row_empty:
            consecutive_empty += 1
        else:
            if consecutive_empty >= 2 and row > region_start_row + 1:
                # End of previous region
                region_end_row = row - consecutive_empty - 1
                if region_end_row >= region_start_row:
                    regions.append(Region(
                        id=f"R{region_id}",
                        range=build_range_address(
                            region_start_row, start_col, region_end_row, end_col
                        ),
                        density=None,
                        headerCandidateRows=[region_start_row],
                    ))
                    region_id += 1
                region_start_row = row
            consecutive_empty = 0

    # Final region
    if region_start_row <= end_row:
        regions.append(Region(
            id=f"R{region_id}",
            range=build_range_address(region_start_row, start_col, end_row, end_col),
            density=None,
            headerCandidateRows=[region_start_row],
        ))

    return regions if regions else [Region(
        id="R1",
        range=build_range_address(start_row, start_col, end_row, end_col),
        density=None,
        headerCandidateRows=[start_row],
    )]


def _detect_outliers_quick(
    ws, start_row: int, start_col: int, end_row: int, end_col: int,
    data_start_row: int, data_start_col: int, data_end_row: int, data_end_col: int
) -> List[Outlier]:
    """Quick outlier detection by checking corners.

    Args:
        ws: Worksheet COM object
        start_row, start_col, end_row, end_col: UsedRange bounds
        data_start_row, data_start_col, data_end_row, data_end_col: Primary data bounds

    Returns:
        List of Outlier objects
    """
    outliers = []
    outlier_id = 1

    # Check areas outside primary data region
    # Top area (above data)
    if data_start_row > start_row:
        found = _check_area_for_data(ws, start_row, start_col, data_start_row - 1, end_col)
        if found:
            outliers.append(Outlier(
                id=f"O{outlier_id}",
                range=build_range_address(start_row, start_col, data_start_row - 1, end_col),
                cellCount=None,
                distanceFromPrimary=None,
            ))
            outlier_id += 1

    # Bottom area (below data)
    if data_end_row < end_row:
        found = _check_area_for_data(ws, data_end_row + 1, start_col, end_row, end_col)
        if found:
            outliers.append(Outlier(
                id=f"O{outlier_id}",
                range=build_range_address(data_end_row + 1, start_col, end_row, end_col),
                cellCount=None,
                distanceFromPrimary=OutlierDistance(
                    rows=1,
                    cols=0,
                ),
            ))
            outlier_id += 1

    # Right area (to the right of data)
    if data_end_col < end_col:
        found = _check_area_for_data(ws, start_row, data_end_col + 1, end_row, end_col)
        if found:
            outliers.append(Outlier(
                id=f"O{outlier_id}",
                range=build_range_address(start_row, data_end_col + 1, end_row, end_col),
                cellCount=None,
                distanceFromPrimary=OutlierDistance(
                    rows=0,
                    cols=1,
                ),
            ))

    return outliers


def _detect_outliers_deep(
    ws, start_row: int, start_col: int, end_row: int, end_col: int,
    data_start_row: int, data_start_col: int, data_end_row: int, data_end_col: int
) -> List[Outlier]:
    """Deep outlier detection with distance calculation.

    Similar to quick but with more accurate bounds and distances.
    """
    outliers = []
    outlier_id = 1

    # Check bottom area
    if data_end_row + 2 <= end_row:
        # Skip one row gap, then check
        for check_row in range(data_end_row + 2, min(end_row + 1, data_end_row + 50)):
            found_row = False
            for col in range(start_col, min(start_col + 10, end_col + 1)):
                try:
                    if not is_cell_empty(ws.Cells(check_row, col).Value2):
                        found_row = True
                        break
                except Exception:
                    pass

            if found_row:
                outliers.append(Outlier(
                    id=f"O{outlier_id}",
                    range=build_range_address(check_row, start_col, end_row, end_col),
                    cellCount=None,
                    distanceFromPrimary=OutlierDistance(
                        rows=check_row - data_end_row,
                        cols=0,
                    ),
                ))
                outlier_id += 1
                break

    # Check right area
    if data_end_col + 2 <= end_col:
        for check_col in range(data_end_col + 2, min(end_col + 1, data_end_col + 20)):
            found_col = False
            for row in range(start_row, min(start_row + 10, end_row + 1)):
                try:
                    if not is_cell_empty(ws.Cells(row, check_col).Value2):
                        found_col = True
                        break
                except Exception:
                    pass

            if found_col:
                outliers.append(Outlier(
                    id=f"O{outlier_id}",
                    range=build_range_address(start_row, check_col, end_row, end_col),
                    cellCount=None,
                    distanceFromPrimary=OutlierDistance(
                        rows=0,
                        cols=check_col - data_end_col,
                    ),
                ))
                break

    return outliers


def _check_area_for_data(
    ws, start_row: int, start_col: int, end_row: int, end_col: int
) -> bool:
    """Check if an area contains any data via sampling.

    Args:
        ws: Worksheet COM object
        start_row, start_col, end_row, end_col: Area bounds

    Returns:
        True if any data found
    """
    # Sample a few positions
    check_positions = [
        (start_row, start_col),
        (start_row, end_col),
        (end_row, start_col),
        (end_row, end_col),
        ((start_row + end_row) // 2, (start_col + end_col) // 2),
    ]

    for row, col in check_positions:
        if row < start_row or row > end_row or col < start_col or col > end_col:
            continue
        try:
            if not is_cell_empty(ws.Cells(row, col).Value2):
                return True
        except Exception:
            pass

    return False


def _get_layout_info(ws, used_range, app) -> Layout:
    """Get layout information for a sheet.

    Args:
        ws: Worksheet COM object
        used_range: UsedRange COM object
        app: Application COM object

    Returns:
        Layout object
    """
    # Tables
    tables = {"count": None, "names": []}
    try:
        list_objects = ws.ListObjects
        if list_objects:
            tables["count"] = list_objects.Count
            for i in range(1, min(list_objects.Count + 1, 11)):
                try:
                    tables["names"].append(list_objects.Item(i).Name)
                except Exception:
                    pass
    except Exception:
        pass

    # Merged cells
    merged_count = None
    merged_sample = []
    try:
        merge_areas = used_range.MergeAreas
        if merge_areas:
            merged_count = merge_areas.Count
            for i in range(1, min(merge_areas.Count + 1, 11)):
                try:
                    merged_sample.append(normalize_address(merge_areas.Item(i).Address))
                except Exception:
                    pass
    except Exception:
        pass

    # Freeze panes
    freeze_at = None
    try:
        window = app.ActiveWindow
        if window and window.FreezePanes:
            split_row = int(window.SplitRow)
            split_col = int(window.SplitColumn)
            if split_row > 0 or split_col > 0:
                freeze_at = f"{number_to_column(split_col + 1)}{split_row + 1}"
    except Exception:
        pass

    # AutoFilter
    auto_filter = None
    try:
        auto_filter = safe_bool(ws.AutoFilterMode)
    except Exception:
        pass

    # Hidden rows/cols - sampled
    hidden_rows = None
    hidden_cols = None
    try:
        hidden_row_count = 0
        row_count = used_range.Rows.Count
        for r in range(1, min(row_count + 1, 501), 25):
            try:
                if used_range.Rows(r).Hidden:
                    hidden_row_count += 1
            except Exception:
                pass
        hidden_rows = hidden_row_count * 25 if hidden_row_count > 0 else 0

        hidden_col_count = 0
        col_count = used_range.Columns.Count
        for c in range(1, min(col_count + 1, 51), 5):
            try:
                if used_range.Columns(c).Hidden:
                    hidden_col_count += 1
            except Exception:
                pass
        hidden_cols = hidden_col_count * 5 if hidden_col_count > 0 else 0
    except Exception:
        pass

    return Layout(
        tables=tables,
        mergedCellsCount=merged_count,
        mergedRangesSample=merged_sample,
        freezePanesAt=freeze_at,
        autoFilter=auto_filter,
        hiddenRows=hidden_rows,
        hiddenCols=hidden_cols,
    )


def _get_comments_info(ws, detailed: bool = False) -> CommentsNotes:
    """Get comments/notes information.

    Args:
        ws: Worksheet COM object
        detailed: If True, include range samples

    Returns:
        CommentsNotes object
    """
    count = None
    ranges_sample = []

    try:
        comments = ws.Comments
        if comments:
            count = comments.Count
            if detailed and count > 0:
                for i in range(1, min(count + 1, 11)):
                    try:
                        comment = comments.Item(i)
                        cell_addr = normalize_address(comment.Parent.Address)
                        ranges_sample.append(cell_addr)
                    except Exception:
                        pass
    except Exception:
        pass

    return CommentsNotes(
        count=count,
        rangesSample=ranges_sample,
    )


def _build_read_hints(
    regions: List[Region], outliers: List[Outlier],
    start_row: int, start_col: int, end_row: int, end_col: int
) -> ReadHints:
    """Build read hints for quick mode.

    Args:
        regions: Detected regions
        outliers: Detected outliers
        start_row, start_col, end_row, end_col: Primary data bounds

    Returns:
        ReadHints object
    """
    primary_region_id = regions[0].id if regions else None

    # Suggested header scan - first 25 rows of primary region
    suggested_header = None
    if regions:
        r_start, r_start_col, r_end, r_end_col = parse_range_bounds(regions[0].range)
        header_end_row = min(r_start + 24, r_end)
        suggested_header = build_range_address(r_start, r_start_col, header_end_row, r_end_col)

    # Suggested tail scan - last 25 rows
    suggested_tail = None
    if regions and end_row > start_row:
        r_start, r_start_col, r_end, r_end_col = parse_range_bounds(regions[0].range)
        tail_start_row = max(r_end - 24, r_start)
        suggested_tail = build_range_address(tail_start_row, r_start_col, r_end, r_end_col)

    # Suggested body - skip first row (header), take rest
    suggested_body = None
    if regions:
        r_start, r_start_col, r_end, r_end_col = parse_range_bounds(regions[0].range)
        if r_end > r_start:
            suggested_body = build_range_address(r_start + 1, r_start_col, r_end, r_end_col)

    # Outlier scans
    outlier_scans = [o.range for o in outliers]

    return ReadHints(
        primaryRegionId=primary_region_id,
        suggestedHeaderScan=suggested_header,
        suggestedTailScan=suggested_tail,
        suggestedBodyRead=suggested_body,
        suggestedOutlierScans=outlier_scans,
        suggestedRegionHeaderScans=None,
        writeSafety=None,
    )


def _build_read_hints_deep(
    regions: List[Region], outliers: List[Outlier],
    start_row: int, start_col: int, end_row: int, end_col: int
) -> ReadHints:
    """Build read hints for deep mode with region header scans.

    Args:
        regions: Detected regions
        outliers: Detected outliers
        start_row, start_col, end_row, end_col: Primary data bounds

    Returns:
        ReadHints object with region-specific hints
    """
    hints = _build_read_hints(regions, outliers, start_row, start_col, end_row, end_col)

    # Add per-region header scans
    region_headers = {}
    for region in regions:
        r_start, r_start_col, r_end, r_end_col = parse_range_bounds(region.range)
        header_end = min(r_start + 24, r_end)
        region_headers[region.id] = build_range_address(r_start, r_start_col, header_end, r_end_col)

    hints.suggestedRegionHeaderScans = region_headers if region_headers else None

    # Write safety hints
    if regions:
        hints.writeSafety = WriteSafety(
            appendColumnToRegionId=regions[0].id,
            avoidUsedRangeRightEdge=end_col > (start_col + 100),
        )

    return hints
