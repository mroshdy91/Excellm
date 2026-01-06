"""Write operations for ExceLLM MCP server.

Contains tools for writing to cells and ranges.
"""

import json
import logging
import re
from typing import Any, Dict, List, Optional

from ..core.connection import (
    _init_com,
    get_excel_app,
    get_workbook,
    get_worksheet,
)
from ..core.errors import ErrorCodes, ToolError
from ..core.utils import (
    column_to_number,
    number_to_column,
)
from ..validators import (
    parse_range,
)

# Maximum cells (rows × cols) allowed per write operation to prevent LLM context issues
# 500 cells = e.g., 125 rows × 4 cols
MAX_CELLS_LIMIT = 250

logger = logging.getLogger(__name__)

# Import audit logging
from ..core.audit import log_write


def _sanitize_value(value: Any) -> Any:
    """Sanitize a single value for Excel (convert dicts/lists to JSON strings)."""
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    # Convert everything else to JSON string
    try:
        return json.dumps(value)
    except Exception:
        return str(value)


def _sanitize_data(data: Any) -> Any:
    """Sanitize data (single value or 2D array) for Excel."""
    if isinstance(data, list):
        sanitized = []
        for row in data:
            if isinstance(row, list):
                sanitized.append([_sanitize_value(v) for v in row])
            else:
                sanitized.append(_sanitize_value(row))
        return sanitized
    return _sanitize_value(data)


def _normalize_jagged_array(data: List[List[Any]]) -> List[List[Any]]:
    """Normalize jagged array to rectangular by padding with empty values."""
    if not data:
        return []

    max_cols = 0
    for row in data:
        if isinstance(row, (list, tuple)):
            max_cols = max(max_cols, len(row))
        else:
            max_cols = max(max_cols, 1)

    normalized = []
    for row in data:
        if isinstance(row, (list, tuple)):
            current_row = list(row)
            if len(current_row) < max_cols:
                current_row.extend([""] * (max_cols - len(current_row)))
            normalized.append(current_row)
        else:
            normalized.append([row] + [""] * (max_cols - 1))

    return normalized


def _verify_against_source(
    worksheet,
    source_column: str,
    data: List[List[Any]],
    key_index: int,
    start_row: int,
    match_mode: str = "contains"
) -> Dict[str, Any]:
    """Verify that written data keys can be found in source column.

    Args:
        worksheet: Excel worksheet COM object
        source_column: Column letter to verify against (e.g., "A")
        data: 2D array of data being written
        key_index: Index of the column containing the key (e.g., 0 for first column)
        start_row: Starting row number for the data
        match_mode: "contains" | "exact" | "regex"

    Returns:
        Dictionary with verification results
    """
    if not data:
        return {"match_rate": 1.0, "mismatches": [], "sample_comparisons": []}

    num_rows = len(data)
    end_row = start_row + num_rows - 1

    # Read source column
    source_range = f"{source_column}{start_row}:{source_column}{end_row}"
    source_values = worksheet.Range(source_range).Value

    # Normalize to list
    if source_values is None:
        source_list = [None] * num_rows
    elif isinstance(source_values, (list, tuple)):
        source_list = [row[0] if isinstance(row, (list, tuple)) else row for row in source_values]
    else:
        source_list = [source_values]

    # Compare each row
    matches = 0
    mismatches = []
    sample_comparisons = []

    for i, row in enumerate(data):
        if i >= len(source_list):
            break

        source_val = str(source_list[i]) if source_list[i] else ""
        written_key = str(row[key_index]) if key_index < len(row) and row[key_index] else ""

        # Check for match
        is_match = False
        if match_mode == "exact":
            is_match = written_key == source_val
        elif match_mode == "contains":
            is_match = written_key in source_val or (written_key and written_key.replace(".0", "") in source_val)
        elif match_mode == "all_columns":
            # Layer 1: Holographic Check (Every output cell must be in source)
            # We use an alphanumeric-only check and strip '.0' for Excel floats
            def to_alphanum(s):
                if s is None:
                    return ""
                s_str = str(s)
                if s_str.endswith(".0"):
                    s_str = s_str[:-2]
                return "".join(c for c in s_str if c.isalnum()).lower()

            src_clean = to_alphanum(source_val)
            cell_mismatches = []

            for col_idx, cell_data in enumerate(row):
                if cell_data is None:
                    continue

                cell_val = str(cell_data)
                if not cell_val.strip():
                    continue

                # Check if cell content exists in source row (alphanumeric check)
                val_clean = to_alphanum(cell_val)
                if val_clean and val_clean not in src_clean:
                     cell_mismatches.append(cell_val)

            # Layer 2: Duplication Guard
            # Count how often this KEY appears in the LOCAL SOURCE vs LOCAL OUTPUT
            src_cnt = sum(1 for s in source_list if s and to_alphanum(written_key) in to_alphanum(s))
            out_cnt = sum(1 for r in data if to_alphanum(r[key_index]) == to_alphanum(written_key))

            is_dupe_safe = out_cnt <= src_cnt

            # Layer 3: Cross-Column Redundancy Guard
            # Checks if a significant value (length > 4) from one column appears as a substring in another column
            redundancy_issues = []
            non_key_values = []
            for col_idx, cell_data in enumerate(row):
                if col_idx == key_index or cell_data is None:
                    continue
                val_clean = to_alphanum(cell_data)
                if len(val_clean) > 4:  # Only check significant values
                    non_key_values.append((col_idx, val_clean, str(cell_data)))

            for v_idx_i, (col_i, val_i, orig_i) in enumerate(non_key_values):
                for v_idx_j, (col_j, val_j, orig_j) in enumerate(non_key_values):
                    if v_idx_i != v_idx_j and val_i in val_j and val_i != val_j:
                        # val_i is a substring of val_j (and they're not equal)
                        redundancy_issues.append(f"Col{col_i+1}:'{orig_i}' is repeated in Col{col_j+1}:'{orig_j}'")

            is_redundancy_safe = len(redundancy_issues) == 0
            is_match = (len(cell_mismatches) == 0) and is_dupe_safe and is_redundancy_safe

            if not is_match:
                diag = []
                if cell_mismatches:
                    diag.append(f"Content Mismatch: {cell_mismatches}")
                if not is_dupe_safe:
                    diag.append(f"Duplication Guard: out_cnt({out_cnt}) > src_cnt({src_cnt}) for key '{written_key}'")
                if redundancy_issues:
                    diag.append(f"Redundancy Guard: {redundancy_issues[0]}")  # Show first issue

        elif match_mode == "regex":
            try:
                is_match = bool(re.search(written_key, source_val))
            except re.error:
                is_match = written_key in source_val

        if is_match:
            matches += 1
        else:
            mismatch_detail = {
                "row": start_row + i,
                "source": source_val[:80] + "..." if len(source_val) > 80 else source_val,
                "written_key": written_key,
                "match": False
            }
            if match_mode == "all_columns" and 'diag' in locals():
                mismatch_detail["diagnostic"] = " | ".join(diag)

            mismatches.append(mismatch_detail)

        # Sample first 3 and last 2 rows
        if i < 3 or i >= num_rows - 2:
            sample_comparisons.append({
                "row": start_row + i,
                "source": source_val[:50] + "..." if len(source_val) > 50 else source_val,
                "written_key": written_key,
                "match": is_match
            })

    match_rate = matches / num_rows if num_rows > 0 else 1.0

    return {
        "match_rate": round(match_rate, 4),
        "total_rows": num_rows,
        "matches": matches,
        "mismatch_count": len(mismatches),
        "mismatches": mismatches[:10],  # Limit to first 10 mismatches
        "sample_comparisons": sample_comparisons,
    }

def write_cell_sync(
    workbook_name: str,
    sheet_name: str,
    cell: str,
    value: Any,
    force_overwrite: bool = False,
    activate: bool = True,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """Write to a single cell.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        cell: Cell reference
        value: Value to write
        force_overwrite: If True, overwrite existing data without checking
        activate: If True, activate the cell after writing
        dry_run: If True, validate without writing

    Returns:
        Dictionary with operation result
    """
    _init_com()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)
    rng = worksheet.Range(cell)

    # Sanitize value
    sanitized_value = _sanitize_value(value)

    # Caution mode check
    if not force_overwrite:
        old_val = rng.Value
        if old_val is not None and str(old_val).strip() != "":
            raise ToolError(
                f"Cell '{cell}' already contains data: '{old_val}'. "
                "Clear cell first or use force_overwrite=True."
            )

    if not dry_run:
        rng.Value = sanitized_value

        if activate:
            try:
                workbook.Activate()
                worksheet.Activate()
                rng.Select()
            except Exception:
                pass

    # Parse coordinates for output
    col_letter, row_num, _, _ = parse_range(cell)
    col_num = column_to_number(col_letter)

    # Audit log the write operation
    log_write(
        tool="write_cell",
        workbook=workbook_name,
        sheet=sheet_name,
        range_str=cell,
        cells=1,
        dry_run=dry_run,
    )

    return {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "cell": cell,
        "row": int(row_num),
        "col": col_num,
        "value": sanitized_value if not dry_run else None,
        "dry_run": dry_run,
    }


def write_range_sync(
    workbook_name: str,
    sheet_name: str,
    range_str: str,
    data: List[List[Any]],
    force_overwrite: bool = False,
    activate: bool = True,
    dry_run: bool = False,
    strict_alignment: bool = False,
    # NEW SAFETY PARAMETERS:
    max_cells: Optional[int] = None,
    verify_source: Optional[Dict[str, Any]] = None,
    abort_threshold: float = 0.0,
) -> Dict[str, Any]:
    """Write a 2D array of values to a range with safety verification.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range_str: Range reference
        data: 2D array of values
        force_overwrite: If True, overwrite existing data
        activate: If True, activate the range after writing
        dry_run: If True, validate without writing
        strict_alignment: If True, require exact dimension match
        max_cells: Maximum cells allowed (rows × cols, default: 100). Set to prevent LLM hallucination.
        verify_source: Source verification config:
            {"column": "A", "key_index": 0, "match_mode": "contains"}
        abort_threshold: Max allowed mismatch rate (default: 0.0 = 0%)
            If verification fails above this rate, write is aborted.

    Returns:
        Dictionary with operation result and verification details
    """
    _init_com()

    # Calculate actual cell count before limit check
    actual_rows = len(data) if data else 0
    actual_cols = len(data[0]) if data and data[0] else 0
    actual_cells = actual_rows * actual_cols

    # Apply max_cells limit
    effective_max_cells = max_cells if max_cells is not None else MAX_CELLS_LIMIT

    if actual_cells > effective_max_cells:
        raise ToolError(
            f"Data has {actual_cells} cells ({actual_rows} rows × {actual_cols} cols), "
            f"but max_cells limit is {effective_max_cells}. "
            f"This limit exists to prevent LLM hallucination. "
            f"Process data in smaller chunks or increase max_cells.",
            ErrorCodes.VALIDATION_ERROR
        )

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    # Sanitize and normalize data
    data = _sanitize_data(data)
    data = _normalize_jagged_array(data)

    # Parse range dimensions
    start_col, start_row, end_col, end_row = parse_range(range_str)

    actual_rows = len(data)
    actual_cols = len(data[0]) if data else 0

    s_row = int(start_row) if start_row else 1
    e_row = int(end_row) if end_row else s_row + actual_rows - 1
    s_col_num = column_to_number(start_col) if start_col else 1
    e_col_num = column_to_number(end_col) if end_col else s_col_num + actual_cols - 1

    expected_rows = e_row - s_row + 1
    expected_cols = e_col_num - s_col_num + 1

    # Strict alignment check
    if strict_alignment:
        if actual_rows != expected_rows or actual_cols != expected_cols:
            raise ToolError(
                f"Data dimensions ({actual_rows}x{actual_cols}) don't match "
                f"range dimensions ({expected_rows}x{expected_cols}). "
                "This may cause row-shifting issues."
            )

    # Calculate actual write range
    rows_to_write = min(actual_rows, expected_rows)
    cols_to_write = min(actual_cols, expected_cols)

    adjusted_end_col = number_to_column(s_col_num + cols_to_write - 1)
    adjusted_end_row = s_row + rows_to_write - 1
    adjusted_range = f"{start_col or 'A'}{s_row}:{adjusted_end_col}{adjusted_end_row}"

    # Trim data if needed
    write_data = [row[:cols_to_write] for row in data[:rows_to_write]]

    rng = worksheet.Range(adjusted_range)

    # Run source verification BEFORE writing
    verification = None
    if verify_source:
        source_col = verify_source.get("column", "A")
        key_idx = verify_source.get("key_index", 0)
        match_mode = verify_source.get("match_mode", "contains")

        verification = _verify_against_source(
            worksheet=worksheet,
            source_column=source_col,
            data=write_data,
            key_index=key_idx,
            start_row=s_row,
            match_mode=match_mode
        )

        # Check abort threshold
        mismatch_rate = 1.0 - verification["match_rate"]
        if mismatch_rate > abort_threshold:
            raise ToolError(
                f"VERIFICATION FAILED: {verification['mismatch_count']} of {verification['total_rows']} rows "
                f"({mismatch_rate:.1%}) failed source verification. Threshold is {abort_threshold:.1%}. "
                f"Write aborted to prevent hallucinated data. "
                f"First mismatch: {verification['mismatches'][0] if verification['mismatches'] else 'N/A'}",
                ErrorCodes.VALIDATION_ERROR
            )

    # Caution mode check
    if not force_overwrite and not dry_run:
        existing = rng.Value
        if existing:
            # Check if any existing data
            has_data = False
            if isinstance(existing, (list, tuple)):
                for row in existing:
                    if isinstance(row, (list, tuple)):
                        for cell in row:
                            if cell is not None and str(cell).strip() != "":
                                has_data = True
                                break
                    elif row is not None and str(row).strip() != "":
                        has_data = True
                        break
                    if has_data:
                        break
            elif existing is not None and str(existing).strip() != "":
                has_data = True

            if has_data:
                raise ToolError(
                    f"Range '{adjusted_range}' already contains data. "
                    "Clear range first or use force_overwrite=True."
                )

    if not dry_run:
        rng.Value = write_data

        if activate:
            try:
                workbook.Activate()
                worksheet.Activate()
                rng.Select()
            except Exception:
                pass

    result = {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "range": adjusted_range,
        "cells_written": rows_to_write * cols_to_write,
        "start_row": s_row,
        "end_row": adjusted_end_row,
        "start_col": s_col_num,
        "end_col": s_col_num + cols_to_write - 1,
        "data_preview": {
            "first_row": write_data[0] if write_data else [],
            "last_row": write_data[-1] if write_data else [],
            "total_rows": rows_to_write
        },
        "dry_run": dry_run,
    }

    # Add verification results if performed
    if verification:
        result["verification"] = verification
        if verification["mismatch_count"] > 0:
            result["warning"] = (
                f"{verification['mismatch_count']} rows had verification warnings. "
                "Review mismatches in verification.mismatches."
            )

    # Audit log the write operation
    log_write(
        tool="write_range",
        workbook=workbook_name,
        sheet=sheet_name,
        range_str=adjusted_range,
        cells=rows_to_write * cols_to_write,
        dry_run=dry_run,
    )

    return result
