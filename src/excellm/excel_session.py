"""Excel COM session manager for live operations.

This module handles the connection to running Excel instances and provides
methods for reading, writing, and managing Excel data via COM automation.
"""

import asyncio
import logging
from typing import Any, Dict, List, Optional, Tuple

import pythoncom
import win32com.client as win32


class ToolError(Exception):
    """Exception raised for tool-related errors."""


# Initialize COM for this module
pythoncom.CoInitialize()


from .validators import (
    validate_cell_format,
    validate_range_format,
    parse_reference_string,
    validate_sheet_name,
    validate_workbook_name,
    validate_value_type,
    get_cell_type,
    get_excel_error_info,
    parse_range,
    validate_data_dimensions,
)

from .filters import FilterEngine, FilterGroup, SingleFilter

logger = logging.getLogger(__name__)


class ExcelSessionManager:
    """Manages Excel COM session for live operations.

    This class connects to a running Excel application and provides
    methods for interacting with open workbooks and worksheets.
    """

    def __init__(self):
        """Initialize the Excel session manager."""
        self.excel_app = None
        self._connected = False

    async def _connect(self) -> None:
        """Connect to running Excel application (async wrapper)."""
        if self._connected:
            logger.debug("Already connected to Excel")
            return

        try:
            # Execute COM operation in thread to avoid blocking event loop
            self.excel_app = await asyncio.to_thread(self._connect_sync)
            self._connected = True
            logger.info("Connected to Excel application")
        except Exception as e:
            error_msg = str(e)[:200] if str(e) else "Unknown connection error"
            raise ToolError(
                f"Could not connect to Excel. Is Excel running? Error: {error_msg}"
            ) from e

    def _connect_sync(self) -> Any:
        """Synchronous COM connection (runs in thread)."""
        try:
            # Initialize COM in this thread
            pythoncom.CoInitialize()
            return win32.GetActiveObject("Excel.Application")
        except Exception as e:
            raise Exception(f"Failed to connect to Excel: {str(e)}")

    def _number_to_column(self, n: int) -> str:
        """Convert column number to Excel column letter (e.g., 1=A, 27=AA).

        Args:
            n: Column number (1-based)

        Returns:
            Excel column letter
        """
        result = ""
        while n > 0:
            n -= 1
            result = chr(n % 26 + ord('A')) + result
            n //= 26
        return result

    def _column_letter_to_number(self, letter: str) -> int:
        """Convert Excel column letter to number (e.g., A=1, AA=27).

        Args:
            letter: Column letter (e.g., "A", "B", "AA")

        Returns:
            Column number (1-based)
        """
        number = 0
        for char in letter:
            if "A" <= char <= "Z":
                number = number * 26 + (ord(char) - ord("A") + 1)
        return number

    def _normalize_jagged_array(self, data: List[List[Any]]) -> List[List[Any]]:
        """Normalize jagged array to rectangular by padding with empty values."""
        if not data:
            return []
        
        # Find maximum columns in any row
        max_cols = 0
        for row in data:
            if isinstance(row, (list, tuple)):
                max_cols = max(max_cols, len(row))
            else:
                max_cols = max(max_cols, 1)
        
        # Pad shorter rows
        normalized = []
        for row in data:
            if isinstance(row, (list, tuple)):
                current_row = list(row)
                if len(current_row) < max_cols:
                    current_row.extend([""] * (max_cols - len(current_row)))
                normalized.append(current_row)
            else:
                # Handle non-list row (shouldn't happen with 2D array input)
                normalized.append([row] + [""] * (max_cols - 1))
        
        return normalized

    def _sanitize_value(self, value: Any) -> Any:
        """Sanitize a single value for Excel (convert dicts/lists to JSON strings)."""
        if value is None:
            return ""
        if isinstance(value, (str, int, float, bool)):
            return value
        # Convert everything else (dict, list, etc.) to JSON string
        import json
        try:
            return json.dumps(value)
        except:
            return str(value)

    def _sanitize_data(self, data: Any) -> Any:
        """Sanitize data (single value or 2D array) for Excel."""
        if isinstance(data, list):
            sanitized = []
            for row in data:
                if isinstance(row, list):
                    sanitized.append([self._sanitize_value(v) for v in row])
                else:
                    sanitized.append(self._sanitize_value(row))
            return sanitized
        return self._sanitize_value(data)

    async def select_range(
        self, workbook_name: str, sheet_name: str, reference: str
    ) -> Dict[str, Any]:
        """Activate a workbook, worksheet, and select a range.
        
        This moves the focus in the Excel UI to the specified location.
        """
        await self._connect()
        return await asyncio.to_thread(self._select_range_sync, workbook_name, sheet_name, reference)

    def _select_range_sync(self, workbook_name: str, sheet_name: str, reference: str) -> Dict[str, Any]:
        """Synchronous range selection."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)
        
        # Activate workbook window if not active
        # (Using Activate on workbook might be needed if multiple Excel windows are open)
        try:
            workbook.Activate()
        except:
            pass # Ignore if fail
            
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

    async def list_workbooks(self) -> List[Dict[str, Any]]:
        """List all open workbooks.

        Returns:
            List of dictionaries containing workbook name and sheets
        """
        await self._connect()

        try:
            workbooks = await asyncio.to_thread(self._list_workbooks_sync)
            return workbooks
        except Exception as e:
            error_msg = str(e)[:200] if str(e) else "Unknown Excel error"
            raise ToolError(f"Failed to list workbooks: {error_msg}") from e

    def _list_workbooks_sync(self) -> List[Dict[str, any]]:
        """Synchronous workbook listing."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbooks = []

        try:
            wb_count = self.excel_app.Workbooks.Count
        except Exception:
            logger.warning("No open workbooks found")
            return workbooks

        for i in range(1, wb_count + 1):
            workbook = app.Workbooks(i)
            sheets = []

            try:
                for j in range(1, workbook.Worksheets.Count + 1):
                    worksheet = workbook.Worksheets(j)
                    sheet_name = worksheet.Name
                    # Check visibility: -1 = visible, 0 = hidden, 2 = very hidden
                    is_hidden = worksheet.Visible != -1
                    sheets.append({
                        "name": sheet_name,
                        "hidden": is_hidden
                    })
            except Exception as e:
                logger.warning(f"Could not list sheets: {str(e)[:100]}")

            try:
                workbooks.append({"name": workbook.Name, "sheets": sheets})
            except Exception:
                workbooks.append({"name": "Unknown", "sheets": sheets})

        return workbooks

    async def get_sheet_names(self, workbook_name: str) -> List[str]:
        """Get sheet names for a workbook.

        Args:
            workbook_name: Name of the workbook

        Returns:
            List of sheet names

        Raises:
            ToolError: If workbook not found
        """
        await self._connect()

        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")

        try:
            sheets = await asyncio.to_thread(self._get_sheet_names_sync, workbook_name)
            return sheets
        except Exception as e:
            raise ToolError(f"Failed to get sheet names: {str(e)}") from e

    def _get_sheet_names_sync(self, workbook_name: str) -> List[str]:
        """Synchronous sheet name retrieval."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        sheets = []

        for i in range(1, workbook.Worksheets.Count + 1):
            sheets.append(workbook.Worksheets(i).Name)

        return sheets

    async def read_cell(self, workbook_name: str, sheet_name: str, cell: str) -> Dict[str, Any]:
        """Read a single cell value.

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            cell: Cell reference (e.g., "A1", "B5")

        Returns:
            Dictionary with cell value and metadata

        Raises:
            ToolError: If any parameter is invalid or read fails
        """
        await self._connect()

        # Validate inputs
        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")

        if not validate_sheet_name(sheet_name):
            raise ToolError(f"Invalid sheet name: '{sheet_name}'")

        if not validate_cell_format(cell):
            raise ToolError(f"Invalid cell reference: '{cell}'. Expected format: A1, B5, Z100")

        try:
            result = await asyncio.to_thread(self._read_cell_sync, workbook_name, sheet_name, cell)
            return result
        except Exception as e:
            error_msg = str(e).lower()
            if "cannot get" in error_msg or "range" in error_msg:
                raise ToolError(
                    f"Invalid cell reference '{cell}' in workbook '{workbook_name}', "
                    f"sheet '{sheet_name}'"
                ) from e
            raise ToolError(f"Failed to read cell: {str(e)}") from e

    def _read_cell_sync(self, workbook_name: str, sheet_name: str, cell: str) -> Dict[str, Any]:
        """Synchronous cell read."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)

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

        # Always add formula field
        try:
            formula = rng.Formula
            # Check if it starts with "=" to determine if it's actually a formula
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

    async def read_range(
        self, workbook_name: str, sheet_name: str, range_str: Optional[str] = None
    ) -> Dict[str, Any]:
        """Read a range of cells.

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            range_str: Range reference (e.g., "A1:C5", "B2:D10")

        Returns:
            Dictionary with range data and metadata

        Raises:
            ToolError: If any parameter is invalid or read fails
        """
        await self._connect()

        # Validate inputs
        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")

        if not validate_sheet_name(sheet_name):
            raise ToolError(f"Invalid sheet name: '{sheet_name}'")

        if range_str and not validate_range_format(range_str):
            raise ToolError(f"Invalid range format: '{range_str}'. Expected format: A1:C5, B2:D10")

        try:
            result = await asyncio.to_thread(
                self._read_range_sync, workbook_name, sheet_name, range_str
            )
            return result
        except Exception as e:
            raise ToolError(f"Failed to read range: {str(e)}") from e

    def _read_range_sync(
        self, workbook_name: str, sheet_name: str, range_str: str
    ) -> Dict[str, Any]:
        """Synchronous range read."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)

        # Get range reference with optimization for large ranges
        if range_str:
            target_rng = worksheet.Range(range_str)
            
            # Optimization: If it's a whole row or column, intersect with UsedRange
            # parse_range returns empty strings for missing components
            start_col, start_row, end_col, end_row = parse_range(range_str)
            is_whole_col = not start_row and not end_row
            is_whole_row = not start_col and not end_col
            
            if is_whole_col or is_whole_row:
                used_rng = worksheet.UsedRange
                # Intersect target range with UsedRange
                intersect_rng = self.excel_app.Intersect(target_rng, used_rng)
                
                if intersect_rng:
                    # Further Optimization: If intersection is large, shrink to actual data
                    # (UsedRange can be bloated by formatting or deleted data)
                    try:
                        # -4163 = xlValues, 1 = xlFormulas
                        # Search backward to find the last cell with actual content/formula
                        # LookAt=2 (xlPart) with What="*" is very reliable for finding any non-empty cell
                        # Specify After=first_cell and SearchDirection=2 (xlPrevious) to wrap to bottom-right
                        first_cell = intersect_rng.Cells(1, 1)
                        s_row = intersect_rng.Row
                        s_col = intersect_rng.Column
                        
                        # Try Values first (usually more consistent for what the user sees)
                        last_row_cell = intersect_rng.Find("*", After=first_cell, LookIn=-4163, LookAt=2, SearchOrder=1, SearchDirection=2)
                        last_col_cell = intersect_rng.Find("*", After=first_cell, LookIn=-4163, LookAt=2, SearchOrder=2, SearchDirection=2)
                        
                        # Fallback to Formulas if Values didn't find much (e.g. only formulas in cells)
                        if not last_row_cell:
                            last_row_cell = intersect_rng.Find("*", After=first_cell, LookIn=1, LookAt=2, SearchOrder=1, SearchDirection=2)
                        if not last_col_cell:
                            last_col_cell = intersect_rng.Find("*", After=first_cell, LookIn=1, LookAt=2, SearchOrder=2, SearchDirection=2)
                        
                        if last_row_cell and last_col_cell:
                            # Recalculate range based on last data row/col
                            if is_whole_col:
                                # Keep columns, limit rows to last_row_cell.Row
                                e_row = last_row_cell.Row
                                e_col = s_col + intersect_rng.Columns.Count - 1
                                intersect_rng = worksheet.Range(
                                    worksheet.Cells(s_row, s_col),
                                    worksheet.Cells(e_row, e_col)
                                )
                            else:
                                # Keep rows, limit columns to last_col_cell.Column
                                e_row = s_row + intersect_rng.Rows.Count - 1
                                e_col = last_col_cell.Column
                                intersect_rng = worksheet.Range(
                                    worksheet.Cells(s_row, s_col),
                                    worksheet.Cells(e_row, e_col)
                                )
                    except Exception as e:
                        logger.debug(f"Deep search optimization failed: {str(e)}")
                    
                    rng = intersect_rng
                    range_str = rng.Address.replace("$", "")
                else:
                    # No intersection means the requested range is empty
                    # We still need a valid range object for metadata
                    # Use a 1x1 range at the start of the request
                    s_row = int(start_row) if start_row else 1
                    s_col = self._column_letter_to_number(start_col) if start_col else 1
                    rng = worksheet.Cells(s_row, s_col)
                    # Force values to None because it might not be empty if UsedRange is elsewhere
                    # but Intersect already proved it's empty relative to whole row/col.
                    # Fast exit with empty result below.
            else:
                rng = target_rng
        else:
            rng = worksheet.UsedRange
            range_str = rng.Address.replace("$", "")

        # Get range values
        values = rng.Value

        # Convert to 2D list with error information
        data = []
        error_codes = []
        error_messages = []
        if values is not None:
            # Check if values is a list/tuple (range values) or scalar (single value)
            if isinstance(values, (list, tuple)):
                # Check if values is 2D (multiple rows) - Excel returns tuples, not lists
                is_2d = len(values) > 0 and isinstance(values[0], (list, tuple))

                if is_2d:
                    # Range spans multiple rows
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

        # Always add formulas field
        try:
            formulas = rng.Formula
            formula_data = []
            if formulas:
                if isinstance(formulas, (list, tuple)):
                    if len(formulas) > 0 and isinstance(formulas[0], (list, tuple)):
                        # 2D array
                        for row in formulas:
                            row_formulas = []
                            for f in row:
                                if isinstance(f, str) and f.startswith("="):
                                    row_formulas.append(f)
                                else:
                                    row_formulas.append(None)
                            formula_data.append(row_formulas)
                    else:
                        # 1D array
                        row_formulas = []
                        for f in formulas:
                            if isinstance(f, str) and f.startswith("="):
                                row_formulas.append(f)
                            else:
                                row_formulas.append(None)
                        formula_data.append(row_formulas)
                else:
                    # Single value
                    if isinstance(formulas, str) and formulas.startswith("="):
                        formula_data.append([formulas])
                    else:
                        formula_data.append([None])

            result["formulas"] = formula_data if formula_data else None
        except Exception:
            result["formulas"] = None

        # Only include error fields if there are errors
        has_errors = any(any(cell) for row in error_codes for cell in row if cell)
        if has_errors:
            result["error_codes"] = error_codes
            result["error_messages"] = error_messages

        return result

    async def write_cell(
        self,
        workbook_name: str,
        sheet_name: str,
        cell: str,
        value: Any,
        force_overwrite: bool = False,
        activate: bool = True,
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        """Write to a single cell."""
        await self._connect()
        
        # Sanitize value
        sanitized_value = self._sanitize_value(value)
        
        if dry_run:
            # For dry run, we still need to validate existence and overwrite constraints
            # but we use a sync wrapper that doesn't actually write
            return await asyncio.to_thread(
                self._write_cell_sync, workbook_name, sheet_name, cell, sanitized_value, force_overwrite, activate, dry_run=True
            )

        return await asyncio.to_thread(
            self._write_cell_sync, workbook_name, sheet_name, cell, sanitized_value, force_overwrite, activate
        )

    def _write_cell_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        cell: str,
        value: Any,
        force_overwrite: bool = False,
        activate: bool = True,
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        """Synchronous cell write."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)
        rng = worksheet.Range(cell)

        # Caution mode check
        if not force_overwrite:
            old_val = rng.Value
            if old_val is not None and str(old_val).strip() != "":
                raise ToolError(
                    f"Cell '{cell}' already contains data: '{old_val}'. "
                    "Clear cell first or use force_overwrite=True."
                )

        if not dry_run:
            worksheet.Range(cell).Value = value
            
            if activate:
                try:
                    workbook.Activate()
                    worksheet.Activate()
                    worksheet.Range(cell).Select()
                except:
                    pass

        # Parse coordinates for output
        col_letter, row_num, _, _ = parse_range(cell)
        col_num = self._column_letter_to_number(col_letter)

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "cell": cell,
            "row": int(row_num),
            "col": col_num,
            "value": value if not dry_run else None,
            "dry_run": dry_run,
        }

    async def write_range(
        self,
        workbook_name: str,
        sheet_name: str,
        range_str: str,
        data: List[List[Any]],
        force_overwrite: bool = False,
        activate: bool = True,
        dry_run: bool = False,
        strict_alignment: bool = False,
    ) -> Dict[str, Any]:
        """Write a 2D array of values to a range.

        Automatically normalizes jagged arrays (rows with different column counts)
        by padding shorter rows with empty values to match the maximum column count.

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            range_str: Range reference (e.g., "A1:C5")
            data: 2D array of values (jagged arrays are auto-normalized)
            force_overwrite: If True, bypass caution mode and overwrite existing data
            activate: If True, activate the workbook, sheet, and select the range after writing.
            dry_run: If True, validate without writing.

        Returns:
            Dictionary with operation result

        Raises:
            ToolError: If any parameter is invalid or write fails

        Example:
            # Jagged array is auto-padded:
            data = [["a", "b", "c"], ["d", "e"],  # Row 2 gets padded to ["d", "e", ""]
                   ["f", "g", "h"]]
            write_range("book.xlsx", "Sheet1", "A1:C3", data)
        """
        await self._connect()

        # Validate inputs
        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")

        if not validate_sheet_name(sheet_name):
            raise ToolError(f"Invalid sheet name: '{sheet_name}'")

        if not validate_range_format(range_str):
            raise ToolError(f"Invalid range format '{range_str}'. Expected format: A1:C5, B2:D10")

        if not data or not isinstance(data, list):
            raise ToolError("Data must be a 2D array (list of lists)")

        # Santize data (recursively convert non-Excel types to JSON strings)
        data = self._sanitize_data(data)

        # Normalize jagged array to rectangular (pad shorter rows with empty strings)
        if data:
            max_cols = 0
            for row in data:
                if isinstance(row, (list, tuple)):
                    max_cols = max(max_cols, len(row))
                else:
                    max_cols = max(max_cols, 1)
            
            normalized_data = []
            for row in data:
                if isinstance(row, (list, tuple)):
                    current_row = list(row)
                    if len(current_row) < max_cols:
                        current_row.extend([""] * (max_cols - len(current_row)))
                    normalized_data.append(current_row)
                else:
                    normalized_data.append([row] + [""] * (max_cols - 1))
            data = normalized_data

        # Parse range to get dimensions
        start_col, start_row, end_col, end_row = parse_range(range_str)

        actual_rows = len(data)
        actual_cols = len(data[0]) if data else 0

        # Handle whole row/column by using Excel limits as defaults
        # Max row: 1,048,576
        # Max col: 16,384 (XFD)
        s_row = int(start_row) if start_row else 1
        e_row = int(end_row) if end_row else 1048576
        s_col_num = self._column_letter_to_number(start_col) if start_col else 1
        e_col_num = self._column_letter_to_number(end_col) if end_col else 16384

        expected_rows = e_row - s_row + 1
        expected_cols = e_col_num - s_col_num + 1

        # Calculate the actual range that will be written (partial write support)
        rows_to_write = min(actual_rows, expected_rows)
        cols_to_write = min(actual_cols, expected_cols)

        # Calculate adjusted end cell
        final_start_col = start_col if start_col else "A"
        final_start_row = str(s_row)
        adjusted_end_col = self._number_to_column(s_col_num + cols_to_write - 1)
        adjusted_end_row = s_row + rows_to_write - 1
        adjusted_range = f"{final_start_col}{final_start_row}:{adjusted_end_col}{adjusted_end_row}"

        # Trim data if it's larger than the range
        write_data = data
        data_was_trimmed = False
        if actual_rows > rows_to_write or actual_cols > cols_to_write:
            write_data = [row[:cols_to_write] for row in data[:rows_to_write]]
            data_was_trimmed = True

        # Check if range was adjusted (data smaller than range)
        range_adjusted = adjusted_range != range_str


        try:
            result = await asyncio.to_thread(
                self._write_range_sync,
                workbook_name,
                sheet_name,
                range_str, # Use ORG range_str for strict check
                data,
                force_overwrite,
                activate,
                dry_run,
                strict_alignment,
            )

            # Add warning if data was trimmed or range was adjusted
            if data_was_trimmed or range_adjusted:
                if data_was_trimmed and not range_adjusted:
                    # Data was larger than range
                    trimmed_rows = actual_rows - rows_to_write
                    trimmed_cols = actual_cols - cols_to_write
                    if trimmed_rows > 0 and trimmed_cols > 0:
                        warning = f"Data had {actual_rows} rows x {actual_cols} cols, range fits {rows_to_write} rows x {cols_to_write} cols - {trimmed_rows} row(s) and {trimmed_cols} col(s) truncated"
                    elif trimmed_rows > 0:
                        warning = f"Data had {actual_rows} rows, range fits {rows_to_write} rows - {trimmed_rows} row(s) truncated"
                    elif trimmed_cols > 0:
                        warning = f"Data had {actual_cols} columns, range fits {cols_to_write} columns - {trimmed_cols} column(s) truncated"
                elif range_adjusted and not data_was_trimmed:
                    # Data was smaller than range
                    if rows_to_write < expected_rows and cols_to_write < expected_cols:
                        warning = f"Wrote {rows_to_write} of {expected_rows} rows, {cols_to_write} of {expected_cols} columns"
                    elif rows_to_write < expected_rows:
                        unused_rows = expected_rows - rows_to_write
                        warning = f"Wrote {rows_to_write} of {expected_rows} rows - {unused_rows} row(s) unused"
                    elif cols_to_write < expected_cols:
                        unused_cols = expected_cols - cols_to_write
                        warning = f"Wrote {cols_to_write} of {expected_cols} columns - {unused_cols} column(s) unused"
                else:
                    # Both trimmed and adjusted (detailed warning)
                    trimmed_rows = actual_rows - rows_to_write
                    trimmed_cols = actual_cols - cols_to_write
                    unused_rows = expected_rows - rows_to_write
                    unused_cols = expected_cols - cols_to_write

                    parts = []
                    if trimmed_rows > 0:
                        parts.append(f"{trimmed_rows} row(s) truncated")
                    if unused_rows > 0:
                        parts.append(f"{unused_rows} row(s) unused")
                    if trimmed_cols > 0:
                        parts.append(f"{trimmed_cols} column(s) truncated")
                    if unused_cols > 0:
                        parts.append(f"{unused_cols} column(s) unused")

                    warning = f"Data ({actual_rows}x{actual_cols}) written to {adjusted_range} ({rows_to_write}x{cols_to_write}): {', '.join(parts)}"

                result["warning"] = warning
                result["requested_range"] = range_str
                result["actual_range"] = adjusted_range

            return result
        except Exception as e:
            error_msg = str(e).lower()
            if "protected" in error_msg:
                raise ToolError(
                    f"Worksheet '{sheet_name}' is protected. "
                    "Please unprotect in worksheet in Excel first (Review → Unprotect Sheet)."
                )
            elif "read-only" in error_msg:
                raise ToolError(
                    f"Workbook '{workbook_name}' is read-only. Please open with write permissions."
                )
            raise ToolError(f"Failed to write range: {str(e)}") from e

    def _write_range_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        range_str: str,
        data: List[List[Any]],
        force_overwrite: bool = False,
        activate: bool = True,
        dry_run: bool = False,
        strict_alignment: bool = False,
    ) -> Dict[str, Any]:
        """Synchronous range write."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)
        target_rng = worksheet.Range(range_str)

        # Parse range for dimensions
        start_col, start_row, end_col, end_row = parse_range(range_str)
        
        # Calculate expected dimensions from range
        # (Handling whole row/column cases)
        exp_rows = target_rng.Rows.Count
        exp_cols = target_rng.Columns.Count
        
        # Calculate actual dimensions from data
        data_rows = len(data)
        data_cols = len(data[0]) if data and data[0] else 0

        # Strict Alignment Check
        if strict_alignment:
            if data_rows != exp_rows or data_cols != exp_cols:
                raise ToolError(
                    f"STRICT ALIGNMENT ERROR: Data dimensions ({data_rows}x{data_cols}) "
                    f"do not match Excel range '{range_str}' ({exp_rows}x{exp_cols}). "
                    "Update your data or range to avoid row-shifting accidents."
                )

        # Check if any cells already have data (caution mode - skip if force_overwrite)
        if not force_overwrite:
            cells_with_data = []  # Track which cells have data
            try:
                # Optimization for caution mode: only check the area we are writing to
                start_col, start_row, end_col, end_row = parse_range(range_str)
                s_row = int(start_row) if start_row else 1
                s_col_num = self._column_letter_to_number(start_col) if start_col else 1
                
                # Calculate required dimensions from data
                data_rows = len(data)
                data_cols = len(data[0]) if data and data[0] else 0
                
                if data_rows == 0 or data_cols == 0:
                     return {"success": True, "message": "No data to write"}

                # Construct sub-range for check
                e_row = s_row + data_rows - 1
                e_col_num = s_col_num + data_cols - 1
                e_col_letter = self._number_to_column(e_col_num)
                check_range_str = f"{self._number_to_column(s_col_num)}{s_row}:{e_col_letter}{e_row}"
                
                range_values = worksheet.Range(check_range_str).Value
                if range_values is not None:

                    # Check if range_values is a 2D structure (multiple rows)
                    # Excel returns tuples, so check for both list and tuple
                    is_2d = isinstance(range_values, (list, tuple)) and len(range_values) > 0 and isinstance(range_values[0], (list, tuple))

                    if is_2d:
                        # Multiple rows - Excel returns ((row1), (row2), ...)
                        for i, row in enumerate(range_values):
                            for j, val in enumerate(row):
                                # Check for actual data (not None, not empty string)
                                if val is not None and val != "" and str(val).strip() != "":
                                    # Calculate cell reference
                                    col_letter = self._number_to_column(s_col_num + j)
                                    row_num = s_row + i
                                    cell_ref = f"{col_letter}{row_num}"
                                    cells_with_data.append(f"{cell_ref}='{val}'")
                    else:
                        # Single row - Excel returns (val1, val2, ...)
                        for j, val in enumerate(range_values):
                            if val is not None and val != "" and str(val).strip() != "":
                                col_letter = self._number_to_column(s_col_num + j)
                                cell_ref = f"{col_letter}{s_row}"
                                cells_with_data.append(f"{cell_ref}='{val}'")

                    # Only raise error if we actually found data
                    if cells_with_data:
                        cells_list = ", ".join(cells_with_data[:10])  # Show first 10
                        if len(cells_with_data) > 10:
                            cells_list += f" ... and {len(cells_with_data) - 10} more"
                        raise ToolError(
                            f"Range '{range_str}' already contains data in {len(cells_with_data)} cell(s): {cells_list}. "
                            "Will not overwrite. Clear the cells first or use force_overwrite=True."
                        )
            except ToolError:
                # Re-raise ToolError (data exists, don't overwrite)
                raise
            except Exception:
                # Continue write if we can't read (might be empty range)
                pass

        if not dry_run:
            # Perform write
            worksheet.Range(range_str).Value = data
            
            if activate:
                try:
                    workbook.Activate()
                    worksheet.Activate()
                    worksheet.Range(range_str).Select()
                except:
                    pass

        rows_written = len(data)
        cols_written = len(data[0]) if data else 0

        # Calculate absolute coordinates
        s_col_letter, s_row_num, e_col_letter, e_row_num = parse_range(range_str)
        s_col_num = self._column_letter_to_number(s_col_letter)
        
        # If e_col_letter or e_row_num are missing (whole row/col), calculate them
        if not e_row_num:
            e_row_num = int(s_row_num) + rows_written - 1
        if not e_col_letter:
            e_col_num = s_col_num + cols_written - 1
            e_col_letter = self._number_to_column(e_col_num)
        else:
            e_col_num = self._column_letter_to_number(e_col_letter)

        result = {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "range": range_str,
            "start_row": int(s_row_num),
            "end_row": int(e_row_num),
            "start_col": s_col_num,
            "end_col": e_col_num,
            "cells_written": rows_written * cols_written if not dry_run else 0,
            "dry_run": dry_run,
        }

        if rows_written > 0:
            result["data_preview"] = {
                "first_row": data[0],
                "last_row": data[-1],
                "total_rows": rows_written
            }

        return result

    async def get_workbook_info(self, workbook_name: str) -> Dict[str, Any]:
        """Get detailed information about a workbook.

        Args:
            workbook_name: Name of open workbook

        Returns:
            Dictionary with workbook metadata

        Raises:
            ToolError: If workbook not found
        """
        await self._connect()

        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")

        try:
            info = await asyncio.to_thread(self._get_workbook_info_sync, workbook_name)
            return info
        except Exception as e:
            raise ToolError(f"Failed to get workbook info: {str(e)}") from e

    def _get_workbook_info_sync(self, workbook_name: str) -> Dict[str, Any]:
        """Synchronous workbook info retrieval."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)

        # Get workbook properties
        path = str(workbook.Path) if workbook.Path else ""
        full_path = str(workbook.FullName) if workbook.FullName else ""
        protected = workbook.ProtectStructure
        read_only = workbook.ReadOnly

        # Get sheet names
        sheets = []
        for i in range(1, workbook.Worksheets.Count + 1):
            sheets.append(workbook.Worksheets(i).Name)

        return {
            "success": True,
            "name": workbook_name,
            "path": path,
            "full_path": full_path,
            "sheet_count": len(sheets),
            "sheets": sheets,
            "protected": bool(protected),
            "read_only": bool(read_only),
        }

    async def read(
        self, workbook_name: str, sheet_name: str, reference: Optional[str] = None
    ) -> Dict[str, Any]:
        """Read from a cell or range - auto-detects based on reference format.

        Supports comma-separated references (e.g., "A1,C3,E5" or "A1:B2,D4:E6").

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            reference: Cell reference (A1), range (A1:C5), or comma-separated references

        Returns:
            Dictionary with read result, including 'scattered' flag for multi-reference
        """
        # Default to UsedRange if no reference provided
        if reference is None:
            return await self.read_range(workbook_name, sheet_name, None)

        # Check if reference contains commas (multiple references)
        if "," in reference:
            # Parse comma-separated references
            refs = parse_reference_string(reference)

            # Read each reference and collect results
            results = []
            errors = []

            for ref in refs:
                try:
                    if validate_cell_format(ref):
                        result = await self.read_cell(workbook_name, sheet_name, ref)
                        results.append(result)
                    elif validate_range_format(ref):
                        result = await self.read_range(workbook_name, sheet_name, ref)
                        results.append(result)
                    else:
                        errors.append(f"Invalid reference '{ref}'")
                except Exception as e:
                    errors.append(f"Failed to read '{ref}': {str(e)}")

            if errors:
                raise ToolError(f"Errors reading scattered references: {'; '.join(errors)}")

            # Return aggregated results
            return {
                "success": True,
                "workbook": workbook_name,
                "sheet": sheet_name,
                "scattered": True,
                "reference": reference,
                "results": results,
                "count": len(results),
            }
        else:
            # Single reference (backward compatible)
            if validate_cell_format(reference):
                return await self.read_cell(workbook_name, sheet_name, reference)
            elif validate_range_format(reference):
                return await self.read_range(workbook_name, sheet_name, reference)
            else:
                raise ToolError(
                    f"Invalid reference '{reference}'. Use cell format (A1) or range format (A1:C5)"
                )

    async def search(
        self,
        workbook_name: str,
        sheet_name: Optional[str] = None,
        range_str: Optional[str] = None,
        filters: dict = None,
        has_header: bool = True,
        all_sheets: bool = False,
    ) -> Dict[str, Any]:
        """Search Excel data and apply filters in memory.

        This method reads data from Excel using read_range(), then applies
        filters to return only matching rows. Does NOT modify the Excel file.

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            range_str: Range reference (e.g., "A1:D1000")
            filters: Filter specification dict with nested AND/OR/NOT logic.
                     Can be a SingleFilter dict or FilterGroup dict.
            has_header: Whether first row contains column headers (default: True)
            all_sheets: If True, searches all sheets in the workbook.
                        If True, sheet_name is ignored.

        Returns:
            Dictionary with:
            - success: True if operation succeeded
            - workbook: Workbook name
            - sheet: Sheet name
            - range: Original range
            - data: Filtered 2D array
            - rows_filtered: Number of rows after filtering
            - rows_removed: Number of rows removed by filter
            - columns: Column names (if headers present)
            - filter_applied: Human-readable filter description
            - warnings: List of warning messages (if any issues found)

        Raises:
            ToolError: If workbook, sheet, or range is invalid

        Example:
            >>> await search(
            ...     "data.xlsx", "Sheet1", "A1:D1000",
            ...     filters={
            ...         "operator": "and",
            ...         "conditions": [
            ...             {"column": {"type": "name", "value": "Status"},
            ...              "operator": "=", "value": "Active"},
            ...             {"column": {"type": "name", "value": "Balance"},
            ...              "operator": ">", "value": 500}
            ...         ]
            ...     },
            ...     case_sensitive=True
            ... )
        """
        await self._connect()

        # Validate workbook
        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")

        # Determine sheets to search
        sheets_to_search = []
        if all_sheets:
            try:
                wb_info = await self.get_workbook_info(workbook_name)
                # Only search visible sheets (standard behavior for global search)
                # Use list comprehension since sheets might be dicts or strings depending on version
                sheets_to_search = [s["name"] if isinstance(s, dict) else s for s in wb_info["sheets"]]
            except Exception as e:
                raise ToolError(f"Failed to get sheets for workbook search: {str(e)}")
        else:
            if not sheet_name:
                raise ToolError("sheet_name is required when all_sheets is False")
            if not validate_sheet_name(sheet_name):
                raise ToolError(f"Invalid sheet name: '{sheet_name}'")
            sheets_to_search = [sheet_name]

        if range_str and not validate_range_format(range_str):
            raise ToolError(f"Invalid range format: '{range_str}'. Expected format: A1:C5, B2:D10")

        # Convert simple search terms (string, number, etc.) to a filter dict
        if not isinstance(filters, dict):
            if filters is None:
                raise ToolError("Filters must be a non-empty dictionary or search term")
            filters = {"value": filters}

        all_results = []
        total_rows_filtered = 0
        total_rows_removed = 0

        for current_sheet in sheets_to_search:
            try:
                # Read the data first using existing read_range
                # Note: read_range handles theUsedRange if range_str is None
                read_result = await self.read_range(workbook_name, current_sheet, range_str)

                if not read_result.get("success"):
                    if not all_sheets:
                        return {
                            "success": False,
                            "error": f"Failed to read data from sheet '{current_sheet}'",
                            "workbook": workbook_name,
                            "sheet": current_sheet,
                        }
                    continue # Skip failed sheets in global search

                # Extract data and headers
                data = read_result.get("data", [])
                if not data:
                    continue

                # Apply headers if present
                headers = []
                if has_header and data:
                    headers = data[0]
                    content_rows = data[1:]
                else:
                    content_rows = data

                # Parse filters using existing FilterEngine
                engine = FilterEngine()
                start_row = read_result.get("start_row", 1)
                start_col = read_result.get("start_col", 1)
                
                # Apply filter returns {data, rows_filtered, rows_removed, columns, filter_applied, cell_locations}
                filter_result = engine.apply_filter(
                    data, filters, headers if has_header else None, 
                    start_row=start_row, start_col=start_col
                )

                filtered_data = filter_result["data"]
                match_count = filter_result["rows_filtered"]
                
                # In all_sheets mode, only include sheets that have actual matches
                if match_count > 0 or not all_sheets:
                    # Formatted result for this sheet
                    sheet_result = {
                        "sheet": current_sheet,
                        "range": read_result.get("range"),
                        "data": filtered_data,  # filter_result["data"] already includes headers once
                        "rows_filtered": match_count,
                        "rows_removed": filter_result["rows_removed"],
                        "columns": filter_result.get("columns", []),
                        "cell_locations": filter_result.get("cell_locations", []),
                    }
                    all_results.append(sheet_result)
                    total_rows_filtered += match_count
                    total_rows_removed += filter_result["rows_removed"]

            except Exception as e:
                # Log warning and skip for global search, raise for single sheet
                if not all_sheets:
                    raise ToolError(f"Search failed on sheet '{current_sheet}': {str(e)}")
                logger.warning(f"Skipping sheet '{current_sheet}' in global search due to error: {str(e)}")

        if all_sheets:
            return {
                "success": True,
                "workbook": workbook_name,
                "all_sheets": True,
                "sheets_searched": len(sheets_to_search),
                "total_rows_filtered": total_rows_filtered,
                "total_rows_removed": total_rows_removed,
                "results": all_results,
                "filter_applied": str(filters),
            }
        else:
            if not all_results:
                 return {
                    "success": True,
                    "workbook": workbook_name,
                    "sheet": sheet_name,
                    "data": [],
                    "rows_filtered": 0,
                    "rows_removed": 0,
                    "columns": [],
                    "filter_applied": str(filters),
                }
            # Single sheet result - return flattened
            res = all_results[0]
            return {
                "success": True,
                "workbook": workbook_name,
                "sheet": res["sheet"],
                "range": res["range"],
                "data": res["data"],
                "rows_filtered": res["rows_filtered"],
                "rows_removed": res["rows_removed"],
                "columns": res["columns"],
                "cell_locations": res["cell_locations"],
                "filter_applied": str(filters),
            }

    async def write(
        self,
        workbook_name: str,
        sheet_name: str,
        reference: str,
        data: Any,
        force_overwrite: bool = False,
        activate: bool = True,
        dry_run: bool = False,
        strict_alignment: bool = False,
    ) -> Dict[str, Any]:
        """Write to a cell or range - auto-detects based on reference format.

        Supports comma-separated references with data array (e.g., "A1,C3,E5" with ["val1", "val2", "val3"]).

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            reference: Cell reference (A1), range (A1:C5), or comma-separated references
            data: Value (for cell), 2D array (for range), or list (for scattered references)
            force_overwrite: If True, bypass caution mode and overwrite existing data
            dry_run: If True, validate without writing.

        Returns:
            Dictionary with write result, including 'scattered' flag for multi-reference
        """
        # Check if reference contains commas (multiple references)
        if "," in reference:
            # Parse comma-separated references
            refs = parse_reference_string(reference)

            # Validate data is a list for scattered writes
            if not isinstance(data, list):
                raise ToolError(
                    "Data must be a list/array for scattered references. "
                    f"Received {len(refs)} references but data is not a list."
                )

            # Validate data array length matches reference count
            if len(data) != len(refs):
                raise ToolError(
                    f"Data array length ({len(data)}) must match reference count ({len(refs)}). "
                    f"References: {reference}"
                )

            # Write to each reference with corresponding data
            results = []
            errors = []

            for i, ref in enumerate(refs):
                try:
                    value = data[i]

                    if validate_cell_format(ref):
                        if not validate_value_type(value):
                            errors.append(
                                f"Reference '{ref}': Invalid value type {type(value).__name__}"
                            )
                            continue
                        result = await self.write_cell(workbook_name, sheet_name, ref, value, force_overwrite, activate=activate, dry_run=dry_run)
                        results.append(result)
                    elif validate_range_format(ref):
                        if not value or not isinstance(value, list):
                            errors.append(
                                f"Reference '{ref}': Range requires 2D array, got {type(value).__name__}"
                            )
                            continue
                        result = await self.write_range(workbook_name, sheet_name, ref, value, force_overwrite, activate=activate, dry_run=dry_run, strict_alignment=strict_alignment)
                        results.append(result)
                    else:
                        errors.append(f"Invalid reference '{ref}'")
                except Exception as e:
                    errors.append(f"Failed to write '{ref}': {str(e)}")

            if errors:
                raise ToolError(f"Errors writing scattered references: {'; '.join(errors)}")

            # Return aggregated results
            return {
                "success": True,
                "workbook": workbook_name,
                "sheet": sheet_name,
                "scattered": True,
                "reference": reference,
                "results": results,
                "count": len(results),
            }
        else:
            # Single reference (backward compatible)
            if validate_cell_format(reference):
                if not validate_value_type(data):
                    raise ToolError(
                        f"Invalid value type: {type(data).__name__}. "
                        "Excel accepts: strings, numbers, booleans"
                    )
                return await self.write_cell(workbook_name, sheet_name, reference, data, force_overwrite, activate=activate, dry_run=dry_run)
            elif validate_range_format(reference):
                if not data or not isinstance(data, list):
                    raise ToolError("Data must be a 2D array (list of lists) for ranges")
                return await self.write_range(workbook_name, sheet_name, reference, data, force_overwrite, activate=activate, dry_run=dry_run, strict_alignment=strict_alignment)
            else:
                raise ToolError(
                    f"Invalid reference '{reference}'. Use cell format (A1) or range format (A1:C5)"
                )

    async def manage_sheet(
        self,
        workbook_name: str,
        sheet_name: str = None,
        sheet_names: list = None,
        action: str = None,
        force_delete: bool = False,
        target_workbook: str = None,
        target_name: str = None,
        position: str = None,
        reference_sheet: str = None,
    ) -> Dict[str, Any]:
        """Add, remove, hide, unhide, copy, or move worksheets in a workbook.

        Args:
            workbook_name: Name of open workbook
            sheet_name: Single sheet name (for add/remove/single operation)
            sheet_names: Multiple sheet names (for bulk hide/unhide)
            action: "add", "remove", "hide", "unhide", "copy", or "move"
            force_delete: If True, bypass empty check when removing sheets
            target_workbook: For copy/move, target workbook name (default: same as source)
            target_name: For copy/move, new name for the sheet
            position: For copy/move, "before" or "after" for positioning
            reference_sheet: For copy/move, sheet to position before/after

        Returns:
            Dictionary with operation result(s)
        """
        await self._connect()

        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")

        if action not in ("add", "remove", "hide", "unhide", "copy", "move", "rename"):
            raise ToolError(
                f"Invalid action '{action}'. Must be 'add', 'remove', 'hide', 'unhide', 'copy', 'move', or 'rename'"
            )

        # Validate sheet name(s)
        if action in ("add", "remove", "copy", "move", "rename"):
            # add/remove/copy/move require single sheet_name
            if not sheet_name:
                raise ToolError(f"sheet_name is required for '{action}' action")
            if sheet_names:
                raise ToolError(f"sheet_names not supported for '{action}' action")
            if action in ("add", "copy", "move", "rename"):
                if not validate_sheet_name(sheet_name):
                    raise ToolError(f"Invalid sheet name: '{sheet_name}'")
        elif action in ("hide", "unhide"):
            # hide/unhide support both single and bulk
            if sheet_name and sheet_names:
                raise ToolError("Use either sheet_name or sheet_names, not both")
            if not sheet_name and not sheet_names:
                raise ToolError("Either sheet_name or sheet_names is required")

        # Validate copy/move specific parameters
        if action in ("copy", "move", "rename"):
            if target_workbook and not validate_workbook_name(target_workbook):
                raise ToolError(f"Invalid target workbook name: '{target_workbook}'")
            if target_name and not validate_sheet_name(target_name):
                raise ToolError(f"Invalid target sheet name: '{target_name}'")
            if action == "rename" and not target_name:
                raise ToolError("target_name is required for 'rename' action")
            if position and position not in ("before", "after"):
                raise ToolError(f"Invalid position '{position}'. Must be 'before' or 'after'")
            if position and not reference_sheet:
                raise ToolError("reference_sheet is required when position is specified")

        try:
            if action == "add":
                result = await asyncio.to_thread(self._add_sheet_sync, workbook_name, sheet_name)
            elif action == "remove":
                result = await asyncio.to_thread(
                    self._remove_sheet_sync, workbook_name, sheet_name, force_delete
                )
            elif action == "copy":
                result = await asyncio.to_thread(
                    self._copy_sheet_sync,
                    workbook_name,
                    sheet_name,
                    target_workbook,
                    target_name,
                    position,
                    reference_sheet,
                )
            elif action == "move":
                result = await asyncio.to_thread(
                    self._move_sheet_sync,
                    workbook_name,
                    sheet_name,
                    target_workbook,
                    target_name,
                    position,
                    reference_sheet,
                )
            elif action == "rename":
                result = await asyncio.to_thread(
                    self._rename_sheet_sync, workbook_name, sheet_name, target_name
                )
            elif action == "hide":
                sheets = [sheet_name] if sheet_name else sheet_names
                result = await asyncio.to_thread(
                    self._hide_sheets_sync, workbook_name, sheets
                )
            else:  # unhide
                sheets = [sheet_name] if sheet_name else sheet_names
                result = await asyncio.to_thread(
                    self._unhide_sheets_sync, workbook_name, sheets
                )
            return result
        except Exception as e:
            error_msg = str(e).lower()
            if "read-only" in error_msg:
                raise ToolError(
                    f"Workbook '{workbook_name}' is read-only. Please open with write permissions."
                )
            raise ToolError(f"Failed to {action} sheet(s): {str(e)}") from e

    def _add_sheet_sync(self, workbook_name: str, sheet_name: str) -> Dict[str, Any]:
        """Synchronous add sheet."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        new_sheet = workbook.Worksheets.Add()
        new_sheet.Name = sheet_name

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "action": "added",
            "message": f"Sheet '{sheet_name}' created successfully",
        }

    def _is_sheet_empty_sync(self, workbook_name: str, sheet_name: str) -> bool:
        """Check if a worksheet is empty (no data in any cell).

        Args:
            workbook_name: Name of workbook
            sheet_name: Name of sheet to check

        Returns:
            True if sheet is completely empty, False otherwise
        """
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)

        # Use Excel's UsedRange to check if sheet has been used
        try:
            used_range = worksheet.UsedRange
            if used_range is None:
                return True  # No used range means empty sheet

            # If used range exists, check if it contains actual data
            # (Excel may create a UsedRange even for formatting-only changes)
            values = used_range.Value

            if values is None:
                return True

            # Check if any cell has actual data (not None, not empty)
            is_2d = isinstance(values, (list, tuple)) and len(values) > 0 and isinstance(values[0], (list, tuple))

            if is_2d:
                for row in values:
                    for val in row:
                        if val is not None and val != "" and str(val).strip() != "":
                            return False  # Found data
            else:
                for val in values:
                    if val is not None and val != "" and str(val).strip() != "":
                        return False  # Found data

            return True  # All cells are empty
        except Exception:
            return True  # If we can't determine, assume empty

    def _remove_sheet_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        force_delete: bool = False,
    ) -> Dict[str, Any]:
        """Synchronous sheet removal."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)

        # Check if sheet exists
        try:
            worksheet = workbook.Worksheets(sheet_name)
        except Exception:
            raise ToolError(f"Sheet '{sheet_name}' not found in workbook '{workbook_name}'")

        # Check if sheet is empty (unless force_delete)
        if not force_delete:
            if not self._is_sheet_empty_sync(workbook_name, sheet_name):
                raise ToolError(
                    f"Sheet '{sheet_name}' contains data. "
                    "Will not delete. Clear the sheet first or use force_delete=True."
                )

        # Don't allow deleting the last sheet
        if workbook.Worksheets.Count <= 1:
            raise ToolError(
                f"Cannot delete '{sheet_name}'. It is the last sheet in the workbook. "
                "At least one sheet must remain."
            )

        # Delete the sheet
        worksheet.Delete()

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "action": "removed",
            "message": f"Sheet '{sheet_name}' deleted successfully",
        }

    def _copy_sheet_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        target_workbook: str = None,
        target_name: str = None,
        position: str = None,
        reference_sheet: str = None,
    ) -> Dict[str, Any]:
        """Synchronous sheet copy within same workbook or to different workbook.

        Args:
            workbook_name: Name of source workbook
            sheet_name: Name of sheet to copy
            target_workbook: Name of target workbook (default: same as source)
            target_name: New name for copied sheet (default: Excel default)
            position: Position relative to reference_sheet ("before" or "after")
            reference_sheet: Sheet to position before/after

        Returns:
            Dictionary with operation result
        """
        pythoncom.CoInitialize()
        source_workbook = self.excel_app.Workbooks(workbook_name)
        source_sheet = source_workbook.Worksheets(sheet_name)

        # Default to same workbook
        if target_workbook is None:
            target_workbook = workbook_name

        # Check if target workbook is open
        try:
            target_workbook_obj = self.excel_app.Workbooks(target_workbook)
        except Exception:
            raise ToolError(f"Target workbook '{target_workbook}' is not open. Please open it first.")

        # Check if source sheet exists
        try:
            source_workbook.Worksheets(sheet_name)
        except Exception:
            raise ToolError(f"Source sheet '{sheet_name}' not found in workbook '{workbook_name}'")

        # Check reference sheet exists for positioning
        if position and reference_sheet:
            try:
                target_workbook_obj.Worksheets(reference_sheet)
            except Exception:
                raise ToolError(f"Reference sheet '{reference_sheet}' not found in workbook '{target_workbook}'")

        # Copy the sheet - IMPORTANT: Use positional arguments for win32com stability
        # sheet.Copy(Before, After)
        if position and reference_sheet:
            # Position before/after reference sheet
            ref_sheet = target_workbook_obj.Sheets(reference_sheet)
            if position == "before":
                source_sheet.Copy(ref_sheet, None)
            else:  # after
                source_sheet.Copy(None, ref_sheet)
        else:
            # Copy to end of target workbook
            last_sheet = target_workbook_obj.Sheets(target_workbook_obj.Sheets.Count)
            source_sheet.Copy(None, last_sheet)

        # Get the newly created sheet
        # When copying within the same workbook, the target_workbook_obj count increases
        # Usually the new sheet is added after the reference sheet or at the end
        if position == "before":
            # If inserted before, the index of the reference sheet is now the index of the new sheet
            ref_idx = target_workbook_obj.Worksheets(reference_sheet).Index
            new_sheet = target_workbook_obj.Worksheets(ref_idx)
        elif position == "after":
            # If inserted after, the index is ref_idx + 1
            ref_idx = target_workbook_obj.Worksheets(reference_sheet).Index
            new_sheet = target_workbook_obj.Worksheets(ref_idx + 1)
        else:
            # Default (copied to end)
            new_sheet = target_workbook_obj.Worksheets(target_workbook_obj.Worksheets.Count)

        actual_name = new_sheet.Name

        # Rename if requested
        if target_name:
            try:
                new_sheet.Name = target_name
                actual_name = target_name
            except Exception as e:
                # Fallback to current name if rename fails (might be duplicate)
                logger.warning(f"Failed to rename copied sheet: {str(e)}")

        same_workbook = target_workbook == workbook_name
        message_parts = [f"Sheet '{sheet_name}' copied"]
        if same_workbook:
            message_parts.append(f"as '{actual_name}' in {workbook_name}")
        else:
            message_parts.append(f"to '{actual_name}' in {target_workbook}")

        return {
            "success": True,
            "source_workbook": workbook_name,
            "source_sheet": sheet_name,
            "target_workbook": target_workbook,
            "target_sheet": actual_name,
            "action": "copied",
            "cross_workbook": not same_workbook,
            "message": " ".join(message_parts),
        }

    def _move_sheet_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        target_workbook: str = None,
        target_name: str = None,
        position: str = None,
        reference_sheet: str = None,
    ) -> Dict[str, Any]:
        """Synchronous sheet move within same workbook or to different workbook.

        Args:
            workbook_name: Name of source workbook
            sheet_name: Name of sheet to move
            target_workbook: Name of target workbook (default: same as source)
            target_name: New name for moved sheet (default: keep original)
            position: Position relative to reference_sheet ("before" or "after")
            reference_sheet: Sheet to position before/after

        Returns:
            Dictionary with operation result
        """
        pythoncom.CoInitialize()
        source_workbook = self.excel_app.Workbooks(workbook_name)

        # Check if source sheet exists
        try:
            source_sheet = source_workbook.Worksheets(sheet_name)
        except Exception:
            raise ToolError(f"Source sheet '{sheet_name}' not found in workbook '{workbook_name}'")

        # Default to same workbook
        if target_workbook is None:
            target_workbook = workbook_name

        # Check if target workbook is open
        try:
            target_workbook_obj = self.excel_app.Workbooks(target_workbook)
        except Exception:
            raise ToolError(f"Target workbook '{target_workbook}' is not open. Please open it first.")

        # Check reference sheet exists for positioning
        if position and reference_sheet:
            try:
                target_workbook_obj.Worksheets(reference_sheet)
            except Exception:
                raise ToolError(f"Reference sheet '{reference_sheet}' not found in workbook '{target_workbook}'")

        # Don't allow moving the last sheet out of a workbook
        if target_workbook != workbook_name and source_workbook.Worksheets.Count <= 1:
            raise ToolError(
                f"Cannot move '{sheet_name}'. It is the last sheet in '{workbook_name}'. "
                "At least one sheet must remain in the source workbook."
            )

        # Store original name for message
        original_name = sheet_name

        # Move the sheet - IMPORTANT: Use positional arguments for win32com stability
        # sheet.Move(Before, After)
        if position and reference_sheet:
            # Position before/after reference sheet
            ref_sheet = target_workbook_obj.Worksheets(reference_sheet)
            if position == "before":
                source_sheet.Move(ref_sheet, None)
            else:  # after
                source_sheet.Move(None, ref_sheet)
        elif target_workbook != workbook_name:
            # Move to different workbook (to end)
            target_last = target_workbook_obj.Sheets(target_workbook_obj.Sheets.Count)
            source_sheet.Move(None, target_last)
        else:
            # Move to end of same workbook
            last_sheet = source_workbook.Sheets(source_workbook.Sheets.Count)
            source_sheet.Move(None, last_sheet)

        # Rename if requested (only after move completes)
        final_name = original_name
        if target_name:
            try:
                # The source_sheet object stays valid even after move
                source_sheet.Name = target_name
                final_name = target_name
            except Exception as e:
                # Fallback to current name if rename fails
                logger.warning(f"Failed to rename moved sheet: {str(e)}")

        same_workbook = target_workbook == workbook_name
        message_parts = [f"Sheet '{original_name}' moved"]
        if same_workbook:
            if target_name:
                message_parts.append(f"and renamed to '{final_name}'")
        else:
            message_parts.append(f"to {target_workbook}")
            if target_name:
                message_parts.append(f"as '{final_name}'")

        return {
            "success": True,
            "source_workbook": workbook_name,
            "target_workbook": target_workbook,
            "sheet": final_name,
            "original_name": original_name,
            "action": "moved",
            "cross_workbook": not same_workbook,
            "message": " ".join(message_parts),
        }

    def _rename_sheet_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        target_name: str,
    ) -> Dict[str, Any]:
        """Synchronous sheet rename."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        sheet = workbook.Worksheets(sheet_name)

        sheet.Name = target_name

        return {
            "success": True,
            "workbook": workbook_name,
            "old_name": sheet_name,
            "new_name": target_name,
            "action": "renamed",
            "message": f"Sheet '{sheet_name}' renamed to '{target_name}' in {workbook_name}",
        }

    def _hide_sheets_sync(
        self,
        workbook_name: str,
        sheet_names: list,
    ) -> Dict[str, Any]:
        """Synchronous sheet hiding (supports bulk)."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)

        results = []
        errors = []

        for sheet_name in sheet_names:
            try:
                # Check if sheet exists
                worksheet = workbook.Worksheets(sheet_name)

                # Check if already hidden (xlSheetHidden = 0)
                if worksheet.Visible == 0:
                    results.append({
                        "sheet": sheet_name,
                        "status": "already_hidden",
                        "message": f"Sheet '{sheet_name}' is already hidden"
                    })
                    continue

                # Hide the sheet (xlSheetHidden = 0)
                worksheet.Visible = 0

                results.append({
                    "sheet": sheet_name,
                    "status": "hidden",
                    "message": f"Sheet '{sheet_name}' hidden successfully"
                })
            except Exception as e:
                errors.append(f"{sheet_name}: {str(e)}")

        if errors and not results:
            raise ToolError(f"Failed to hide sheets: {'; '.join(errors)}")

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "hidden",
            "results": results,
            "count": len(results),
            "errors": errors if errors else None,
        }

    def _unhide_sheets_sync(
        self,
        workbook_name: str,
        sheet_names: list,
    ) -> Dict[str, Any]:
        """Synchronous sheet unhiding (supports bulk)."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)

        results = []
        errors = []

        for sheet_name in sheet_names:
            try:
                # Check if sheet exists
                worksheet = workbook.Worksheets(sheet_name)

                # Check if already visible (xlSheetVisible = -1)
                if worksheet.Visible == -1:
                    results.append({
                        "sheet": sheet_name,
                        "status": "already_visible",
                        "message": f"Sheet '{sheet_name}' is already visible"
                    })
                    continue

                # Unhide the sheet (xlSheetVisible = -1)
                worksheet.Visible = -1

                results.append({
                    "sheet": sheet_name,
                    "status": "visible",
                    "message": f"Sheet '{sheet_name}' made visible successfully"
                })
            except Exception as e:
                errors.append(f"{sheet_name}: {str(e)}")

        if errors and not results:
            raise ToolError(f"Failed to unhide sheets: {'; '.join(errors)}")

        return {
            "success": True,
            "workbook": workbook_name,
            "action": "unhidden",
            "results": results,
            "count": len(results),
            "errors": errors if errors else None,
        }

    async def get_selection_info(self) -> Dict[str, Any]:
        """Get information about the current Excel selection.

        Returns:
            Dictionary with selection details including address, sheet,
            row/column info, and values
        """
        await self._connect()

        try:
            result = await asyncio.to_thread(self._get_selection_info_sync)
            return result
        except Exception as e:
            error_msg = str(e)[:200] if str(e) else "Unknown Excel error"
            raise ToolError(f"Failed to get selection: {error_msg}") from e

    def _get_selection_info_sync(self) -> Dict[str, Any]:
        """Synchronous selection info retrieval."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        # Get active window
        try:
            active_window = app.ActiveWindow
            if not active_window:
                return {
                    "success": True,
                    "has_selection": False,
                    "message": "No active window found"
                }
        except Exception as e:
            raise ToolError(f"Could not access active window: {str(e)}")

        # Get selection
        try:
            selection = active_window.Selection
            if not selection:
                return {
                    "success": True,
                    "has_selection": False,
                    "message": "No selection found"
                }
        except Exception as e:
            raise ToolError(f"Could not access selection: {str(e)}")

        # Build result dictionary
        result = {
            "success": True,
            "has_selection": True,
            "selection_type": str(type(selection).__name__),
        }

        # Handle Range selections (most common - cells, ranges, rows, columns)
        if hasattr(selection, 'Address'):
            try:
                result['address'] = selection.Address
                result['address_local'] = selection.AddressLocal  # e.g., $A$1 instead of A1

                # Get row/column count
                result['rows'] = selection.Rows.Count
                result['columns'] = selection.Columns.Count
                result['cell_count'] = selection.Cells.Count

                # Determine selection type
                # Excel max columns = 16384 (XFD), max rows = 1048576
                is_entire_row = result['rows'] == 1 and result['columns'] >= 16384
                is_entire_column = result['columns'] == 1 and result['rows'] >= 1048576

                result['selection_type'] = 'entire_row' if is_entire_row else (
                    'entire_column' if is_entire_column else (
                        'single_cell' if result['cell_count'] == 1 else 'range'
                    )
                )

                # Get workbook and sheet names
                if hasattr(selection, 'Parent'):
                    result['sheet'] = selection.Parent.Name
                    if hasattr(selection.Parent, 'Parent'):
                        result['workbook'] = selection.Parent.Parent.Name

                # Get value(s) for selection (skip for entire rows/columns due to size)
                if not is_entire_row and not is_entire_column:
                    try:
                        selection_value = selection.Value
                        if selection_value is not None:
                            if result['cell_count'] == 1:
                                # Single cell - return the value directly
                                if isinstance(selection_value, (int, float, bool)):
                                    result['value'] = selection_value
                                else:
                                    result['value'] = str(selection_value) if selection_value else ""
                            else:
                                # Multiple cells - return 2D array (like read_range)
                                values = []
                                if isinstance(selection_value, (list, tuple)):
                                    if len(selection_value) > 0 and isinstance(selection_value[0], (list, tuple)):
                                        # 2D array
                                        for row in selection_value:
                                            row_data = []
                                            for cell in row:
                                                row_data.append(str(cell) if cell is not None else "")
                                            values.append(row_data)
                                    else:
                                        # 1D array (single row)
                                        for cell in selection_value:
                                            values.append([str(cell) if cell is not None else ""])
                                result['values'] = values
                                result['value_rows'] = len(values)
                                result['value_cols'] = len(values[0]) if values else 0
                        else:
                            result['value'] = ""
                    except Exception:
                        result['value'] = None

            except Exception as e:
                logger.warning(f"Could not get range properties: {str(e)}")

        return result

    async def insert(
        self,
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
            position: Position (row number like "5", or column letter like "C", or number like "3")
            count: Number of rows/columns to insert

        Returns:
            Dictionary with operation result
        """
        await self._connect()

        # Validation
        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")
        if not validate_sheet_name(sheet_name):
            raise ToolError(f"Invalid sheet name: '{sheet_name}'")
        if insert_type not in ("row", "column"):
            raise ToolError(f"Invalid insert_type '{insert_type}'. Must be 'row' or 'column'")
        if count < 1:
            raise ToolError("Count must be 1 or greater")

        try:
            result = await asyncio.to_thread(
                self._insert_sync, workbook_name, sheet_name, insert_type, position, count
            )
            return result
        except Exception as e:
            error_msg = str(e).lower()
            if "protected" in error_msg:
                raise ToolError(
                    f"Worksheet '{sheet_name}' is protected. "
                    "Please unprotect in Excel first."
                )
            raise ToolError(f"Failed to insert: {str(e)}") from e

    def _insert_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        insert_type: str,
        position: str,
        count: int,
    ) -> Dict[str, Any]:
        """Synchronous insert operation."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)

        if insert_type == "row":
            # Parse position - could be "5" or a row number
            position = position.strip()
            if ":" in position:
                # Range format like "5:10" - use as is
                range_to_insert = position
                # Extract starting row
                start = int(position.split(":")[0])
            else:
                # Single position like "5"
                try:
                    start_row = int(position)
                    end_row = start_row + count - 1
                    range_to_insert = f"{start_row}:{end_row}"
                    start = start_row
                except ValueError:
                    raise ToolError(f"Invalid row position: '{position}'. Use row number like '5'")

            # Insert rows
            worksheet.Rows(range_to_insert).Insert()

            return {
                "success": True,
                "workbook": workbook_name,
                "sheet": sheet_name,
                "action": "rows_inserted",
                "count": count,
                "at": str(start),
                "message": f"Inserted {count} row(s) at row {start}"
            }

        else:  # column
            # Parse position - could be "C" or "3"
            position = position.strip()

            # Check if it's a number
            if position.isdigit():
                col_num = int(position)
                if col_num < 1 or col_num > 16384:
                    raise ToolError(f"Column number out of range: {col_num}")
                col_letter = self._number_to_column(col_num)
            else:
                # Assume it's a letter like "C"
                col_letter = position.upper()
                col_num = self._column_letter_to_number(col_letter)

            end_col_num = col_num + count - 1
            if end_col_num > 16384:
                raise ToolError(f"Cannot insert {count} columns - would exceed Excel's limit")

            end_col_letter = self._number_to_column(end_col_num)

            # Insert columns
            range_to_insert = f"{col_letter}:{end_col_letter}"
            worksheet.Columns(range_to_insert).Insert()

            return {
                "success": True,
                "workbook": workbook_name,
                "sheet": sheet_name,
                "action": "columns_inserted",
                "count": count,
                "at": col_letter,
                "message": f"Inserted {count} column(s) at column {col_letter}"
            }

    # Predefined formatting styles
    PREDEFINED_STYLES = {
        "header": {
            "font_bold": True,
            "font_size": 12,
            "fill_color": "4472C4",
            "font_color": "FFFFFF",
            "horizontal": "center",
            "vertical": "center"
        },
        "currency": {
            "number_format": "$#,##0.00",
            "horizontal": "right"
        },
        "percent": {
            "number_format": "0.00%",
            "horizontal": "right"
        },
        "warning": {
            "font_bold": True,
            "font_color": "FF0000",
            "fill_color": "FFFF00"
        },
        "success": {
            "font_color": "006100",
            "fill_color": "C6EFCE"
        },
        "border": {
            "border": True,
            "border_color": "000000",
            "border_weight": 2
        },
        "center": {
            "horizontal": "center",
            "vertical": "center"
        },
        "wrap": {
            "wrap_text": True
        }
    }

    async def format(
        self,
        workbook_name: str,
        sheet_name: str,
        reference: str,
        style: str = None,
        format: dict = None,
        activate: bool = True,
    ) -> Dict[str, Any]:
        """Apply formatting to cells/ranges.

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            reference: Cell/range reference or comma-separated references
            style: Predefined style name (optional)
            format: Custom format properties (optional)
            activate: If True, activate the workbook, sheet, and select the range after formatting.

        Returns:
            Dictionary with format result
        """
        await self._connect()

        # Validation
        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")
        if not validate_sheet_name(sheet_name):
            raise ToolError(f"Invalid sheet name: '{sheet_name}'")
        if not style and not format:
            raise ToolError("Either style or format must be specified")

        # Resolve style to format dict if style provided
        # If both style and format are provided, merge them (format overrides style)
        combined_format = {}
        if style:
            if style not in self.PREDEFINED_STYLES:
                available = ", ".join(self.PREDEFINED_STYLES.keys())
                raise ToolError(f"Unknown style '{style}'. Available: {available}")
            combined_format.update(self.PREDEFINED_STYLES[style])
        if format:
            combined_format.update(format)

        try:
            result = await asyncio.to_thread(
                self._format_sync, workbook_name, sheet_name, reference, combined_format, activate
            )
            return result
        except Exception as e:
            error_msg = str(e).lower()
            if "protected" in error_msg:
                raise ToolError(
                    f"Worksheet '{sheet_name}' is protected. "
                    "Please unprotect in Excel first."
                )
            raise ToolError(f"Failed to apply format: {str(e)}") from e

    def _format_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        reference: str,
        format: dict,
        activate: bool = True,
    ) -> Dict[str, Any]:
        """Synchronous format application.

        Args:
            workbook_name: Name of workbook
            sheet_name: Name of sheet
            reference: Cell/range reference or comma-separated
            format: Format properties dictionary
            activate: If True, activate the workbook, sheet, and select the range after formatting.

        Returns:
            Dictionary with format result
        """
        # Check if reference contains commas (multiple references)
        if "," in reference:
            refs = parse_reference_string(reference)

            # If format is a list, validate it matches refs count
            if isinstance(format, list):
                if len(format) != len(refs):
                    raise ToolError(
                        f"Format array length ({len(format)}) must match reference count ({len(refs)})"
                    )
                format_list = format
            else:
                # Single format applied to all references
                format_list = [format] * len(refs)

            results = []
            errors = []

            for i, ref in enumerate(refs):
                try:
                    result = self._apply_format_sync(
                    workbook_name, sheet_name, ref, format_list[i], activate
                )
                    results.append(result)
                except Exception as e:
                    errors.append(f"{ref}: {str(e)}")

            if errors and not results:
                raise ToolError(f"Failed to format: {'; '.join(errors)}")

            return {
                "success": True,
                "workbook": workbook_name,
                "sheet": sheet_name,
                "scattered": True,
                "reference": reference,
                "results": results,
                "count": len(results),
                "errors": errors if errors else None
            }

        else:
            # Single reference
            result = self._apply_format_sync(workbook_name, sheet_name, reference, format)
            result["scattered"] = False
            return result

    def _apply_format_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        reference: str,
        format: dict,
        activate: bool = True,
    ) -> Dict[str, Any]:
        """Apply format to a single cell or range."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)
        rng = worksheet.Range(reference)

        # Font properties
        font = rng.Font
        if "font_bold" in format:
            font.Bold = format["font_bold"]
        if "font_italic" in format:
            font.Italic = format["font_italic"]
        if "font_underline" in format:
            font.Underline = format["font_underline"]
        if "font_strikethrough" in format:
            font.Strikethrough = format["font_strikethrough"]
        if "font_size" in format:
            font.Size = format["font_size"]
        if "font_color" in format:
            font.Color = int(format["font_color"], 16)
        if "font_name" in format:
            font.Name = format["font_name"]

        # Fill (background)
        if "fill_color" in format:
            rng.Interior.Color = int(format["fill_color"], 16)

        # Alignment
        alignment_map = {"left": -4131, "center": -4108, "right": -4152, "justify": -4130}
        if "horizontal" in format:
            rng.HorizontalAlignment = alignment_map.get(format["horizontal"].lower(), -4108)

        vertical_map = {"top": -4160, "center": -4108, "bottom": -4107, "justify": -4116}
        if "vertical" in format:
            rng.VerticalAlignment = vertical_map.get(format["vertical"].lower(), -4108)

        # Wrap text
        if "wrap_text" in format:
            rng.WrapText = format["wrap_text"]

        # AutoFit
        if format.get("autofit"):
            rng.Columns.AutoFit()
            rng.Rows.AutoFit()
        if format.get("autofit_columns"):
            rng.Columns.AutoFit()
        if format.get("autofit_rows"):
            rng.Rows.AutoFit()

        # Column width
        if "column_width" in format:
            rng.ColumnWidth = format["column_width"]

        # Row height
        if "row_height" in format:
            rng.RowHeight = format["row_height"]

        # Merge/Unmerge
        if format.get("merge"):
            rng.Merge()
        if format.get("unmerge"):
            rng.UnMerge()

        # Number format
        if "number_format" in format:
            rng.NumberFormat = format["number_format"]

        # Borders
        if any(k.startswith("border") for k in format) or format.get("border"):
            borders = rng.Borders
            border_style_map = {
                "thin": 1, "continuous": 1, "dash": -4118,
                "medium": -4138, "thick": 4, "double": -4118
            }

            if "border_color" in format:
                borders.Color = int(format["border_color"], 16)
            if "border_style" in format:
                style = border_style_map.get(format["border_style"].lower(), 1)
                borders.LineStyle = style
            if "border_weight" in format:
                borders.Weight = format["border_weight"]

            # Apply to all edges if 'border' is true
            if format.get("border"):
                for edge in [7, 8, 9, 10]:  # xlLeft, xlRight, xlTop, xlBottom
                    borders.Item(edge).LineStyle = 1
                    borders.Item(edge).Weight = format.get("border_weight", 2)

        # Return result
        cell_count = rng.Cells.Count
        is_range = cell_count > 1

        if activate:
            try:
                workbook.Activate()
                worksheet.Activate()
                rng.Select()
            except:
                pass

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "cell": reference if not is_range else None,
            "range": reference if is_range else None,
            "cells_formatted": cell_count,
            "format": format
        }

    async def get_format(
        self,
        workbook_name: str,
        sheet_name: str,
        reference: str,
    ) -> Dict[str, Any]:
        """Get formatting properties from cells/ranges.

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            reference: Cell/range reference or comma-separated references

        Returns:
            Dictionary with formatting properties
        """
        await self._connect()

        # Validation
        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")
        if not validate_sheet_name(sheet_name):
            raise ToolError(f"Invalid sheet name: '{sheet_name}'")

        try:
            result = await asyncio.to_thread(
                self._get_format_sync, workbook_name, sheet_name, reference
            )
            return result
        except Exception as e:
            error_msg = str(e).lower()
            if "protected" in error_msg:
                raise ToolError(
                    f"Worksheet '{sheet_name}' is protected. "
                    "Please unprotect in Excel first."
                )
            raise ToolError(f"Failed to get format: {str(e)}") from e

    def _get_format_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        reference: str,
    ) -> Dict[str, Any]:
        """Synchronous format retrieval."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        # Check if reference contains commas (multiple references)
        if "," in reference:
            refs = parse_reference_string(reference)

            results = []
            errors = []

            for ref in refs:
                try:
                    result = self._get_single_format_sync(workbook_name, sheet_name, ref)
                    results.append(result)
                except Exception as e:
                    errors.append(f"{ref}: {str(e)}")

            if errors and not results:
                raise ToolError(f"Failed to get format: {'; '.join(errors)}")

            return {
                "success": True,
                "workbook": workbook_name,
                "sheet": sheet_name,
                "scattered": True,
                "reference": reference,
                "results": results,
                "count": len(results),
                "errors": errors if errors else None
            }

        else:
            # Single reference
            result = self._get_single_format_sync(workbook_name, sheet_name, reference)
            result["scattered"] = False
            return result

    def _get_single_format_sync(
        self,
        workbook_name: str,
        sheet_name: str,
        reference: str,
    ) -> Dict[str, Any]:
        """Get format from a single cell or range."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)
        rng = worksheet.Range(reference)

        # Reverse alignment maps for readable output
        horizontal_reverse_map = {-4131: "left", -4108: "center", -4152: "right", -4130: "justify"}
        vertical_reverse_map = {-4160: "top", -4108: "center", -4107: "bottom", -4116: "justify"}

        result = {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "reference": reference,
        }

        # Get font properties
        try:
            font = rng.Font
            result["font"] = {
                "bold": font.Bold if hasattr(font, 'Bold') else False,
                "italic": font.Italic if hasattr(font, 'Italic') else False,
                "underline": font.Underline if hasattr(font, 'Underline') else False,
                "strikethrough": font.Strikethrough if hasattr(font, 'Strikethrough') else False,
                "size": font.Size if hasattr(font, 'Size') else None,
                "color": hex(font.Color)[2:].upper() if hasattr(font, 'Color') else None,
                "name": font.Name if hasattr(font, 'Name') else None,
            }
        except Exception:
            result["font"] = None

        # Get fill (background) color
        try:
            interior = rng.Interior
            color = interior.Color if hasattr(interior, 'Color') else None
            result["fill"] = {
                "color": hex(color)[2:].upper() if color else None
            }
        except Exception:
            result["fill"] = None

        # Get alignment
        try:
            h_align = rng.HorizontalAlignment if hasattr(rng, 'HorizontalAlignment') else None
            v_align = rng.VerticalAlignment if hasattr(rng, 'VerticalAlignment') else None
            result["alignment"] = {
                "horizontal": horizontal_reverse_map.get(h_align, "general"),
                "vertical": vertical_reverse_map.get(v_align, "bottom"),
            }
        except Exception:
            result["alignment"] = None

        # Get wrap text
        try:
            result["wrap_text"] = rng.WrapText if hasattr(rng, 'WrapText') else False
        except Exception:
            result["wrap_text"] = None

        # Get number format
        try:
            result["number_format"] = rng.NumberFormat if hasattr(rng, 'NumberFormat') else "General"
        except Exception:
            result["number_format"] = None

        # Get borders info (simplified)
        try:
            borders = rng.Borders
            result["borders"] = {
                "has_borders": bool(borders and hasattr(borders, 'LineStyle') and borders.LineStyle != -4142),
            }
        except Exception:
            result["borders"] = None

        # Get column width - return array for multi-column ranges
        try:
            if rng.Columns.Count > 1:
                column_widths = []
                for i in range(1, rng.Columns.Count + 1):
                    column_widths.append(round(rng.Columns(i).ColumnWidth, 2))
                result["column_width"] = column_widths
            else:
                result["column_width"] = round(rng.ColumnWidth, 2)
        except Exception:
            result["column_width"] = None

        # Get row height - return array for multi-row ranges
        try:
            if rng.Rows.Count > 1:
                row_heights = []
                for i in range(1, rng.Rows.Count + 1):
                    row_heights.append(round(rng.Rows(i).RowHeight, 2))
                result["row_height"] = row_heights
            else:
                result["row_height"] = round(rng.RowHeight, 2)
        except Exception:
            result["row_height"] = None

        # Get merge status
        try:
            if rng.MergeCells:
                result["is_merged"] = True
                merge_area = rng.MergeArea.Address(False, False)
                result["merge_area"] = merge_area
            else:
                result["is_merged"] = False
                result["merge_area"] = None
        except Exception:
            result["is_merged"] = False
            result["merge_area"] = None

        # Determine if range or cell
        cell_count = rng.Cells.Count
        is_range = cell_count > 1

        if is_range:
            result["range"] = reference
            result["cell_count"] = cell_count
        else:
            result["cell"] = reference

        return result

    async def get_unique_values(
        self, workbook_name: str, sheet_name: str, range_str: Optional[str] = None
    ) -> Dict[str, Any]:
        """Get unique values from a range.

        Args:
            workbook_name: Name of open workbook
            sheet_name: Name of worksheet
            range_str: Range reference (e.g., "A1:A100"). Defaults to UsedRange.

        Returns:
            Dictionary with unique values and counts
        """
        await self._connect()

        if not validate_workbook_name(workbook_name):
            raise ToolError(f"Invalid workbook name: '{workbook_name}'")

        if not validate_sheet_name(sheet_name):
            raise ToolError(f"Invalid sheet name: '{sheet_name}'")

        if range_str and not validate_range_format(range_str) and not validate_cell_format(range_str):
             raise ToolError(f"Invalid reference format: '{range_str}'")

        try:
            result = await asyncio.to_thread(
                self._get_unique_values_sync, workbook_name, sheet_name, range_str
            )
            return result
        except Exception as e:
            raise ToolError(f"Failed to get unique values: {str(e)}") from e

    def _get_unique_values_sync(
        self, workbook_name: str, sheet_name: str, range_str: Optional[str] = None
    ) -> Dict[str, Any]:
        """Synchronous unique values retrieval."""
        pythoncom.CoInitialize()
        app = win32.GetActiveObject("Excel.Application")
        workbook = app.Workbooks(workbook_name)
        worksheet = workbook.Worksheets(sheet_name)

        if range_str:
            rng = worksheet.Range(range_str)
        else:
            rng = worksheet.UsedRange

        values = rng.Value
        from collections import Counter
        
        counts = Counter()
        if values is not None:
            # Flatten values
            if isinstance(values, (list, tuple)):
                for item in values:
                    if isinstance(item, (list, tuple)):
                        for subitem in item:
                            if subitem is not None and str(subitem).strip() != "":
                                counts[subitem] += 1
                    else:
                        if item is not None and str(item).strip() != "":
                            counts[item] += 1
            else:
                if values is not None and str(values).strip() != "":
                    counts[values] += 1

        # Sort by frequency (descending) then by value
        unique_results = []
        for val, count in counts.most_common():
            unique_results.append({"value": val, "count": count})

        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "range": rng.Address,
            "unique_count": len(counts),
            "values": unique_results,
        }
