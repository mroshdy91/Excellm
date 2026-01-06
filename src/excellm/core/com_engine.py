"""COM-based Excel engine for Windows.

Uses pywin32 for live Excel automation.
Wraps existing ExceLLM COM functionality.
"""

from typing import Any, Dict, List

from .connection import get_excel_app, get_workbook, get_worksheet
from .engine import ExcelEngine
from .errors import ErrorCodes, ToolError


class COMEngine(ExcelEngine):
    """Windows COM automation engine using pywin32.

    Requires:
    - Windows OS
    - Excel running
    - Workbook open in Excel
    """

    def __init__(self):
        """Initialize COM engine."""
        self.app = get_excel_app()

    @property
    def engine_name(self) -> str:
        return "COM"

    @property
    def supports_live_excel(self) -> bool:
        return True

    def list_workbooks(self) -> List[Dict[str, Any]]:
        """List all open workbooks in Excel."""
        workbooks = []

        try:
            for wb_idx in range(1, self.app.Workbooks.Count + 1):
                wb = self.app.Workbooks(wb_idx)

                sheets = []
                for sheet_idx in range(1, wb.Worksheets.Count + 1):
                    sheet = wb.Worksheets(sheet_idx)
                    sheets.append({
                        "name": sheet.Name,
                        "hidden": sheet.Visible != -1  # -1 = xlSheetVisible
                    })

                workbooks.append({
                    "name": wb.Name,
                    "sheets": sheets,
                })

            return workbooks
        except Exception as e:
            raise ToolError(
                f"Failed to list workbooks: {str(e)}",
                code=ErrorCodes.EXCEL_NOT_RUNNING
           )

    def read_range(
        self,
        workbook_path: str,
        sheet_name: str,
        range_ref: str,
    ) -> List[List[Any]]:
        """Read range using COM."""
        try:
            # workbook_path is actually workbook name for COM
            workbook = get_workbook(self.app, workbook_path)
            worksheet = get_worksheet(workbook, sheet_name)

            range_obj = worksheet.Range(range_ref)
            values = range_obj.Value2

            if values is None:
                return [[]]

            # Handle single cell
            if not isinstance(values, tuple):
                return [[values]]

            # Handle multiple cells
            result = []
            for row in values:
                if isinstance(row, tuple):
                    result.append(list(row))
                else:
                    result.append([row])

            return result

        except Exception as e:
            raise ToolError(
                f"Failed to read range: {str(e)}",
                code=ErrorCodes.READ_FAILED
            )

    def write_range(
        self,
        workbook_path: str,
        sheet_name: str,
        range_ref: str,
        data: List[List[Any]],
    ) -> Dict[str, Any]:
        """Write range using COM."""
        try:
            workbook = get_workbook(self.app, workbook_path)
            worksheet = get_worksheet(workbook, sheet_name)

            range_obj = worksheet.Range(range_ref)

            # Convert data to tuple format for COM
            if len(data) == 1 and len(data[0]) == 1:
                # Single cell
                range_obj.Value = data[0][0]
            else:
                # Multiple cells - convert to tuple of tuples
                tuple_data = tuple(tuple(row) for row in data)
                range_obj.Value = tuple_data

            # Save workbook
            workbook.Save()

            return {
                "success": True,
                "cells_written": len(data) * len(data[0]) if data else 0,
                "range": range_ref,
            }

        except Exception as e:
            raise ToolError(
                f"Failed to write range: {str(e)}",
                code=ErrorCodes.WRITE_FAILED
            )

    def get_sheet_names(self, workbook_path: str) -> List[str]:
        """Get sheet names using COM."""
        try:
            workbook = get_workbook(self.app, workbook_path)
            sheets = []

            for sheet_idx in range(1, workbook.Worksheets.Count + 1):
                sheet = workbook.Worksheets(sheet_idx)
                sheets.append(sheet.Name)

            return sheets

        except Exception as e:
            raise ToolError(
                f"Failed to get sheet names: {str(e)}",
                code=ErrorCodes.WORKBOOK_NOT_FOUND
            )

    def create_sheet(
        self,
        workbook_path: str,
        sheet_name: str,
    ) -> Dict[str, Any]:
        """Create sheet using COM."""
        try:
            workbook = get_workbook(self.app, workbook_path)

            # Add new worksheet
            new_sheet = workbook.Worksheets.Add()
            new_sheet.Name = sheet_name

            workbook.Save()

            return {
                "success": True,
                "sheet_name": sheet_name,
                "message": f"Created sheet '{sheet_name}'",
            }

        except Exception as e:
            raise ToolError(
                f"Failed to create sheet: {str(e)}",
                code=ErrorCodes.WRITE_FAILED
            )

    def delete_sheet(
        self,
        workbook_path: str,
        sheet_name: str,
    ) -> Dict[str, Any]:
        """Delete sheet using COM."""
        try:
            workbook = get_workbook(self.app, workbook_path)
            worksheet = get_worksheet(workbook, sheet_name)

            # Disable alerts temporarily
            original_display_alerts = self.app.DisplayAlerts
            self.app.DisplayAlerts = False

            try:
                worksheet.Delete()
                workbook.Save()
            finally:
                self.app.DisplayAlerts = original_display_alerts

            return {
                "success": True,
                "sheet_name": sheet_name,
                "message": f"Deleted sheet '{sheet_name}'",
            }

        except Exception as e:
            raise ToolError(
                f"Failed to delete sheet: {str(e)}",
                code=ErrorCodes.WRITE_FAILED
            )

    # Advanced features supported by COM

    def execute_vba(
        self,
        workbook_path: str,
        vba_code: str,
        **kwargs
    ) -> Dict[str, Any]:
        """Execute VBA using existing implementation."""
        from ..tools.vba_execution import execute_vba_sync
        return execute_vba_sync(workbook_path, vba_code, **kwargs)

    def capture_sheet(
        self,
        workbook_path: str,
        sheet_name: str,
        **kwargs
    ) -> Dict[str, Any]:
        """Capture sheet using existing implementation."""
        from ..tools.screen_capture import capture_sheet_sync
        return capture_sheet_sync(workbook_path, sheet_name, **kwargs)
