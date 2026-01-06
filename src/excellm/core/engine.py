"""Excel engine abstraction layer.

Provides a unified interface for Excel operations that can be
implemented by different backends:
- COMEngine: Windows + Excel running (pywin32)
- FileEngine: Cross-platform file-based (openpyxl)
"""

import platform
from abc import ABC, abstractmethod
from typing import Any, Dict, List


class ExcelEngine(ABC):
    """Abstract base class for Excel operation engines."""

    @property
    @abstractmethod
    def engine_name(self) -> str:
        """Name of the engine (e.g., 'COM', 'File')."""
        pass

    @property
    @abstractmethod
    def supports_live_excel(self) -> bool:
        """Whether this engine supports live Excel operations."""
        pass

    @abstractmethod
    def list_workbooks(self) -> List[Dict[str, Any]]:
        """List all available workbooks.

        Returns:
            List of workbook info dictionaries
        """
        pass

    @abstractmethod
    def read_range(
        self,
        workbook_path: str,
        sheet_name: str,
        range_ref: str,
    ) -> List[List[Any]]:
        """Read a range of cells.

        Args:
            workbook_path: Path or name of workbook
            sheet_name: Name of worksheet
            range_ref: Range reference (e.g., "A1:C5")

        Returns:
            2D list of cell values
        """
        pass

    @abstractmethod
    def write_range(
        self,
        workbook_path: str,
        sheet_name: str,
        range_ref: str,
        data: List[List[Any]],
    ) -> Dict[str, Any]:
        """Write data to a range of cells.

        Args:
            workbook_path: Path or name of workbook
            sheet_name: Name of worksheet
            range_ref: Range reference (e.g., "A1:C5")
            data: 2D list of values to write

        Returns:
            Dictionary with operation result
        """
        pass

    @abstractmethod
    def get_sheet_names(self, workbook_path: str) -> List[str]:
        """Get all sheet names in a workbook.

        Args:
            workbook_path: Path or name of workbook

        Returns:
            List of sheet names
        """
        pass

    @abstractmethod
    def create_sheet(
        self,
        workbook_path: str,
        sheet_name: str,
    ) -> Dict[str, Any]:
        """Create a new worksheet.

        Args:
            workbook_path: Path or name of workbook
            sheet_name: Name for new sheet

        Returns:
            Dictionary with operation result
        """
        pass

    @abstractmethod
    def delete_sheet(
        self,
        workbook_path: str,
        sheet_name: str,
    ) -> Dict[str, Any]:
        """Delete a worksheet.

        Args:
            workbook_path: Path or name of workbook
            sheet_name: Name of sheet to delete

        Returns:
            Dictionary with operation result
        """
        pass

    # Optional advanced features (may not be supported by all engines)

    def execute_vba(
        self,
        workbook_path: str,
        vba_code: str,
        **kwargs
    ) -> Dict[str, Any]:
        """Execute VBA code (COM engine only).

        Raises:
            NotImplementedError: If engine doesn't support VBA
        """
        raise NotImplementedError(f"{self.engine_name} engine does not support VBA execution")

    def capture_sheet(
        self,
        workbook_path: str,
        sheet_name: str,
        **kwargs
    ) -> Dict[str, Any]:
        """Capture sheet screenshot (COM engine only).

        Raises:
            NotImplementedError: If engine doesn't support screen capture
        """
        raise NotImplementedError(f"{self.engine_name} engine does not support screen capture")


def is_windows() -> bool:
    """Check if running on Windows."""
    return platform.system() == "Windows"


def is_excel_running() -> bool:
    """Check if Excel is running on Windows."""
    if not is_windows():
        return False

    try:
        import win32com.client
        excel = win32com.client.GetActiveObject("Excel.Application")
        # Try to access a property to verify connection
        _ = excel.Workbooks.Count
        return True
    except Exception:
        return False


class EngineFactory:
    """Factory for creating appropriate Excel engine based on context."""

    @staticmethod
    def create_engine(
        prefer_live: bool = True,
        workbook_path: str = None,
    ) -> ExcelEngine:
        """Create the most appropriate engine for the context.

        Args:
            prefer_live: Prefer live Excel if available
            workbook_path: Path to workbook (helps determine if file-based is needed)

        Returns:
            ExcelEngine instance (COMEngine or FileEngine)
        """
        # If workbook_path looks like a full path (not just a name), prefer file mode
        is_full_path = workbook_path and ('/' in workbook_path or '\\' in workbook_path)

        if prefer_live and not is_full_path and is_windows() and is_excel_running():
            # Use COM engine for live Excel
            from .com_engine import COMEngine
            return COMEngine()
        else:
            # Use file-based engine
            from .file_engine import FileEngine
            return FileEngine()

    @staticmethod
    def get_default_engine() -> ExcelEngine:
        """Get the default engine for the current environment."""
        return EngineFactory.create_engine(prefer_live=True)
