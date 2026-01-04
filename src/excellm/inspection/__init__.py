"""Excel workbook and sheet inspection tools.

This module provides deterministic, stable tools for inspecting Excel
workbooks and worksheets via COM automation:

- inspect_workbook: Fast workbook-level radar across all sheets
- explore: Sheet-level radar with quick/deep modes
"""

from .types import (
    InspectWorkbookResult,
    SheetIndex,
    ExploreResult,
    Region,
    Outlier,
    TOOL_VERSION,
)
from .inspect_workbook import inspect_workbook_sync
from .explore import explore_sync

__all__ = [
    "InspectWorkbookResult",
    "SheetIndex", 
    "ExploreResult",
    "Region",
    "Outlier",
    "TOOL_VERSION",
    "inspect_workbook_sync",
    "explore_sync",
]
