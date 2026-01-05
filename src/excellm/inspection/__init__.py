"""Excel workbook and sheet inspection tools.

This module provides deterministic, stable tools for inspecting Excel
workbooks and worksheets via COM automation:

- inspect_workbook: Fast workbook-level radar across all sheets
- explore: Sheet-level radar with quick/deep modes
"""

from .explore import explore_sync
from .inspect_workbook import inspect_workbook_sync
from .types import (
    TOOL_VERSION,
    ExploreResult,
    InspectWorkbookResult,
    Outlier,
    Region,
    SheetIndex,
)

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
