
import re
from typing import Any, Dict, List

import pythoncom
import win32com.client as win32

from ..core.errors import ToolError

# Command Bar Control IDs
ID_UNDO = 128
ID_REDO = 129

def get_recent_changes_sync(limit: int = 10) -> Dict[str, List[Dict[str, Any]]]:
    """Retrieve items from Excel's Undo and Redo stacks.

    Args:
        limit: Max number of items to retrieve per stack

    Returns:
        Dict with keys 'undo' and 'redo', each containing list of history items
    """
    pythoncom.CoInitialize()
    try:
        try:
            excel = win32.GetActiveObject("Excel.Application")
        except Exception as e:
            raise ToolError("Could not connect to Excel. Is it running?") from e

        def get_stack(control_id):
            """Helper to extract items from a CommandBar control"""
            control = None
            try:
                # Iterate to find "Standard" safely and then the control
                # This is more robust than FindControl for some versions
                for cb in excel.CommandBars:
                    if cb.Name == "Standard":
                        for c in cb.Controls:
                            if c.Id == control_id:
                                control = c
                                break
                        break

                # Fallback
                if not control:
                    control = excel.CommandBars.FindControl(Id=control_id)

                if not control:
                    return []

                items = []
                # ListCount access check
                try:
                    count = control.ListCount
                except Exception:
                    return []

                num_items = min(limit, count)
                for i in range(1, num_items + 1):
                    try:
                        desc = control.List(i)

                        probable_address = None
                        match = re.search(r"in\s+((?:['\w\s]+!)?\$?[A-Z]{1,3}\$?[0-9]{1,7})$", desc, re.IGNORECASE)
                        if match:
                            probable_address = match.group(1).strip()

                        items.append({
                            "index": i,
                            "description": desc,
                            "probable_address": probable_address
                        })
                    except Exception:
                        continue
                return items
            except Exception:
                return []

        return {
            "undo": get_stack(ID_UNDO),
            "redo": get_stack(ID_REDO)
        }

    except Exception as e:
        if isinstance(e, ToolError):
            raise
        return {"undo": [], "redo": []}
