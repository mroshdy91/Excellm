
import re
import pythoncom
import win32com.client as win32
from typing import List, Dict, Any, Optional

from ..core.errors import ToolError

# Command Bar Control IDs
ID_UNDO = 128
ID_REDO = 129

def get_recent_changes_sync(limit: int = 10, history_type: str = "undo") -> List[Dict[str, Any]]:
    """Retrieve items from Excel's Undo or Redo stack.
    
    Args:
        limit: Max number of items to retrieve
        history_type: "undo" or "redo"
        
    Returns:
        List of history items with descriptions and probable addresses
    """
    pythoncom.CoInitialize()
    try:
        try:
            excel = win32.GetActiveObject("Excel.Application")
        except Exception as e:
            raise ToolError("Could not connect to Excel. Is it running?") from e
            
        target_id = ID_REDO if history_type.lower() == "redo" else ID_UNDO
        target_name = "Redo" if target_id == ID_REDO else "Undo"
        
        # Locate the control
        # Strategy: Look in the "Standard" CommandBar first (most reliable for List access)
        # Note: In Ribbon UI, this is a legacy command bar but often still populated
        control = None
        try:
            std_bar = excel.CommandBars("Standard")
            # Iterate to find by ID
            for i in range(1, std_bar.Controls.Count + 1):
                c = std_bar.Controls(i)
                if c.Id == target_id:
                    control = c
                    break
        except Exception:
            # Fallback: FindControl (global search)
            # Warning: logic for .List on FindControl result is sometimes flaky if not cast correctly
            control = excel.CommandBars.FindControl(Id=target_id)
            
        if not control:
            # It's possible the Standard bar is missing or ID changed (unlikely for 128/129)
            return []
            
        items = []
        
        # Access the list
        # Check if list exists
        try:
            # ListCount might fail if control is disabled or empty
            count = control.ListCount
        except Exception:
            return []
            
        # Get items
        num_items = min(limit, count)
        for i in range(1, num_items + 1):
            try:
                # Excel 1-based index
                desc = control.List(i)
                
                # Heuristic parsing for address
                # Looks for patterns like " in A1" or " in Sheet1!A1"
                # Regex: "in (Sheet\d+!)?\$?[A-Z]{1,3}\$?\d+" (simplified)
                
                probable_address = None
                # Basic pattern for address at end of string
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
        
    except Exception as e:
        if isinstance(e, ToolError):
            raise
        # Return empty list on failure rather than crashing, as this is inspection
        return []
