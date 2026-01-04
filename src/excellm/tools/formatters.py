"""Formatting operations for ExceLLM MCP server.

Contains tools for applying and retrieving cell/range formatting.
"""

import logging
from typing import Any, Dict, List, Optional, Union

from ..core.connection import (
    get_excel_app,
    get_workbook,
    get_worksheet,
    _init_com,
)
from ..core.errors import ToolError, ErrorCodes
from ..core.utils import normalize_address

logger = logging.getLogger(__name__)

# Predefined styles
STYLES = {
    "header": {
        "font_bold": True,
        "fill_color": "4472C4",  # Blue
        "font_color": "FFFFFF",  # White
        "horizontal": "center",
    },
    "currency": {
        "number_format": "$#,##0.00",
        "horizontal": "right",
    },
    "percent": {
        "number_format": "0.00%",
        "horizontal": "right",
    },
    "warning": {
        "font_bold": True,
        "font_color": "FF0000",  # Red
        "fill_color": "FFFF00",  # Yellow
    },
    "success": {
        "font_color": "008000",  # Green
        "fill_color": "C6EFCE",  # Light green
    },
    "border": {
        "border": True,
    },
    "center": {
        "horizontal": "center",
        "vertical": "center",
    },
    "wrap": {
        "wrap_text": True,
    },
}


def _apply_format(rng, format_props: Dict[str, Any]) -> None:
    """Apply formatting properties to a range."""
    
    # Font properties
    if format_props.get("font_bold") is not None:
        rng.Font.Bold = format_props["font_bold"]
    
    if format_props.get("font_italic") is not None:
        rng.Font.Italic = format_props["font_italic"]
    
    if format_props.get("font_underline") is not None:
        rng.Font.Underline = 2 if format_props["font_underline"] else -4142  # xlUnderlineStyleSingle or xlNone
    
    if format_props.get("font_strikethrough") is not None:
        rng.Font.Strikethrough = format_props["font_strikethrough"]
    
    if format_props.get("font_size") is not None:
        rng.Font.Size = format_props["font_size"]
    
    if format_props.get("font_color"):
        # Convert hex RGB to Excel color
        hex_color = format_props["font_color"].lstrip("#")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        rng.Font.Color = r + (g * 256) + (b * 65536)
    
    if format_props.get("font_name"):
        rng.Font.Name = format_props["font_name"]
    
    # Fill color
    if format_props.get("fill_color"):
        hex_color = format_props["fill_color"].lstrip("#")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        rng.Interior.Color = r + (g * 256) + (b * 65536)
    
    # Alignment
    if format_props.get("horizontal"):
        h_align = format_props["horizontal"].lower()
        if h_align == "left":
            rng.HorizontalAlignment = -4131  # xlLeft
        elif h_align == "center":
            rng.HorizontalAlignment = -4108  # xlCenter
        elif h_align == "right":
            rng.HorizontalAlignment = -4152  # xlRight
    
    if format_props.get("vertical"):
        v_align = format_props["vertical"].lower()
        if v_align == "top":
            rng.VerticalAlignment = -4160  # xlTop
        elif v_align == "center":
            rng.VerticalAlignment = -4108  # xlCenter
        elif v_align == "bottom":
            rng.VerticalAlignment = -4107  # xlBottom
    
    if format_props.get("wrap_text") is not None:
        rng.WrapText = format_props["wrap_text"]
    
    # Number format
    if format_props.get("number_format"):
        rng.NumberFormat = format_props["number_format"]
    
    # Border
    if format_props.get("border"):
        # xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10
        # xlInsideVertical=11, xlInsideHorizontal=12
        for edge in [7, 8, 9, 10, 11, 12]:
            try:
                rng.Borders(edge).LineStyle = 1  # xlContinuous
                if format_props.get("border_weight"):
                    rng.Borders(edge).Weight = format_props["border_weight"]
                if format_props.get("border_color"):
                    hex_color = format_props["border_color"].lstrip("#")
                    r = int(hex_color[0:2], 16)
                    g = int(hex_color[2:4], 16)
                    b = int(hex_color[4:6], 16)
                    rng.Borders(edge).Color = r + (g * 256) + (b * 65536)
            except Exception:
                pass
    
    # Column width
    if format_props.get("column_width") is not None:
        rng.ColumnWidth = format_props["column_width"]
    
    # Row height
    if format_props.get("row_height") is not None:
        rng.RowHeight = format_props["row_height"]
    
    # AutoFit
    if format_props.get("autofit") or format_props.get("autofit_columns"):
        rng.Columns.AutoFit()
    
    if format_props.get("autofit") or format_props.get("autofit_rows"):
        rng.Rows.AutoFit()
    
    # Merge/Unmerge
    if format_props.get("merge"):
        rng.Merge()
    
    if format_props.get("unmerge"):
        rng.UnMerge()


def format_range_sync(
    workbook_name: str,
    sheet_name: str,
    reference: str,
    style: Optional[str] = None,
    format: Optional[Dict[str, Any]] = None,
    activate: bool = True,
) -> Dict[str, Any]:
    """Apply formatting to cells/ranges in Excel.
    
    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        reference: Cell/range reference (supports comma-separated)
        style: Predefined style name (header, currency, percent, warning, success, border, center, wrap)
        format: Custom format properties
        activate: If True, activate the range after formatting
        
    Returns:
        Dictionary with operation result
    """
    _init_com()
    
    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)
    
    # Parse comma-separated references
    references = [r.strip() for r in reference.split(",")]
    
    results = []
    total_cells = 0
    
    for ref in references:
        rng = worksheet.Range(ref)
        
        # Build format properties
        format_props = {}
        
        # Apply style first (base formatting)
        if style and style.lower() in STYLES:
            format_props.update(STYLES[style.lower()])
        
        # Override with custom format
        if format:
            format_props.update(format)
        
        if format_props:
            _apply_format(rng, format_props)
        
        cells_count = rng.Cells.Count
        total_cells += cells_count
        
        results.append({
            "reference": ref,
            "cells_formatted": cells_count,
        })
    
    if activate:
        try:
            workbook.Activate()
            worksheet.Activate()
            worksheet.Range(references[0]).Select()
        except Exception:
            pass
    
    if len(references) == 1:
        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "cells_formatted": total_cells,
        }
    else:
        return {
            "success": True,
            "workbook": workbook_name,
            "sheet": sheet_name,
            "scattered": True,
            "results": results,
            "count": len(results),
        }


def get_format_sync(
    workbook_name: str,
    sheet_name: str,
    reference: str,
) -> Dict[str, Any]:
    """Get formatting properties from cells/ranges in Excel.
    
    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        reference: Cell/range reference
        
    Returns:
        Dictionary with formatting properties
    """
    _init_com()
    
    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)
    
    # Parse comma-separated references
    references = [r.strip() for r in reference.split(",")]
    
    if len(references) > 1:
        results = []
        for ref in references:
            result = _get_single_format(worksheet, ref)
            results.append(result)
        
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
        result = _get_single_format(worksheet, references[0])
        result["success"] = True
        result["workbook"] = workbook_name
        result["sheet"] = sheet_name
        return result


def _get_single_format(worksheet, reference: str) -> Dict[str, Any]:
    """Get formatting for a single reference."""
    rng = worksheet.Range(reference)
    
    # Font properties
    font = {}
    try:
        font["bold"] = bool(rng.Font.Bold)
        font["italic"] = bool(rng.Font.Italic)
        font["underline"] = rng.Font.Underline != -4142  # xlNone
        font["strikethrough"] = bool(rng.Font.Strikethrough)
        font["size"] = rng.Font.Size
        font["name"] = rng.Font.Name
        
        # Font color
        color = rng.Font.Color
        if color:
            r = color % 256
            g = (color // 256) % 256
            b = (color // 65536) % 256
            font["color"] = f"{r:02X}{g:02X}{b:02X}"
    except Exception:
        pass
    
    # Fill
    fill = {}
    try:
        color = rng.Interior.Color
        if color:
            r = color % 256
            g = (color // 256) % 256
            b = (color // 65536) % 256
            fill["color"] = f"{r:02X}{g:02X}{b:02X}"
    except Exception:
        pass
    
    # Alignment
    alignment = {}
    try:
        h_align = rng.HorizontalAlignment
        if h_align == -4131:
            alignment["horizontal"] = "left"
        elif h_align == -4108:
            alignment["horizontal"] = "center"
        elif h_align == -4152:
            alignment["horizontal"] = "right"
        
        v_align = rng.VerticalAlignment
        if v_align == -4160:
            alignment["vertical"] = "top"
        elif v_align == -4108:
            alignment["vertical"] = "center"
        elif v_align == -4107:
            alignment["vertical"] = "bottom"
    except Exception:
        pass
    
    # Other properties
    wrap_text = None
    number_format = None
    try:
        wrap_text = bool(rng.WrapText)
        number_format = rng.NumberFormat
    except Exception:
        pass
    
    # Borders
    borders = {"has_borders": False}
    try:
        # Check if any border exists
        for edge in [7, 8, 9, 10]:  # Left, Top, Bottom, Right
            if rng.Borders(edge).LineStyle != -4142:  # xlNone
                borders["has_borders"] = True
                break
    except Exception:
        pass
    
    result = {
        "reference": reference,
        "font": font,
        "fill": fill,
        "alignment": alignment,
        "wrap_text": wrap_text,
        "number_format": number_format,
        "borders": borders,
    }
    
    if ":" in reference:
        result["range"] = reference
        result["cell_count"] = rng.Cells.Count
    else:
        result["cell"] = reference
    
    return result
