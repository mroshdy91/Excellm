"""Formatting operations for ExceLLM MCP server.

Contains tools for applying and retrieving cell/range formatting.
"""

import logging
from typing import Any, Dict, Optional

from ..core.connection import (
    _init_com,
    get_excel_app,
    get_workbook,
    get_worksheet,
)

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


def _apply_conditional_format(rng, cond_format: Dict[str, Any]) -> None:
    """Apply conditional formatting to a range.

    Supports:
    - colorScale: 2 or 3 color gradient based on values
    - dataBar: Data bars showing value proportion
    - iconSet: Icons based on value thresholds
    - cellIs: Cell comparison rules (greater than, less than, etc.)
    """
    format_type = cond_format.get("type", "").lower()

    if format_type == "colorscale":
        # Color scale - gradient from min to max
        # Type 2 = xlColorScale
        cf = rng.FormatConditions.AddColorScale(ColorScaleType=2)

        # Set min color (default: red)
        min_color = cond_format.get("min_color", "F8696B")  # Light red
        if min_color:
            hex_color = min_color.lstrip("#")
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            cf.ColorScaleCriteria(1).FormatColor.Color = r + (g * 256) + (b * 65536)

        # Set max color (default: green)
        max_color = cond_format.get("max_color", "63BE7B")  # Light green
        if max_color:
            hex_color = max_color.lstrip("#")
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            cf.ColorScaleCriteria(2).FormatColor.Color = r + (g * 256) + (b * 65536)

    elif format_type == "databar":
        # Data bars - horizontal bars proportional to value
        cf = rng.FormatConditions.AddDatabar()

        # Set bar color (default: blue)
        bar_color = cond_format.get("bar_color", "638EC6")  # Light blue
        if bar_color:
            hex_color = bar_color.lstrip("#")
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            cf.BarColor.Color = r + (g * 256) + (b * 65536)

        # Show/hide values
        if cond_format.get("show_value") is False:
            cf.ShowValue = False

    elif format_type == "iconset":
        # Icon sets - traffic lights, arrows, etc.
        # IconSetType values: 1=3Arrows, 2=3ArrowsGray, 3=3Flags, 4=3TrafficLights, etc.
        icon_styles = {
            "3arrows": 1,
            "3arrowsgray": 2,
            "3flags": 3,
            "3trafficlights": 4,
            "3signs": 5,
            "3symbols": 6,
            "4arrows": 8,
            "4arrowsgray": 9,
            "4trafficlights": 11,
            "5arrows": 13,
            "5arrowsgray": 14,
            "5quarters": 16,
        }

        icon_style = cond_format.get("icon_style", "3trafficlights").lower()
        icon_type = icon_styles.get(icon_style, 4)  # Default to traffic lights

        cf = rng.FormatConditions.AddIconSetCondition()
        cf.IconSet = rng.Application.ActiveWorkbook.IconSets(icon_type)

        # Reverse icons if requested
        if cond_format.get("reverse", False):
            cf.ReverseOrder = True

    elif format_type == "cellis":
        # Cell comparison rules
        # Operator values: 1=Between, 3=Equal, 5=GreaterThan, 6=LessThan, 7=GreaterEqual, 8=LessEqual
        operators = {
            "between": 1,
            "notbetween": 2,
            "equal": 3,
            "notequal": 4,
            "greaterthan": 5,
            "lessthan": 6,
            "greaterequal": 7,
            "lessequal": 8,
        }

        operator = cond_format.get("operator", "greaterthan").lower().replace("_", "")
        op_value = operators.get(operator, 5)

        value1 = cond_format.get("value", cond_format.get("value1", 0))
        value2 = cond_format.get("value2")  # For between

        if value2 is not None:
            cf = rng.FormatConditions.Add(Type=1, Operator=op_value, Formula1=str(value1), Formula2=str(value2))
        else:
            cf = rng.FormatConditions.Add(Type=1, Operator=op_value, Formula1=str(value1))

        # Apply formatting for matching cells
        fill_color = cond_format.get("fill_color", "FFEB9C")  # Light yellow
        if fill_color:
            hex_color = fill_color.lstrip("#")
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            cf.Interior.Color = r + (g * 256) + (b * 65536)

        font_color = cond_format.get("font_color")
        if font_color:
            hex_color = font_color.lstrip("#")
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            cf.Font.Color = r + (g * 256) + (b * 65536)


def format_range_sync(
    workbook_name: str,
    sheet_name: str,
    reference: str,
    style: Optional[str] = None,
    format: Optional[Dict[str, Any]] = None,
    conditional_format: Optional[Dict[str, Any]] = None,
    activate: bool = True,
) -> Dict[str, Any]:
    """Apply formatting to cells/ranges in Excel.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        reference: Cell/range reference (supports comma-separated)
        style: Predefined style name (header, currency, percent, warning, success, border, center, wrap)
        format: Custom format properties
        conditional_format: Conditional formatting rules with keys:
            - type: 'colorScale', 'dataBar', 'iconSet', or 'cellIs'
            - For colorScale: min_color, max_color
            - For dataBar: bar_color, show_value
            - For iconSet: icon_style ('3arrows', '3trafficlights', etc.), reverse
            - For cellIs: operator, value, fill_color, font_color
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

        # Apply conditional formatting if specified
        if conditional_format:
            try:
                _apply_conditional_format(rng, conditional_format)
            except Exception as cf_error:
                logger.warning(f"Failed to apply conditional format: {cf_error}")

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


def merge_cells_sync(
    workbook_name: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
) -> Dict[str, Any]:
    """Merge a range of cells.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        start_cell: Starting cell of range (e.g., "A1")
        end_cell: Ending cell of range (e.g., "D1")

    Returns:
        Dictionary with operation result
    """
    _init_com()

    try:
        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)
        worksheet = get_worksheet(workbook, sheet_name)

        range_ref = f"{start_cell}:{end_cell}"
        rng = worksheet.Range(range_ref)

        # Get the value of the first cell (will be preserved after merge)
        first_value = rng.Cells(1, 1).Value

        rng.Merge()

        logger.info(f"Merged cells {range_ref} in {sheet_name}")

        return {
            "success": True,
            "message": f"Cells {range_ref} merged successfully",
            "range": range_ref,
            "preserved_value": first_value,
        }

    except Exception as e:
        logger.error(f"Failed to merge cells: {e}")
        return {
            "success": False,
            "error": f"Failed to merge cells: {str(e)}",
        }


def unmerge_cells_sync(
    workbook_name: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
) -> Dict[str, Any]:
    """Unmerge a previously merged range of cells.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        start_cell: Starting cell of range (e.g., "A1")
        end_cell: Ending cell of range (e.g., "D1")

    Returns:
        Dictionary with operation result
    """
    _init_com()

    try:
        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)
        worksheet = get_worksheet(workbook, sheet_name)

        range_ref = f"{start_cell}:{end_cell}"
        rng = worksheet.Range(range_ref)

        rng.UnMerge()

        logger.info(f"Unmerged cells {range_ref} in {sheet_name}")

        return {
            "success": True,
            "message": f"Cells {range_ref} unmerged successfully",
            "range": range_ref,
        }

    except Exception as e:
        logger.error(f"Failed to unmerge cells: {e}")
        return {
            "success": False,
            "error": f"Failed to unmerge cells: {str(e)}",
        }


def get_merged_cells_sync(
    workbook_name: str,
    sheet_name: str,
) -> Dict[str, Any]:
    """Get all merged cell ranges in a worksheet.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet

    Returns:
        Dictionary with list of merged ranges
    """
    _init_com()

    try:
        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)
        worksheet = get_worksheet(workbook, sheet_name)

        merged_ranges = []

        # Get UsedRange to scan for merged cells
        used_range = worksheet.UsedRange

        # Track processed merge areas to avoid duplicates
        processed = set()

        for cell in used_range:
            try:
                if cell.MergeCells:
                    merge_area = cell.MergeArea
                    address = merge_area.Address

                    if address not in processed:
                        processed.add(address)
                        merged_ranges.append({
                            "range": address.replace("$", ""),
                            "rows": merge_area.Rows.Count,
                            "columns": merge_area.Columns.Count,
                            "value": merge_area.Cells(1, 1).Value,
                        })
            except Exception:
                continue

        return {
            "success": True,
            "merged_ranges": merged_ranges,
            "count": len(merged_ranges),
        }

    except Exception as e:
        logger.error(f"Failed to get merged cells: {e}")
        return {
            "success": False,
            "error": f"Failed to get merged cells: {str(e)}",
        }

