"""Excel Table creation tool for ExceLLM MCP server.

Enables creation of Excel Table objects (ListObjects) for better
data management, auto-filtering, and structured references.
"""


from ..core.connection import get_excel_app, get_workbook, get_worksheet
from ..core.errors import ErrorCodes, ToolError

# Common Excel table styles
TABLE_STYLES = {
    # Light styles
    "light1": "TableStyleLight1",
    "light2": "TableStyleLight2",
    "light9": "TableStyleLight9",
    "light14": "TableStyleLight14",
    # Medium styles (most common)
    "medium1": "TableStyleMedium1",
    "medium2": "TableStyleMedium2",
    "medium3": "TableStyleMedium3",
    "medium4": "TableStyleMedium4",
    "medium9": "TableStyleMedium9",
    "medium15": "TableStyleMedium15",
    # Dark styles
    "dark1": "TableStyleDark1",
    "dark2": "TableStyleDark2",
    "dark11": "TableStyleDark11",
}


def create_table_sync(
    workbook_name: str,
    sheet_name: str,
    range_ref: str,
    table_name: str,
    has_headers: bool = True,
    table_style: str = "medium2",
) -> dict:
    """Create an Excel Table (ListObject) from a range.

    Excel Tables provide:
    - Automatic formatting
    - Built-in filtering
    - Structured references
    - Easy expansion when adding data

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range_ref: Range for the table (e.g., "A1:D100")
        table_name: Name for the table (must be unique in workbook)
        has_headers: Whether first row contains headers (default: True)
        table_style: Style name - use shortcuts like "medium2" or full name
                    like "TableStyleMedium2" (default: "medium2")

    Returns:
        Dictionary with:
        {
            "success": True,
            "table_name": "SalesData",
            "range": "A1:D100",
            "rows": 100,
            "columns": 4,
            "has_headers": True,
            "style": "TableStyleMedium2",
            "headers": ["Name", "Date", "Amount", "Status"]
        }

    Raises:
        ToolError: If table creation fails
    """
    try:
        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)
        worksheet = get_worksheet(workbook, sheet_name)

        # Validate range
        try:
            table_range = worksheet.Range(range_ref)
        except Exception:
            raise ToolError(
                f"Invalid range: '{range_ref}'",
                code=ErrorCodes.INVALID_REFERENCE
            )

        # Check if table name already exists
        try:
            for table in worksheet.ListObjects:
                if table.Name.lower() == table_name.lower():
                    raise ToolError(
                        f"Table '{table_name}' already exists on this sheet.",
                        code=ErrorCodes.WRITE_FAILED
                    )
        except Exception as e:
            if "already exists" in str(e):
                raise
            # ListObjects might not be accessible, continue

        # Also check workbook-level for duplicate names
        try:
            for sheet in workbook.Worksheets:
                for table in sheet.ListObjects:
                    if table.Name.lower() == table_name.lower():
                        raise ToolError(
                            f"Table '{table_name}' already exists in workbook (on sheet '{sheet.Name}').",
                            code=ErrorCodes.WRITE_FAILED
                        )
        except Exception as e:
            if "already exists" in str(e):
                raise

        # Resolve table style
        if table_style.lower() in TABLE_STYLES:
            full_style_name = TABLE_STYLES[table_style.lower()]
        elif table_style.startswith("TableStyle"):
            full_style_name = table_style
        else:
            # Try to make it a valid style name
            full_style_name = f"TableStyleMedium{table_style}" if table_style.isdigit() else "TableStyleMedium2"

        # xlSrcRange = 1, xlYes = 1, xlNo = 2
        xl_has_headers = 1 if has_headers else 2

        try:
            # Create the table
            # ListObjects.Add(SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination)
            new_table = worksheet.ListObjects.Add(
                SourceType=1,  # xlSrcRange
                Source=table_range,
                XlListObjectHasHeaders=xl_has_headers
            )

            # Set the table name
            new_table.Name = table_name

            # Apply style
            try:
                new_table.TableStyle = full_style_name
            except Exception:
                # Style might not exist, use default
                try:
                    new_table.TableStyle = "TableStyleMedium2"
                except Exception:
                    pass

            # Get table info
            table_range_addr = table_range.Address.replace("$", "")
            row_count = table_range.Rows.Count
            col_count = table_range.Columns.Count

            # Get headers if available
            headers = []
            if has_headers:
                try:
                    for col in new_table.HeaderRowRange:
                        val = col.Value
                        if val:
                            headers.append(str(val))
                        else:
                            headers.append("")
                except Exception:
                    pass

            return {
                "success": True,
                "table_name": table_name,
                "range": table_range_addr,
                "rows": row_count,
                "columns": col_count,
                "has_headers": has_headers,
                "style": full_style_name,
                "headers": headers if headers else None,
                "message": f"Created table '{table_name}' with {row_count} rows and {col_count} columns"
            }

        except Exception as e:
            error_msg = str(e)
            if "1004" in error_msg:
                raise ToolError(
                    f"Cannot create table in range '{range_ref}'. Range may overlap with existing table or be invalid.",
                    code=ErrorCodes.WRITE_FAILED
                )
            raise ToolError(
                f"Failed to create table: {error_msg}",
                code=ErrorCodes.WRITE_FAILED
            )

    except ToolError:
        raise
    except Exception as e:
        raise ToolError(
            f"Failed to create table: {str(e)}",
            code=ErrorCodes.EXCEL_NOT_RUNNING
        )


def list_tables_sync(
    workbook_name: str,
    sheet_name: str = None,
) -> dict:
    """List all tables in a workbook or specific sheet.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Optional - specific sheet (None = all sheets)

    Returns:
        Dictionary with list of tables
    """
    try:
        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)

        tables = []

        if sheet_name:
            sheets_to_check = [get_worksheet(workbook, sheet_name)]
        else:
            sheets_to_check = [workbook.Worksheets(i) for i in range(1, workbook.Worksheets.Count + 1)]

        for sheet in sheets_to_check:
            try:
                for table in sheet.ListObjects:
                    table_info = {
                        "name": table.Name,
                        "sheet": sheet.Name,
                        "range": table.Range.Address.replace("$", ""),
                        "rows": table.Range.Rows.Count,
                        "columns": table.Range.Columns.Count,
                        "style": str(table.TableStyle) if table.TableStyle else None,
                    }

                    # Get headers
                    try:
                        headers = []
                        if table.HeaderRowRange:
                            for col in table.HeaderRowRange:
                                val = col.Value
                                headers.append(str(val) if val else "")
                        table_info["headers"] = headers
                    except Exception:
                        pass

                    tables.append(table_info)
            except Exception:
                continue

        return {
            "success": True,
            "tables": tables,
            "count": len(tables),
        }

    except ToolError:
        raise
    except Exception as e:
        raise ToolError(
            f"Failed to list tables: {str(e)}",
            code=ErrorCodes.EXCEL_NOT_RUNNING
        )


def delete_table_sync(
    workbook_name: str,
    sheet_name: str,
    table_name: str,
    keep_data: bool = True,
) -> dict:
    """Delete an Excel Table, optionally keeping the data.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        table_name: Name of table to delete
        keep_data: If True, convert to normal range; if False, delete data too

    Returns:
        Dictionary with operation result
    """
    try:
        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)
        worksheet = get_worksheet(workbook, sheet_name)

        # Find the table
        target_table = None
        for table in worksheet.ListObjects:
            if table.Name.lower() == table_name.lower():
                target_table = table
                break

        if not target_table:
            raise ToolError(
                f"Table '{table_name}' not found on sheet '{sheet_name}'.",
                code=ErrorCodes.SHEET_NOT_FOUND
            )

        table_range = target_table.Range.Address.replace("$", "")

        if keep_data:
            # Convert to range (removes table formatting but keeps data)
            target_table.Unlist()
            message = f"Table '{table_name}' converted to normal range. Data preserved."
        else:
            # Delete table and data
            target_table.Delete()
            message = f"Table '{table_name}' and its data deleted."

        return {
            "success": True,
            "table_name": table_name,
            "range": table_range,
            "data_preserved": keep_data,
            "message": message,
        }

    except ToolError:
        raise
    except Exception as e:
        raise ToolError(
            f"Failed to delete table: {str(e)}",
            code=ErrorCodes.EXCEL_NOT_RUNNING
        )
