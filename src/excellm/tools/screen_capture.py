"""Screen capture tool for ExceLLM MCP server.

Enables capturing visual screenshots of Excel sheets.
This is a COM-only feature (Windows + Excel running).
"""

import base64
import os
import tempfile
import time

from ..core.connection import get_excel_app, get_workbook, get_worksheet
from ..core.errors import ErrorCodes, ToolError


def capture_sheet_sync(
    workbook_name: str,
    sheet_name: str,
    range_ref: str = None,
    output_format: str = "file",
    output_path: str = None,
) -> dict:
    """Capture screenshot of Excel sheet or range.

    Uses Excel's CopyPicture method to capture the visual representation
    of cells including formatting, colors, borders, etc.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        range_ref: Specific range to capture (None = entire UsedRange)
        output_format: "base64" (default) or "file"
        output_path: File path if output_format is "file"

    Returns:
        Dictionary with:
        {
            "success": True,
            "format": "base64" or "file",
            "image_data": "..." (if base64),
            "file_path": "..." (if file),
            "mime_type": "image/png",
            "range_captured": "A1:H20",
            "sheet_name": "Sheet1"
        }

    Raises:
        ToolError: If capture fails
    """
    try:
        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)
        worksheet = get_worksheet(workbook, sheet_name)

        # Determine range to capture
        if range_ref:
            try:
                capture_range = worksheet.Range(range_ref)
            except Exception:
                raise ToolError(
                    f"Invalid range: '{range_ref}'",
                    code=ErrorCodes.INVALID_REFERENCE
                )
        else:
            # Use UsedRange
            capture_range = worksheet.UsedRange

        # Get the actual range address for reporting
        range_address = capture_range.Address.replace("$", "")

        # Activate workbook and sheet for proper capture
        workbook.Activate()
        worksheet.Activate()
        capture_range.Select()

        # Create temporary file for image
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, f"excel_capture_{int(time.time() * 1000)}.png")

        try:
            # Method 1: Export chart containing the range as picture
            # This is more reliable across Excel versions

            # Copy range as picture
            # xlScreen = 1, xlBitmap = 2
            capture_range.CopyPicture(Appearance=1, Format=2)

            # Brief pause to allow clipboard to populate
            time.sleep(0.2)

            # Create a temporary chart to paste and export
            chart_obj = worksheet.ChartObjects().Add(
                Left=capture_range.Left,
                Top=capture_range.Top,
                Width=capture_range.Width,
                Height=capture_range.Height
            )

            # Activate chart object to ensure Paste targets it
            chart_obj.Activate()
            chart = chart_obj.Chart

            # Paste the picture and export
            chart.Paste()

            # Export to file
            chart.Export(temp_file, "PNG")

            # Delete the temporary chart
            chart_obj.Delete()

            # Clear clipboard
            app.CutCopyMode = False

        except Exception as e:
            # Method 2: Try alternative approach using chart export
            try:
                # Alternative: Use Windows clipboard and save

                import win32clipboard

                # Copy range as picture
                capture_range.CopyPicture(Appearance=1, Format=2)

                # Get from clipboard
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
                    # Convert DIB to image
                    # This is simplified - in practice you might need more complex handling
                except Exception:
                    pass
                finally:
                    win32clipboard.CloseClipboard()

                app.CutCopyMode = False

            except Exception as e2:
                raise ToolError(
                    f"Failed to capture sheet: {str(e)}. Alternative method also failed: {str(e2)}",
                    code=ErrorCodes.WRITE_FAILED
                )

        # Read the image file
        if not os.path.exists(temp_file):
            raise ToolError(
                "Screenshot capture failed - no image file created",
                code=ErrorCodes.WRITE_FAILED
            )

        if output_format == "file":
            # Move to destination
            if output_path:
                import shutil
                shutil.move(temp_file, output_path)
                final_path = output_path
            else:
                final_path = temp_file

            return {
                "success": True,
                "format": "file",
                "file_path": final_path,
                "mime_type": "image/png",
                "range_captured": range_address,
                "sheet_name": sheet_name,
            }
        else:
            # Read and encode as base64
            with open(temp_file, 'rb') as f:
                image_data = base64.b64encode(f.read()).decode('utf-8')

            # Clean up temp file
            try:
                os.remove(temp_file)
            except Exception:
                pass

            return {
                "success": True,
                "format": "base64",
                "image_data": image_data,
                "mime_type": "image/png",
                "range_captured": range_address,
                "sheet_name": sheet_name,
            }

    except ToolError:
        raise
    except Exception as e:
        raise ToolError(
            f"Failed to capture sheet: {str(e)}",
            code=ErrorCodes.EXCEL_NOT_RUNNING
        )
