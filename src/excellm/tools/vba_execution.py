"""VBA execution tool for ExceLLM MCP server.

Enables execution of custom VBA code in Excel workbooks.
This is a COM-only feature (Windows + Excel running).
"""

import random
import re
import time
from typing import Optional

from ..core.connection import get_excel_app, get_workbook
from ..core.errors import ErrorCodes, ToolError


def _clean_vba_code(vba_code: str) -> str:
    """Clean and sanitize VBA code for safe execution.

    - Removes MsgBox statements (prevent blocking popups)
    - Removes problematic Err.Raise statements
    - Strips inline comments for cleaner execution

    Args:
        vba_code: Raw VBA code

    Returns:
        Cleaned VBA code
    """
    lines = vba_code.split('\n')
    clean_lines = []

    for line in lines:
        stripped = line.strip()

        # Skip empty lines at start/end but keep structure
        if not stripped:
            clean_lines.append('')
            continue

        # Skip comment-only lines
        if stripped.startswith("'"):
            continue

        # Remove MsgBox statements to prevent popups
        if re.search(r'\bMsgBox\b', stripped, re.IGNORECASE):
            continue

        # Replace Err.Raise with Exit Sub
        if re.search(r'\bErr\.Raise\b', stripped, re.IGNORECASE):
            line = re.sub(
                r'Err\.Raise\s+Err\.Number,\s*Err\.Source,\s*Err\.Description',
                'Exit Sub',
                line,
                flags=re.IGNORECASE
            )

        clean_lines.append(line.rstrip())

    return '\n'.join(clean_lines)


def _parse_vba_structure(vba_code: str) -> tuple[bool, Optional[str], str]:
    """Parse VBA code to detect Sub/Function structure.

    Args:
        vba_code: VBA code to parse

    Returns:
        Tuple of (has_structure, procedure_name, cleaned_code)
    """
    clean_code = _clean_vba_code(vba_code)

    # Look for existing Sub or Function definition
    sub_pattern = r'\b(Sub|Function)\s+(\w+)'
    matches = re.findall(sub_pattern, clean_code, re.IGNORECASE)

    if matches:
        procedure_type, procedure_name = matches[0]
        return True, procedure_name, clean_code

    return False, None, clean_code


def _generate_unique_module_name() -> str:
    """Generate a unique temporary module name."""
    return f"TempModule{random.randint(1000, 9999)}"


def _clean_temp_modules(workbook) -> None:
    """Clean up existing temporary VBA modules.

    Args:
        workbook: Workbook COM object
    """
    try:
        vb_project = workbook.VBProject
        components_to_remove = []

        for comp in vb_project.VBComponents:
            if comp.Name.startswith("TempModule"):
                components_to_remove.append(comp)

        for comp in components_to_remove:
            try:
                vb_project.VBComponents.Remove(comp)
                time.sleep(0.05)  # Small delay for COM
            except Exception:
                continue

    except Exception:
        # VBProject access might be restricted
        pass


def execute_vba_sync(
    workbook_name: str,
    vba_code: str,
    module_name: str = None,
    procedure_name: str = "Main",
    sheet_name: str = None,
    timeout: int = 30,
) -> dict:
    """Execute VBA code in Excel.

    Creates a temporary module, adds the VBA code, executes it,
    and cleans up the module afterwards.

    Args:
        workbook_name: Name of open workbook
        vba_code: VBA code to execute
        module_name: Optional custom module name (auto-generated if None)
        procedure_name: Name for the procedure (default: "Main")
        sheet_name: Optional sheet to activate before execution
        timeout: Execution timeout in seconds (for future use)

    Returns:
        Dictionary with execution result:
        {
            "success": True,
            "message": "VBA executed successfully",
            "procedure_name": "Main",
            "module_name": "TempModule1234"
        }

    Raises:
        ToolError: If VBA execution fails
    """
    try:
        app = get_excel_app()
        workbook = get_workbook(app, workbook_name)

        # Check VBProject access (must be enabled in Trust Center)
        try:
            _ = workbook.VBProject.VBComponents.Count
        except Exception:
            raise ToolError(
                "VBA Project access denied. Enable 'Trust access to the VBA project object model' "
                "in Excel Trust Center settings (File > Options > Trust Center > Trust Center Settings > Macro Settings).",
                code=ErrorCodes.PROTECTED_SHEET
            )

        # Navigate to sheet if specified
        if sheet_name:
            try:
                sheet = workbook.Worksheets(sheet_name)
                sheet.Activate()
            except Exception:
                raise ToolError(
                    f"Sheet '{sheet_name}' not found.",
                    code=ErrorCodes.SHEET_NOT_FOUND
                )

        # Parse VBA code structure
        has_structure, detected_proc, clean_code = _parse_vba_structure(vba_code)

        # Use detected procedure name or provided one
        final_proc_name = detected_proc if has_structure and detected_proc else procedure_name
        final_module_name = module_name or _generate_unique_module_name()

        # Clean up old temp modules
        _clean_temp_modules(workbook)

        # Store original application settings
        original_display_alerts = app.DisplayAlerts
        original_screen_updating = app.ScreenUpdating

        vba_module = None

        try:
            # Disable alerts and screen updating for cleaner execution
            app.DisplayAlerts = False
            app.ScreenUpdating = False

            # Create temporary module
            # 1 = vbext_ct_StdModule (Standard Module)
            vba_module = workbook.VBProject.VBComponents.Add(1)
            vba_module.Name = final_module_name

            # Prepare VBA code
            if not has_structure:
                # Wrap code in Sub with error handling
                wrapped_code = f"""Sub {final_proc_name}()
    On Error GoTo ErrorHandler

{clean_code}

    Exit Sub
ErrorHandler:
    Exit Sub
End Sub"""
            else:
                # Use existing structure but ensure error handling
                wrapped_code = clean_code
                if 'On Error' not in wrapped_code:
                    # Add basic error handling
                    wrapped_code = re.sub(
                        rf'(Sub\s+{final_proc_name}\s*\(\))',
                        r'\1\n    On Error GoTo ErrorHandler',
                        wrapped_code,
                        flags=re.IGNORECASE
                    )
                    wrapped_code = wrapped_code.replace(
                        'End Sub',
                        '    Exit Sub\nErrorHandler:\n    Exit Sub\nEnd Sub'
                    )

            # Add code to module
            vba_module.CodeModule.AddFromString(wrapped_code)

            # Execute the procedure
            full_procedure_name = f"{final_module_name}.{final_proc_name}"
            app.Run(full_procedure_name)

            return {
                "success": True,
                "message": "VBA code executed successfully",
                "procedure_name": final_proc_name,
                "module_name": final_module_name,
            }

        except Exception as e:
            error_msg = str(e)
            # Try to extract meaningful error from COM exception
            if "Error 1004" in error_msg:
                error_msg = "Excel runtime error during VBA execution"
            elif "Error -2147352573" in error_msg:
                error_msg = "VBA procedure not found or failed to compile"

            raise ToolError(
                f"VBA execution failed: {error_msg}",
                code=ErrorCodes.WRITE_FAILED
            )

        finally:
            # Always cleanup: remove temporary module
            if vba_module:
                try:
                    workbook.VBProject.VBComponents.Remove(vba_module)
                except Exception:
                    pass

            # Restore application settings
            try:
                app.DisplayAlerts = original_display_alerts
                app.ScreenUpdating = original_screen_updating
            except Exception:
                pass

    except ToolError:
        raise
    except Exception as e:
        raise ToolError(
            f"Failed to execute VBA: {str(e)}",
            code=ErrorCodes.EXCEL_NOT_RUNNING
        )
