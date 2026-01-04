"""Input validation utilities for Excel operations."""

import re
import logging
from typing import Any, List, Tuple, Union, Optional


class ToolError(Exception):
    """Exception raised for tool-related errors."""


logger = logging.getLogger(__name__)


def validate_cell_format(cell: str) -> bool:
    """Validate Excel cell reference format (e.g., A1, B5, Z100).

    Args:
        cell: Cell reference string to validate

    Returns:
        True if valid, False otherwise
    """
    if not cell or not isinstance(cell, str):
        return False

    # Match pattern: 1-3 letters followed by 1-7 digits
    pattern = r"^[A-Za-z]{1,3}[0-9]{1,7}$"
    return bool(re.match(pattern, cell))


def validate_workbook_name(name: str) -> bool:
    """Validate workbook name.

    Args:
        name: Workbook name to validate

    Returns:
        True if valid, False otherwise
    """
    if not name or not isinstance(name, str):
        return False

    # Excel will do thorough validation
    # Basic checks: not empty, reasonable length
    if len(name.strip()) == 0:
        return False

    if len(name) > 255:  # Excel max filename length
        return False

    # Check for invalid characters
    invalid_chars = ["<", ">", ":", '"', "/", "\\", "|", "?", "*"]
    if any(char in name for char in invalid_chars):
        return False

    return True


def validate_sheet_name(name: str) -> bool:
    """Validate Excel sheet name.

    Args:
        name: Sheet name to validate

    Returns:
        True if valid, False otherwise
    """
    if not name or not isinstance(name, str):
        return False

    if len(name.strip()) == 0:
        return False

    if len(name) > 31:  # Excel max sheet name length
        return False

    # Excel forbids these characters in sheet names
    invalid_chars = ["\\", "?", "*", "/", ":", "[", "]"]
    if any(char in name for char in invalid_chars):
        return False

    return True


def validate_range_format(range_str: str) -> bool:
    """Validate Excel range format.
    
    Supports:
    - Single cell: A1, F5, Z100
    - Standard range: A1:C5
    - Whole columns: A:C
    - Whole rows: 1:5
    """
    if not range_str or not isinstance(range_str, str):
        return False

    # Single cell A1, F5, Z100
    single_cell_pattern = r"^[A-Za-z]{1,3}[0-9]{1,7}$"
    # Standard A1:C5
    standard_pattern = r"^[A-Za-z]{1,3}[0-9]{1,7}:[A-Za-z]{1,3}[0-9]{1,7}$"
    # Whole columns A:C
    column_pattern = r"^[A-Za-z]{1,3}:[A-Za-z]{1,3}$"
    # Whole rows 1:5
    row_pattern = r"^[0-9]{1,7}:[0-9]{1,7}$"
    
    return bool(re.match(single_cell_pattern, range_str) or
                re.match(standard_pattern, range_str) or 
                re.match(column_pattern, range_str) or 
                re.match(row_pattern, range_str))


def parse_range(range_str: str) -> Tuple[str, str, str, str]:
    """Parse Excel range into start and end coordinates.

    Returns:
        Tuple of (start_col, start_row, end_col, end_row)
        Missing components are returned as empty strings.
        For single cells (e.g., F1), returns (F, 1, F, 1).
    """
    if not validate_range_format(range_str):
        raise ToolError(f"Invalid range format: '{range_str}'.")

    # Helper to split cell into alpha and digit parts
    def split_part(part):
        col = re.sub(r"[0-9]", "", part)
        row = re.sub(r"[A-Za-z]", "", part)
        return col, row

    # Handle single cell (no colon)
    if ":" not in range_str:
        col, row = split_part(range_str)
        return (col, row, col, row)  # Treat as 1x1 range

    parts = range_str.split(":")
    start_part, end_part = parts

    start_col, start_row = split_part(start_part)
    end_col, end_row = split_part(end_part)

    return (start_col, start_row, end_col, end_row)


def validate_data_dimensions(data: List[List[Any]], range_str: str) -> Tuple[bool, str]:
    """Validate data dimensions match range.

    Args:
        data: 2D array of values
        range_str: Excel range reference

    Returns:
        Tuple of (is_valid, error_message)
        - is_valid: True if dimensions match
        - error_message: Description of mismatch if invalid, empty string if valid
    """
    try:
        # First check data is rectangular
        is_rectangular, shape_error = validate_rectangular_data(data)
        if not is_rectangular:
            return (False, shape_error)

        start_col, start_row, end_col, end_row = parse_range(range_str)

        # Convert letters to numbers
        def col_letter_to_num(letter):
            if not letter:
                return None
            num = 0
            for char in letter:
                num = num * 26 + (ord(char.upper()) - ord("A") + 1)
            return num

        start_col_num = col_letter_to_num(start_col)
        end_col_num = col_letter_to_num(end_col)

        # Calculate expected sizes (handle whole row/column)
        if start_row and end_row:
            expected_rows = int(end_row) - int(start_row) + 1
        else:
            expected_rows = len(data)  # Assume it matches if unbound

        if start_col_num and end_col_num:
            expected_cols = end_col_num - start_col_num + 1
        else:
            expected_cols = len(data[0]) if data and data[0] else 0

        actual_rows = len(data)
        actual_cols = len(data[0]) if data and data[0] else 0

        if (start_row and end_row and actual_rows != expected_rows) or \
           (start_col_num and end_col_num and actual_cols != expected_cols):
            return (
                False,
                f"Data dimensions ({actual_rows}x{actual_cols}) do not match range '{range_str}'."
            )

        return (True, "")

    except Exception as e:
        return (False, f"Error validating dimensions: {str(e)}")


def validate_rectangular_data(data: List[List[Any]]) -> Tuple[bool, str]:
    """Validate that data is rectangular (all rows have same length).

    Args:
        data: 2D array of values

    Returns:
        Tuple of (is_valid, error_message)
        - is_valid: True if data is rectangular
        - error_message: Description of mismatch if invalid, empty string if valid

    Example:
        >>> validate_rectangular_data([[1,2], [3,4], [5,6]])
        (True, "")
        >>> validate_rectangular_data([[1,2,3], [4,5], [6,7,8]])
        (False, "Row 0 has 3 columns but row 1 has 2 columns")
    """
    if not data or not isinstance(data, list):
        return (False, "Data is empty or not a list")

    # Get column count from first non-empty row
    first_row_cols = None
    for row in data:
        if isinstance(row, list) and row:
            first_row_cols = len(row)
            break

    if first_row_cols is None:
        return (True, "")  # All empty rows is valid

    # Check all rows have the same column count
    for i, row in enumerate(data):
        if not isinstance(row, list):
            return (False, f"Row {i} is not a list (got {type(row).__name__})")

        actual_cols = len(row)
        if actual_cols != first_row_cols:
            return (
                False,
                f"Row {i} has {actual_cols} columns but row 0 has {first_row_cols} columns. "
                f"All rows must have the same number of columns."
            )

    return (True, "")


def validate_value_type(value: Any) -> bool:
    """Validate value type for Excel cell.

    Args:
        value: Value to validate

    Returns:
        True if valid, False otherwise
    """
    # Excel accepts: strings, numbers, dates, booleans
    valid_types = (str, int, float, bool)

    if value is None:
        return True  # Empty cells are fine

    if isinstance(value, valid_types):
        return True

    # Try to convert to string
    try:
        str(value)
        return True
    except Exception:
        return False


def get_cell_type(value: Any) -> str:
    """Get Excel-compatible type of a value.

    Args:
        value: Value to categorize

    Returns:
        Type string: 'string', 'number', 'boolean', 'empty'
    """
    if value is None:
        return "empty"

    if isinstance(value, bool):
        return "boolean"

    if isinstance(value, (int, float)):
        return "number"

    if isinstance(value, str):
        return "string"

    return "string"  # Default to string


# Excel error code mappings
EXCEL_ERROR_CODES = {
    -2146826281: ("#DIV/0!", "The formula or function used is divided by zero or empty cell."),
    -2146826246: ("#N/A", "A value is not available to a formula or function."),
    -2146826273: ("#VALUE!", "The wrong type of operand or argument is used."),
    -2146826288: ("#REF!", "A cell reference is not valid."),
    -2146826275: ("#NAME?", "Excel does not recognize a name or function."),
    -2146826280: ("#NUM!", "A formula or function contains invalid numeric values."),
    -2146826287: ("#NULL!", "A reference intersection is not valid."),
}


def get_excel_error_info(value: Any) -> tuple[str, str] | None:
    """Get Excel error code and description for an error code.

    Args:
        value: Cell value that might be an error code

    Returns:
        Tuple of (error_code, error_message) if value is an error code, None otherwise
        Example: ("#DIV/0!", "The formula or function used is divided by zero or empty cell.")
    """
    if isinstance(value, (int, float)):
        # Check if it's a negative Excel error code
        if value in EXCEL_ERROR_CODES:
            return EXCEL_ERROR_CODES[value]
    return None


def parse_reference_string(reference: str) -> List[str]:
    """Parse comma-separated Excel references (cells and ranges).

    Supports formats like:
    - Single cell: "A1"
    - Single range: "A1:C5"
    - Multiple cells: "A1,C3,E5"
    - Multiple ranges: "A1:B2,D4:E6"
    - Mixed: "A1,B2:D3,F5"

    Args:
        reference: Comma-separated reference string

    Returns:
        List of individual references (cells or ranges)

    Raises:
        ToolError: If reference string is invalid
    """
    if not reference or not isinstance(reference, str):
        raise ToolError("Reference must be a non-empty string")

    # Split by comma and strip whitespace
    refs = [ref.strip() for ref in reference.split(",") if ref.strip()]

    if not refs:
        raise ToolError("No valid references found in string")

    # Validate each reference
    for ref in refs:
        if not (validate_cell_format(ref) or validate_range_format(ref)):
            raise ToolError(
                f"Invalid reference '{ref}'. "
                "Expected format: A1, B5, Z100 (cells) or A1:C5, B2:D10 (ranges)"
            )

    return refs


# ============================================================================
# Filter Validation Functions
# ============================================================================

VALID_FILTER_OPERATORS = {
    # Comparison operators
    "=", "!=", ">", ">=", "<", "<=",
    # String search operators
    "contains", "starts_with", "ends_with", "regex",
    # Range operators
    "between", "in", "not_in",
    # Empty check operators
    "is_empty", "not_empty",
}

VALID_LOGICAL_OPERATORS = {"and", "or", "not"}

VALID_COLUMN_REFERENCE_TYPES = {"letter", "name", "index"}


def validate_filter_operator(operator: str) -> Tuple[bool, str]:
    """Validate filter operator.

    Args:
        operator: Operator string to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not operator or not isinstance(operator, str):
        return (False, "Operator must be a non-empty string")

    if operator.lower() not in VALID_FILTER_OPERATORS:
        return (
            False,
            f"Invalid operator '{operator}'. Valid operators: {', '.join(sorted(VALID_FILTER_OPERATORS))}"
        )

    return (True, "")


def validate_logical_operator(operator: str) -> Tuple[bool, str]:
    """Validate logical operator (AND/OR/NOT).

    Args:
        operator: Logical operator string to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not operator or not isinstance(operator, str):
        return (False, "Logical operator must be a non-empty string")

    if operator.lower() not in VALID_LOGICAL_OPERATORS:
        return (
            False,
            f"Invalid logical operator '{operator}'. Valid operators: {', '.join(sorted(VALID_LOGICAL_OPERATORS))}"
        )

    return (True, "")


def validate_column_reference(column_ref: Any) -> Tuple[bool, str]:
    """Validate column reference structure.

    Args:
        column_ref: Column reference (dict, ColumnReference, or None)

    Returns:
        Tuple of (is_valid, error_message)
    """
    # None is valid (means "all columns")
    if column_ref is None:
        return (True, "")

    # Must be a dict
    if not isinstance(column_ref, dict):
        return (False, f"Column reference must be a dict or None, got {type(column_ref).__name__}")

    # Check for required keys
    if "type" not in column_ref:
        return (False, "Column reference missing 'type' key")

    if "value" not in column_ref:
        return (False, "Column reference missing 'value' key")

    # Validate type
    ref_type = column_ref["type"]
    if ref_type not in VALID_COLUMN_REFERENCE_TYPES:
        return (
            False,
            f"Invalid column reference type '{ref_type}'. Valid types: {', '.join(sorted(VALID_COLUMN_REFERENCE_TYPES))}"
        )

    # Validate value based on type
    value = column_ref["value"]

    if ref_type == "letter":
        if not isinstance(value, str) or not re.match(r"^[A-Za-z]+$", value):
            return (False, f"Column letter must be alphabetic (e.g., 'A', 'AB'), got: {value}")

    elif ref_type == "name":
        if not isinstance(value, str) or not value.strip():
            return (False, "Column name must be a non-empty string")

    elif ref_type == "index":
        if not isinstance(value, (int, float, str)):
            return (False, f"Column index must be a number, got: {type(value).__name__}")
        try:
            idx = int(value)
            if idx < 1:
                return (False, f"Column index must be >= 1, got: {idx}")
        except (ValueError, TypeError):
            return (False, f"Column index must be a valid integer, got: {value}")

    return (True, "")


def validate_single_filter(filter_dict: Any) -> Tuple[bool, str]:
    """Validate a single filter structure.

    Args:
        filter_dict: Filter dictionary to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not filter_dict or not isinstance(filter_dict, dict):
        return (False, "Filter must be a non-empty dictionary")

    # Check for operator
    if "operator" not in filter_dict:
        return (False, "Filter missing 'operator' key")

    # Validate operator
    is_valid, error_msg = validate_filter_operator(filter_dict["operator"])
    if not is_valid:
        return (False, error_msg)

    # Validate column reference (optional for "all columns" search)
    if "column" in filter_dict:
        is_valid, error_msg = validate_column_reference(filter_dict["column"])
        if not is_valid:
            return (False, f"Invalid column reference: {error_msg}")

    # Check for required value based on operator
    operator = filter_dict["operator"].lower()
    if operator not in ("is_empty", "not_empty") and "value" not in filter_dict:
        return (False, f"Operator '{operator}' requires 'value' key")

    # For BETWEEN, check for value2
    if operator == "between" and "value2" not in filter_dict:
        return (False, "Operator 'between' requires 'value2' key")

    return (True, "")


def validate_filter_group(filter_dict: Any) -> Tuple[bool, str]:
    """Validate a filter group structure.

    Args:
        filter_dict: Filter group dictionary to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not filter_dict or not isinstance(filter_dict, dict):
        return (False, "Filter group must be a non-empty dictionary")

    # Check for operator or logical (both are accepted)
    logical_op = filter_dict.get("operator") or filter_dict.get("logical")
    if not logical_op:
        return (False, "Filter group missing 'operator' key")

    # Validate logical operator
    is_valid, error_msg = validate_logical_operator(logical_op)
    if not is_valid:
        return (False, error_msg)

    # Check for conditions or filters (both are accepted)
    conditions = filter_dict.get("conditions") or filter_dict.get("filters")
    if conditions is None:
        return (False, "Filter group missing 'conditions' key")

    if not isinstance(conditions, list):
        return (False, f"Conditions must be a list, got: {type(conditions).__name__}")

    if not conditions:
        return (False, "Filter group must have at least one condition")

    # Validate each condition
    for i, cond in enumerate(conditions):
        # Check if it's a filter group (has 'conditions', 'filters', or 'logical')
        if isinstance(cond, dict):
            if "conditions" in cond or "filters" in cond or "logical" in cond:
                is_valid, error_msg = validate_filter_group(cond)
                if not is_valid:
                    return (False, f"Condition {i + 1}: {error_msg}")
            else:
                is_valid, error_msg = validate_single_filter(cond)
                if not is_valid:
                    return (False, f"Condition {i + 1}: {error_msg}")
        else:
            return (False, f"Condition {i + 1}: must be a dictionary")

    return (True, "")


def validate_filter_structure(filter_dict: Any) -> Tuple[bool, str]:
    """Validate any filter structure (single filter or filter group).

    Args:
        filter_dict: Filter dictionary to validate

    Returns:
        Tuple of (is_valid, error_message)

    Example:
        >>> validate_filter_structure({"column": {"type": "letter", "value": "A"}, "operator": ">", "value": 100})
        (True, "")
        >>> validate_filter_structure({"operator": "and", "conditions": [...]})
        (True, "")
    """
    if not filter_dict or not isinstance(filter_dict, dict):
        return (False, "Filter must be a non-empty dictionary")

    # Check if it's a filter group (has conditions, filters, or logical)
    if "conditions" in filter_dict or "filters" in filter_dict or "logical" in filter_dict:
        return validate_filter_group(filter_dict)
    else:
        return validate_single_filter(filter_dict)


# ============================================================================
# Formula Validation Functions
# ============================================================================


VALID_EXCEL_FUNCTIONS = {
    # Math functions
    "SUM", "AVERAGE", "COUNT", "COUNTA", "COUNTBLANK", "MAX", "MIN",
    "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS",
    "SUMPRODUCT",
    "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS", "SQRT", "POWER", "MOD",
    "INT", "CEILING", "FLOOR", "RAND", "RANDBETWEEN", "SEQUENCE", "RANDARRAY",
    # Text functions
    "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "PROPER",
    "CONCATENATE", "CONCAT", "TEXTJOIN", "TEXT", "VALUE", "FIND", "SEARCH",
    "SUBSTITUTE", "REPLACE", "REPT", "CHAR", "CODE", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER",
    # Lookup functions
    "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP", "LOOKUP", "XMATCH",
    "OFFSET", "INDIRECT", "ROW", "COLUMN", "ROWS", "COLUMNS", "CHOOSECOLS", "CHOOSEROWS",
    "FILTER", "SORT", "SORTBY", "UNIQUE", "TRANSPOSE",
    # Logical functions
    "IF", "IFS", "AND", "OR", "NOT", "XOR", "TRUE", "FALSE",
    "IFERROR", "IFNA", "SWITCH", "CHOOSE", "LET", "LAMBDA",
    "MAP", "REDUCE", "SCAN", "MAKEARRAY", "BYROW", "BYCOL", "ISOMITTED",
    # Date functions
    "DATE", "TODAY", "NOW", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND",
    "WEEKDAY", "WEEKNUM", "DATEVALUE", "TIMEVALUE", "EDATE", "EOMONTH",
    "NETWORKDAYS", "WORKDAY", "DATEDIF",
    # Info functions
    "ISBLANK", "ISERROR", "ISNA", "ISNUMBER", "ISTEXT", "ISLOGICAL", "ISFORMULA",
    "TYPE", "N", "NA", "INFO", "CELL",
    # Financial functions
    "PMT", "PV", "FV", "RATE", "NPER", "NPV", "IRR",
    # Statistical functions
    "STDEV", "STDEVP", "VAR", "VARP", "MEDIAN", "MODE", "LARGE", "SMALL",
    "PERCENTILE", "QUARTILE", "RANK", "CORREL", "COVAR",
}


def validate_cell_reference(cell: str) -> bool:
    """Validate a cell reference (e.g., A1, $B$5, Sheet1!A1).
    
    Args:
        cell: Cell reference to validate
        
    Returns:
        True if valid cell reference
    """
    if not cell or not isinstance(cell, str):
        return False
    
    # Remove sheet reference if present
    if "!" in cell:
        parts = cell.split("!")
        if len(parts) != 2:
            return False
        cell = parts[1]
    
    # Remove $ signs (absolute reference markers)
    cell = cell.replace("$", "")
    
    # Match pattern: 1-3 letters followed by 1-7 digits
    pattern = r"^[A-Za-z]{1,3}[0-9]{1,7}$"
    return bool(re.match(pattern, cell))


def validate_formula_sync(formula: str) -> dict:
    """Validate Excel formula syntax without applying it.
    
    Checks for:
    - Proper formula start (=)
    - Balanced parentheses
    - Valid function names (including modern 365 functions)
    - Valid cell references
    - Proper string quoting
    - Ambiguous ranges
    
    Args:
        formula: Formula string to validate
        
    Returns:
        Dictionary with validation result:
        {
            "valid": bool,
            "error": str or None,
            "suggestion": str or None,
            "warnings": [
                {
                    "type": str,      # invalid_function, suspicious_range, syntax, logic
                    "severity": str,  # low, medium, high
                    "message": str,
                    "location": str or None
                }
            ],
            "functions_used": list of str,
        }
    """
    result = {
        "valid": True,
        "error": None,
        "suggestion": None,
        "warnings": [],
        "functions_used": [],
    }
    
    if not formula or not isinstance(formula, str):
        result["valid"] = False
        result["error"] = "Formula must be a non-empty string"
        return result
    
    formula = formula.strip()
    
    # Check if formula starts with =
    if not formula.startswith("="):
        # Suggest correction
        result["valid"] = False
        result["error"] = "Formula must start with '='"
        result["suggestion"] = f"Did you mean: ={formula}?"
        return result
    
    # Remove the leading =
    formula_body = formula[1:]
    
    if not formula_body:
        result["valid"] = False
        result["error"] = "Formula body is empty after '='"
        return result
    
    # Check for balanced parentheses
    paren_count = 0
    in_string = False
    string_char = None
    
    for i, char in enumerate(formula_body):
        # Track string literals
        if char in ('"', "'") and (i == 0 or formula_body[i-1] != "\\"):
            if not in_string:
                in_string = True
                string_char = char
            elif char == string_char:
                in_string = False
                string_char = None
            continue
        
        if not in_string:
            if char == "(":
                paren_count += 1
            elif char == ")":
                paren_count -= 1
                
            if paren_count < 0:
                result["valid"] = False
                result["error"] = f"Unbalanced parentheses: extra ')' at position {i + 2}"
                return result
    
    if paren_count != 0:
        result["valid"] = False
        result["error"] = f"Unbalanced parentheses: {abs(paren_count)} unclosed '('"
        return result
    
    # Check for unclosed string
    if in_string:
        result["valid"] = False
        result["error"] = "Unclosed string literal"
        return result
    
    # Extract and validate function names
    function_pattern = r'\b([A-Z][A-Z0-9_]*)\s*\('
    functions = re.findall(function_pattern, formula_body.upper())
    
    for func in functions:
        if func not in VALID_EXCEL_FUNCTIONS:
            # Check if it might be a named range or just unknown
            # Named ranges usually don't have '(' immediately after, but in some contexts they might
            # Simple heuristic: if it looks like a function call but isn't known
            
            # Check against range pattern (A1, AA123) to avoid false positives for multiplication e.g. AB(
            is_range_like = re.match(r'^[A-Z]{1,3}[0-9]+$', func)
            
            if not is_range_like:
                result["warnings"].append({
                    "type": "unknown_function",
                    "severity": "medium",
                    "message": f"Unknown function '{func}'. Use generic functions or check for typos.",
                    "function": func
                })
        else:
            if func not in result["functions_used"]:
                result["functions_used"].append(func)
    
    # Range Validation (Advanced)
    # Check for whole column references in potential array contexts
    # Pattern: A:A or A:C
    whole_col_pattern = r'\b[A-Za-z]{1,3}:[A-Za-z]{1,3}\b'
    whole_cols = re.findall(whole_col_pattern, formula_body)
    
    # Functions that are dangerous with whole columns (slow calculation)
    dangerous_funcs_with_whole_cols = {"SUMPRODUCT", "FILTER", "SORT", "UNIQUE"}
    used_dangerous_func = any(f in result["functions_used"] for f in dangerous_funcs_with_whole_cols)
    
    if whole_cols and used_dangerous_func:
        result["warnings"].append({
            "type": "performance_risk",
            "severity": "medium",
            "message": f"Performance risk: Using whole column references ({whole_cols[0]}) with array functions references >1M rows.",
        })
    elif whole_cols:
        # Just a low severity warning for general awareness
         result["warnings"].append({
            "type": "range_suspicious",
            "severity": "low",
            "message": f"Note: Whole column reference {whole_cols[0]} used. Ensure this is intentional.",
        })
        
        
    # Check for common syntax errors
    # Double operators
    if re.search(r'[+\-*/^]{2,}', formula_body):
        result["valid"] = False
        result["error"] = "Invalid: consecutive operators"
        return result
    
    # Operator at end (except for ranges)
    if re.search(r'[+\-*/^=<>]$', formula_body) and not formula_body.endswith(":"):
        result["valid"] = False
        result["error"] = "Formula cannot end with an operator"
        return result
    
    # Empty parentheses
    if "()" in formula_body:
        # Check if it's a valid no-argument function
        no_arg_functions = {"NOW", "TODAY", "TRUE", "FALSE", "NA", "RAND", "ROW", "COLUMN"}
        match = re.search(r'(\w+)\(\)', formula_body.upper())
        if match and match.group(1) not in no_arg_functions:
             result["warnings"].append({
                "type": "syntax",
                "severity": "low",
                "message": f"Empty parentheses in {match.group(1)}(). Verify if arguments are required.",
            })
    
    return result


