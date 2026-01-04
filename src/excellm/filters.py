"""Filter data structures and engine for Excel data filtering.

This module provides in-memory filtering capabilities for 2D Excel data arrays.
Supports nested boolean logic (AND/OR/NOT), equations, and advanced operators.
"""

import re
import logging
from datetime import datetime
from enum import Enum
from typing import Any, Optional, Union, List, Tuple

logger = logging.getLogger(__name__)


class FilterOperator(str, Enum):
    """Supported comparison operators."""

    # Comparison operators
    EQ = "="
    NE = "!="
    GT = ">"
    GTE = ">="
    LT = "<"
    LTE = "<="

    # String search operators
    CONTAINS = "contains"
    STARTS_WITH = "starts_with"
    ENDS_WITH = "ends_with"
    REGEX = "regex"

    # Range operators
    BETWEEN = "between"
    IN = "in"
    NOT_IN = "not_in"

    # Empty check operators
    IS_EMPTY = "is_empty"
    NOT_EMPTY = "not_empty"


class LogicalOperator(str, Enum):
    """Logical operators for combining filters."""

    AND = "and"
    OR = "or"
    NOT = "not"


class ColumnReference:
    """Reference to a column by letter, name, or index."""

    def __init__(self, type: str, value: Union[str, int]):
        """
        Args:
            type: Reference type - "letter", "name", or "index"
            value: Column letter (A), name (Price), or index (1)
        """
        if type not in ("letter", "name", "index"):
            raise ValueError(
                f"Invalid column reference type: {type}. Must be 'letter', 'name', or 'index'"
            )
        self.type = type
        self.value = value

    def resolve(self, headers: Optional[List[str]] = None) -> int:
        """
        Resolve column reference to zero-based index.

        Args:
            headers: List of column headers (required for name-based reference)

        Returns:
            Zero-based column index
        """
        if self.type == "letter":
            return self._letter_to_index(str(self.value))
        elif self.type == "index":
            return int(self.value) - 1  # Convert 1-based to 0-based
        elif self.type == "name":
            if headers is None:
                raise ValueError("Headers required for name-based column reference")
            for i, header in enumerate(headers):
                if str(header).strip().lower() == str(self.value).strip().lower():
                    return i
            raise ValueError(f"Column '{self.value}' not found in headers: {headers}")
        return 0

    @staticmethod
    def _letter_to_index(letter: str) -> int:
        """Convert Excel column letter to zero-based index (A=0, B=1, AA=26)."""
        letter = str(letter).upper()
        result = 0
        for char in letter:
            if not char.isalpha():
                raise ValueError(f"Invalid column letter: {letter}")
            result = result * 26 + (ord(char) - ord("A") + 1)
        return result - 1

    def to_dict(self) -> dict:
        """Convert to dictionary representation."""
        return {"type": self.type, "value": self.value}

    @classmethod
    def from_dict(cls, data: dict) -> "ColumnReference":
        """Create from dictionary representation."""
        return cls(type=data["type"], value=data["value"])

    def __repr__(self) -> str:
        return f"ColumnReference({self.type}, {self.value})"


class SingleFilter:
    """A single filter condition."""

    def __init__(
        self,
        column: Optional[Union[ColumnReference, dict]] = None,
        operator: Union[FilterOperator, str] = FilterOperator.CONTAINS,
        value: Optional[Any] = None,
        value2: Optional[Any] = None,
        strict_type: bool = False,
    ):
        """
        Args:
            column: ColumnReference, dict, or None (None = search all columns)
            operator: FilterOperator or string
            value: Comparison value
            value2: Second value for BETWEEN operator
            strict_type: If True, disable type coercion for = operator (exact type matching only)
        """
        if isinstance(column, dict):
            column = ColumnReference.from_dict(column)
        elif column is not None and not isinstance(column, ColumnReference):
            raise ValueError(f"column must be ColumnReference, dict, or None, got {type(column)}")

        self.column = column
        self.operator = FilterOperator(operator) if isinstance(operator, str) else operator
        self.value = value
        self.value2 = value2
        self.strict_type = strict_type

    def to_dict(self) -> dict:
        """Convert to dictionary representation."""
        result = {
            "operator": self.operator.value,
            "value": self.value,
            "strict_type": self.strict_type,
        }
        if self.column is not None:
            result["column"] = self.column.to_dict()
        if self.value2 is not None:
            result["value2"] = self.value2
        return result

    @classmethod
    def from_dict(cls, data: dict) -> "SingleFilter":
        """Create from dictionary representation."""
        return cls(
            column=data.get("column"),
            operator=data.get("operator", "contains"),  # Changed from "=" to "contains"
            value=data.get("value"),
            value2=data.get("value2"),
            strict_type=data.get("strict_type", False),
        )

    def __repr__(self) -> str:
        col = self.column if self.column else "all_columns"
        return f"SingleFilter({col}, {self.operator.value}, {self.value})"


class FilterGroup:
    """A group of filters with logical AND/OR/NOT operators."""

    def __init__(
        self,
        operator: Union[LogicalOperator, str] = LogicalOperator.AND,
        conditions: Optional[List[Union[dict, SingleFilter, "FilterGroup"]]] = None,
    ):
        """
        Args:
            operator: LogicalOperator (and/or/not)
            conditions: List of SingleFilter or FilterGroup objects
        """
        self.operator = LogicalOperator(operator) if isinstance(operator, str) else operator
        self.conditions: List[Union[SingleFilter, FilterGroup]] = []

        if conditions:
            for cond in conditions:
                if isinstance(cond, dict):
                    # Detect if it's a SingleFilter or FilterGroup
                    if "conditions" in cond or ("operator" in cond and "logical" in cond):
                        self.conditions.append(FilterGroup.from_dict(cond))
                    else:
                        self.conditions.append(SingleFilter.from_dict(cond))
                elif isinstance(cond, (SingleFilter, FilterGroup)):
                    self.conditions.append(cond)
                else:
                    raise ValueError(f"Invalid condition type: {type(cond)}")

    def to_dict(self) -> dict:
        """Convert to dictionary representation."""
        return {
            "operator": self.operator.value,
            "conditions": [c.to_dict() if hasattr(c, "to_dict") else c for c in self.conditions],
        }

    @classmethod
    def from_dict(cls, data: dict) -> "FilterGroup":
        """Create from dictionary representation."""
        return cls(
            operator=data.get("operator", "and"),
            conditions=data.get("conditions", data.get("filters", [])),
        )

    def __repr__(self) -> str:
        return f"FilterGroup({self.operator.value}, {len(self.conditions)} conditions)"


class FilterEngine:
    """Engine for applying filters to 2D Excel data arrays."""

    def __init__(self):
        self._cache_column_types = {}
        self._warnings: List[str] = []
        self._trim = True

    def apply_filter(
        self,
        data: List[List[Any]],
        filters: Union[dict, FilterGroup, SingleFilter],
        headers: Optional[List[str]] = None,
        trim: bool = True,
        start_row: int = 1,
        start_col: int = 1,
    ) -> dict:
        """
        Apply filter to 2D data array.

        Args:
            data: 2D array of values (including header row if present)
            filters: FilterGroup, SingleFilter, or dict representation
            headers: List of column headers (extracted from data[0] if not provided)
            trim: If True, trim whitespace from strings before comparison (default: True)

        Returns:
            Dictionary with filtered data and metadata:
            {
                "success": True,
                "data": filtered 2D array,
                "rows_filtered": number of data rows returned,
                "rows_removed": number of rows filtered out,
                "columns": column names (if headers provided),
                "filter_applied": human-readable filter description,
                "warnings": list of warning messages (if any)
            }
        """
        # Set trim for this operation
        self._trim = trim

        # Clear previous warnings
        self._warnings = []

        if not data or not isinstance(data, list):
            return {
                "success": True,
                "data": [],
                "rows_filtered": 0,
                "rows_removed": 0,
                "columns": headers or [],
                "filter_applied": str(filters),
                "cell_locations": [],
                "warnings": [],
            }

        # Parse filters if dict
        if isinstance(filters, dict):
            filters = self._parse_filter_dict(filters)

        # Extract headers from first row if not provided
        has_header = headers is not None
        header_row = None
        data_rows = data

        if has_header and headers:
            header_row = headers
            data_rows = data[1:] if len(data) > 1 else []
        elif headers is None and data:
            # Assume first row is headers
            header_row = data[0] if data else []
            data_rows = data[1:] if len(data) > 1 else []
            headers = header_row

        # Validate columns in filters against headers
        if headers:
            self._validate_filter_columns(filters, headers, len(data[0]) if data else 0)

        # Apply filter to each row, tracking indices
        filtered_rows = []
        matched_indices = []
        data_start_index = 1 if has_header else 0

        for i, row in enumerate(data_rows):
            col_indices = self._matches_filter(row, filters, headers)
            if col_indices:
                filtered_rows.append(row)
                # Track all matching cells in this row
                actual_row_offset = data_start_index + i
                for c_idx in col_indices:
                    matched_indices.append((actual_row_offset, c_idx))

        # Check for zero matches warning
        if len(filtered_rows) == 0 and len(data_rows) > 0:
            self._warnings.append(
                f"Filter returned zero matching rows from {len(data_rows)} data rows. "
                "Check that column names exist and values match the expected format."
            )

        # Reattach header if present
        result_data = []
        if header_row is not None:
            result_data.append(header_row)
        result_data.extend(filtered_rows)

        result = {
            "success": True,
            "data": result_data,
            "rows_filtered": len(filtered_rows),
            "rows_removed": len(data_rows) - len(filtered_rows),
            "columns": list(headers) if headers else [],
            "filter_applied": str(filters),
            "cell_locations": self._get_cell_locations(matched_indices, start_row, start_col),
        }

        if self._warnings:
            result["warnings"] = self._warnings

        return result

    def _validate_filter_columns(
        self,
        filter_obj: Union[FilterGroup, SingleFilter],
        headers: List[str],
        num_columns: int,
    ) -> None:
        """Validate that columns referenced in filters exist in headers/range."""
        if isinstance(filter_obj, SingleFilter):
            self._validate_single_filter_columns(filter_obj, headers, num_columns)
        elif isinstance(filter_obj, FilterGroup):
            for condition in filter_obj.conditions:
                self._validate_filter_columns(condition, headers, num_columns)

    def _validate_single_filter_columns(
        self,
        filter_obj: SingleFilter,
        headers: List[str],
        num_columns: int,
    ) -> None:
        """Validate column reference in a single filter."""
        if filter_obj.column is None:
            return  # Search all columns - no validation needed

        col_ref = filter_obj.column
        headers_lower = [h.lower() if isinstance(h, str) else str(h).lower() for h in headers]

        if col_ref.type == "name":
            # Check if column name exists in headers
            col_name_lower = str(col_ref.value).lower()
            if col_name_lower not in headers_lower:
                similar_names = [
                    h
                    for h in headers
                    if str(h).lower().replace(" ", "") == col_name_lower.replace(" ", "")
                ]
                if similar_names:
                    self._warnings.append(
                        f"Column '{col_ref.value}' not found. Did you mean '{similar_names[0]}' (with different spacing)?"
                    )
                else:
                    # Find closest match
                    closest = self._find_closest_match(col_ref.value, headers)
                    if closest:
                        self._warnings.append(
                            f"Column '{col_ref.value}' not found. Did you mean '{closest}'?"
                        )
                    else:
                        self._warnings.append(
                            f"Column '{col_ref.value}' not found in headers. Available columns: {', '.join(headers[:10])}{'...' if len(headers) > 10 else ''}"
                        )

        elif col_ref.type == "index":
            # Check if column index is within range
            try:
                idx = int(col_ref.value)
                if idx < 1 or idx > num_columns:
                    self._warnings.append(
                        f"Column index {idx} is out of range. Data has {num_columns} columns (valid range: 1-{num_columns})."
                    )
            except (ValueError, TypeError):
                pass  # Invalid index will be caught during resolution

    def _find_closest_match(self, target: str, options: List[str]) -> Optional[str]:
        """Find the closest matching string from options (simple edit distance)."""
        target_lower = target.lower()
        best_match = None
        best_score = 0

        for opt in options:
            opt_lower = str(opt).lower()
            # Check for contains
            if target_lower in opt_lower or opt_lower in target_lower:
                return opt
            # Simple similarity check
            common_chars = sum(1 for c in target_lower if c in opt_lower)
            score = common_chars / max(len(target_lower), len(opt_lower))
            if score > best_score and score > 0.5:
                best_score = score
                best_match = opt

        return best_match

    def _get_cell_locations(
        self, matched_indices: List[Tuple[int, Optional[int]]], start_row: int, start_col: int
    ) -> List[str]:
        """Convert row/col offset indices to Excel cell references (A1 format)."""
        locations = []
        for row_idx, col_idx in matched_indices:
            # Excel is 1-indexed, start_row/col are absolute Excel coordinates
            abs_row = start_row + row_idx
            # If no specific column matched, default to the start of the range
            actual_col = start_col + (col_idx if col_idx is not None else 0)
            col_letter = self._number_to_column(actual_col)
            locations.append(f"{col_letter}{abs_row}")
        return locations

    def _number_to_column(self, n: int) -> str:
        """Convert column number to Excel column letter (e.g., 1=A, 27=AA)."""
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _parse_filter_dict(self, filter_dict: dict) -> Union[FilterGroup, SingleFilter]:
        """Parse dictionary into FilterGroup or SingleFilter."""
        # Check if it's a FilterGroup (has 'conditions' or 'filters' or 'logical')
        if "conditions" in filter_dict or "filters" in filter_dict or "logical" in filter_dict:
            return FilterGroup.from_dict(filter_dict)
        # Otherwise it's a SingleFilter
        return SingleFilter.from_dict(filter_dict)

    def _matches_filter(
        self,
        row: List[Any],
        filter_obj: Union[FilterGroup, SingleFilter],
        headers: Optional[List[str]] = None,
    ) -> List[Optional[int]]:
        """Check if row matches the filter and return matching column indices.
        Returns empty list if no match.
        """
        if isinstance(filter_obj, SingleFilter):
            return self._matches_single_filter(row, filter_obj, headers)
        elif isinstance(filter_obj, FilterGroup):
            return self._matches_filter_group(row, filter_obj, headers)
        return []

    def _matches_filter_group(
        self,
        row: List[Any],
        filter_group: FilterGroup,
        headers: Optional[List[str]] = None,
    ) -> List[Optional[int]]:
        """Check if row matches a filter group (AND/OR/NOT)."""
        operator = filter_group.operator
        
        if not filter_group.conditions:
            return [None]  # Empty group matches everything

        if operator == LogicalOperator.AND:
            all_indices = []
            for cond in filter_group.conditions:
                indices = self._matches_filter(row, cond, headers)
                if not indices:
                    return []
                all_indices.extend(indices)
            # Return unique indices, maintaining None if present
            return list(dict.fromkeys(all_indices))

        elif operator == LogicalOperator.OR:
            any_indices = []
            matched = False
            for cond in filter_group.conditions:
                indices = self._matches_filter(row, cond, headers)
                if indices:
                    matched = True
                    any_indices.extend(indices)
            if matched:
                return list(dict.fromkeys(any_indices)) or [None]
            return []

        elif operator == LogicalOperator.NOT:
            if not filter_group.conditions:
                return [None]
            indices = self._matches_filter(row, filter_group.conditions[0], headers)
            if not indices:
                return [None]
            return []

        return []

    def _matches_single_filter(
        self,
        row: List[Any],
        filter_obj: SingleFilter,
        headers: Optional[List[str]] = None,
    ) -> List[Optional[int]]:
        """Check if row matches a single filter condition."""
        # If no column specified, search across all columns
        if filter_obj.column is None:
            return self._matches_any_column(row, filter_obj)

        # Resolve column index
        try:
            col_idx = filter_obj.column.resolve(headers)
        except (ValueError, IndexError):
            return []  # Column not found

        if col_idx >= len(row):
            return []  # Column doesn't exist in this row

        cell_value = row[col_idx]
        if self._compare_values(cell_value, filter_obj, self._trim):
            return [col_idx]
        return []

    def _matches_any_column(self, row: List[Any], filter_obj: SingleFilter) -> List[Optional[int]]:
        """Check if any column in the row matches the filter."""
        matched_cols = []
        for i, cell_value in enumerate(row):
            if self._compare_values(cell_value, filter_obj, self._trim):
                matched_cols.append(i)
        return matched_cols

    def _compare_values(self, cell_value: Any, filter_obj: SingleFilter, trim: bool = True) -> bool:
        """Compare a cell value against a filter."""
        operator = filter_obj.operator
        filter_value = filter_obj.value
        filter_value2 = filter_obj.value2
        strict_type = filter_obj.strict_type

        # Handle empty checks first
        if operator == FilterOperator.IS_EMPTY:
            return cell_value in [None, "", " ", []]
        elif operator == FilterOperator.NOT_EMPTY:
            return cell_value not in [None, "", " ", []]

        # Convert cell value to string for comparison
        cell_str = str(cell_value) if cell_value is not None else ""

        # Numeric comparisons
        if operator in (
            FilterOperator.GT,
            FilterOperator.GTE,
            FilterOperator.LT,
            FilterOperator.LTE,
        ):
            cell_num = self._to_number(cell_value)
            filter_num = self._to_number(filter_value)
            if cell_num is None or filter_num is None:
                return False
            if operator == FilterOperator.GT:
                return cell_num > filter_num
            elif operator == FilterOperator.GTE:
                return cell_num >= filter_num
            elif operator == FilterOperator.LT:
                return cell_num < filter_num
            elif operator == FilterOperator.LTE:
                return cell_num <= filter_num

        # Equality operators
        elif operator == FilterOperator.EQ:
            return self._values_equal(cell_value, filter_value, strict_type=strict_type)
        elif operator == FilterOperator.NE:
            return not self._values_equal(cell_value, filter_value, strict_type=strict_type)

        # String search operators
        elif operator == FilterOperator.CONTAINS:
            return self._string_contains(cell_str, filter_value)
        elif operator == FilterOperator.STARTS_WITH:
            return self._string_starts_with(cell_str, filter_value)
        elif operator == FilterOperator.ENDS_WITH:
            return self._string_ends_with(cell_str, filter_value)
        elif operator == FilterOperator.REGEX:
            return self._regex_match(cell_str, filter_value)

        # Range operators
        elif operator == FilterOperator.BETWEEN:
            val_num = self._to_number(cell_value)
            min_num = self._to_number(filter_value)
            max_num = self._to_number(filter_value2)
            if val_num is None or min_num is None or max_num is None:
                return False
            return min_num <= val_num <= max_num
        elif operator == FilterOperator.IN:
            if isinstance(filter_value, (list, tuple)):
                return self._values_equal(
                    cell_value, filter_value, strict_type=strict_type, is_in=True
                )
            return False
        elif operator == FilterOperator.NOT_IN:
            if isinstance(filter_value, (list, tuple)):
                return not self._values_equal(
                    cell_value, filter_value, strict_type=strict_type, is_in=True
                )
            return True

        return False

    def _values_equal(
        self, cell_value: Any, filter_value: Any, strict_type: bool = False, is_in: bool = False
    ) -> bool:
        """Check if values are equal with optional automatic type coercion for numeric values."""
        # Always trim strings before comparison
        if isinstance(cell_value, str):
            cell_value = cell_value.strip()
        if isinstance(filter_value, str):
            filter_value = filter_value.strip()

        if is_in and isinstance(filter_value, (list, tuple)):
            # Check if cell_value is in the list
            for item in filter_value:
                if self._values_equal(cell_value, item, strict_type=strict_type):
                    return True
            return False

        # Automatic type coercion for numeric comparisons (skip if strict_type=True)
        if not strict_type:
            cell_is_numeric = isinstance(cell_value, (int, float))
            filter_is_numeric = isinstance(filter_value, (int, float))

            if cell_is_numeric and isinstance(filter_value, str):
                # Cell is numeric, filter is string - try to convert filter to numeric
                filter_num = self._to_number(filter_value)
                if filter_num is not None:
                    return float(cell_value) == filter_num
            elif filter_is_numeric and isinstance(cell_value, str):
                # Filter is numeric, cell is string - try to convert cell to numeric
                cell_num = self._to_number(cell_value)
                if cell_num is not None:
                    return cell_num == float(filter_value)
            elif isinstance(cell_value, str) and isinstance(filter_value, str):
                # Both are strings - try to convert both to numeric for flexible matching
                cell_num = self._to_number(cell_value)
                filter_num = self._to_number(filter_value)
                if cell_num is not None and filter_num is not None:
                    return cell_num == filter_num

        # Direct comparison (always case-insensitive for strings)
        if isinstance(cell_value, str) and isinstance(filter_value, str):
            return cell_value.lower() == filter_value.lower()
        return str(cell_value) == str(filter_value)

    def _string_contains(self, haystack: str, needle: Any) -> bool:
        """Check if string contains substring (always case-insensitive and trimmed)."""
        if haystack is None or needle is None:
            return False
        hs = str(haystack).strip()
        n = str(needle).strip()
        return n.lower() in hs.lower()

    def _string_starts_with(self, haystack: str, prefix: Any) -> bool:
        """Check if string starts with prefix (always case-insensitive and trimmed)."""
        if haystack is None or prefix is None:
            return False
        hs = str(haystack).strip()
        p = str(prefix).strip()
        return hs.lower().startswith(p.lower())

    def _string_ends_with(self, haystack: str, suffix: Any) -> bool:
        """Check if string ends with suffix (always case-insensitive and trimmed)."""
        if haystack is None or suffix is None:
            return False
        hs = str(haystack).strip()
        s = str(suffix).strip()
        return hs.lower().endswith(s.lower())

    def _regex_match(self, haystack: str, pattern: Any) -> bool:
        """Check if string matches regex pattern (always case-insensitive and trimmed)."""
        if haystack is None or pattern is None:
            return False
        try:
            hs = str(haystack).strip()
            return bool(re.search(str(pattern), hs, re.IGNORECASE))
        except re.error:
            return False

    def _to_number(self, value: Any) -> Optional[float]:
        """Convert value to number if possible."""
        if value is None:
            return None
        try:
            # Handle common numeric formats
            if isinstance(value, (int, float)):
                return float(value)
            # Remove common currency/thousand separators
            if isinstance(value, str):
                value = value.strip().replace(",", "").replace("$", "")
            return float(value)
        except (ValueError, TypeError):
            return None

    def filter_data(
        self,
        data_rows: List[List[Any]],
        filters: Union[str, dict, FilterGroup, SingleFilter],
        headers: Optional[List[str]] = None,
    ) -> Tuple[List[List[Any]], List[int]]:
        """
        Filter data rows and return filtered data with row indices.
        
        This is a convenience method that returns a tuple instead of a dict,
        suitable for use cases that need both the filtered data and original row indices.
        
        Args:
            data_rows: 2D array of data values (without header row)
            filters: Filter specification (string for simple contains search, dict for complex)
            headers: Optional list of column headers for name-based column references
            
        Returns:
            Tuple of (filtered_data, row_indices) where:
            - filtered_data: List of rows that matched the filter
            - row_indices: Original indices of matched rows in data_rows
        """
        # Handle simple string filter (contains search across all columns)
        if isinstance(filters, str):
            filters = SingleFilter(
                column=None,  # Search all columns
                operator=FilterOperator.CONTAINS,
                value=filters
            )
        elif isinstance(filters, dict):
            filters = self._parse_filter_dict(filters)
        
        # Apply filter to each row, tracking indices
        filtered_rows = []
        matched_indices = []
        
        for i, row in enumerate(data_rows):
            col_indices = self._matches_filter(row, filters, headers)
            if col_indices:
                filtered_rows.append(row)
                matched_indices.append(i)
        
        return (filtered_rows, matched_indices)

    def get_warnings(self) -> List[str]:
        """Return any warnings generated during the last filter operation."""
        return self._warnings


def parse_filters(filter_dict: dict) -> Union[FilterGroup, SingleFilter]:
    """
    Parse filter dictionary into FilterGroup or SingleFilter object.

    This is a convenience function for creating filter objects from dict input.

    Args:
        filter_dict: Dictionary representation of filter

    Returns:
        FilterGroup or SingleFilter object

    Example:
        >>> parse_filters({"operator": "and", "conditions": [...]})
        FilterGroup(...)
        >>> parse_filters({"column": {"type": "letter", "value": "A"}, "operator": ">", "value": 100})
        SingleFilter(...)
    """
    engine = FilterEngine()
    return engine._parse_filter_dict(filter_dict)


def apply_filter_to_data(
    data: List[List[Any]],
    filters: dict,
    headers: Optional[List[str]] = None,
) -> dict:
    """
    Apply filters to 2D data array.

    Convenience function that creates a FilterEngine and applies filters.

    Args:
        data: 2D array of values
        filters: Filter dictionary
        headers: Optional list of column headers

    Returns:
        Dictionary with filtered data and metadata
    """
    engine = FilterEngine()
    return engine.apply_filter(data, filters, headers)
