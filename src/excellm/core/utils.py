"""Consolidated utility functions for ExceLLM MCP server.

This module provides all shared utility functions for Excel operations,
consolidating previously duplicated functions from multiple modules.
"""

import re
from typing import Any, List, Optional, Tuple


def normalize_address(addr: str) -> str:
    """Normalize Excel address by removing $ signs.

    Args:
        addr: Excel address like "$A$1:$B$2" or "A1:B2"

    Returns:
        Normalized address like "A1:B2"
    """
    if not addr:
        return addr
    return addr.replace("$", "")


def column_to_number(col: str) -> int:
    """Convert Excel column letter to number (1-based).

    Args:
        col: Column letter like "A", "Z", "AA"

    Returns:
        Column number (1-based)

    Examples:
        >>> column_to_number("A")
        1
        >>> column_to_number("Z")
        26
        >>> column_to_number("AA")
        27
    """
    result = 0
    for char in col.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def number_to_column(n: int) -> str:
    """Convert column number to Excel letter (1-based).

    Args:
        n: Column number (1-based)

    Returns:
        Column letter like "A", "Z", "AA"

    Examples:
        >>> number_to_column(1)
        'A'
        >>> number_to_column(26)
        'Z'
        >>> number_to_column(27)
        'AA'
    """
    result = ""
    while n > 0:
        n -= 1
        result = chr(ord('A') + (n % 26)) + result
        n //= 26
    return result


def parse_range_bounds(range_str: str) -> Tuple[int, int, int, int]:
    """Parse A1:B2 format range into row/col bounds.

    Args:
        range_str: Range like "A1:Z99" or "A1"

    Returns:
        Tuple of (start_row, start_col, end_row, end_col) - all 1-based
    """
    # Normalize first
    range_str = normalize_address(range_str)

    # Match A1:B2 or just A1
    pattern = r"([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?"
    match = re.match(pattern, range_str)

    if not match:
        return (1, 1, 1, 1)

    start_col = column_to_number(match.group(1))
    start_row = int(match.group(2))

    if match.group(3) and match.group(4):
        end_col = column_to_number(match.group(3))
        end_row = int(match.group(4))
    else:
        end_col = start_col
        end_row = start_row

    return (start_row, start_col, end_row, end_col)


def build_range_address(
    start_row: int, start_col: int, end_row: int, end_col: int
) -> str:
    """Build A1:B2 format range from row/col bounds.

    Args:
        start_row, start_col, end_row, end_col: 1-based row/col numbers

    Returns:
        Range string like "A1:B2"
    """
    start = f"{number_to_column(start_col)}{start_row}"
    end = f"{number_to_column(end_col)}{end_row}"

    if start == end:
        return start
    return f"{start}:{end}"


def is_cell_empty(value: Any) -> bool:
    """Check if a cell value is empty.

    Args:
        value: Cell value from Excel

    Returns:
        True if cell is empty (None or empty string)
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def safe_int(value: Any, default: int = 0) -> int:
    """Safely convert to int with default.

    Args:
        value: Value to convert
        default: Default value if conversion fails

    Returns:
        Integer value or default
    """
    try:
        return int(value) if value is not None else default
    except (ValueError, TypeError):
        return default


def safe_bool(value: Any, default: bool = False) -> bool:
    """Safely convert to bool with default.

    Args:
        value: Value to convert
        default: Default value if conversion fails

    Returns:
        Boolean value or default
    """
    try:
        return bool(value) if value is not None else default
    except (ValueError, TypeError):
        return default


def compute_density(non_empty: int, total: int) -> Optional[float]:
    """Compute density ratio.

    Args:
        non_empty: Number of non-empty cells
        total: Total cells

    Returns:
        Density 0..1 or None if total is 0
    """
    if total <= 0:
        return None
    return round(non_empty / total, 3)


def get_cell_type(value: Any) -> str:
    """Determine the type of a cell value.

    Args:
        value: Cell value

    Returns:
        Type string: "string", "number", "boolean", "date", "empty", "error"
    """
    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        # Check for Excel error values (negative integers in specific range)
        if isinstance(value, int) and value < -2000000000:
            return "error"
        return "number"
    if isinstance(value, str):
        if value.strip() == "":
            return "empty"
        return "string"
    # pywintypes.datetime
    type_name = type(value).__name__
    if "datetime" in type_name.lower() or "time" in type_name.lower():
        return "date"
    return "string"


# Sheet name heuristics for scoring
PRIORITY_SHEET_PATTERNS = [
    r"^data$",
    r"^input$",
    r"^main$",
    r"^sheet1$",
    r"^summary$",
    r"^dashboard$",
    r"^report$",
    r"details",
    r"list",
]


def sheet_name_priority_boost(name: str) -> float:
    """Compute priority boost based on sheet name.

    Args:
        name: Sheet name

    Returns:
        Boost value 0..0.2
    """
    name_lower = name.lower().strip()

    for i, pattern in enumerate(PRIORITY_SHEET_PATTERNS):
        if re.search(pattern, name_lower):
            # Earlier patterns get higher boost
            return 0.2 - (i * 0.02)

    return 0.0


def generate_sample_positions(
    start_row: int, start_col: int, end_row: int, end_col: int,
    max_tiles: int = 12, probes_per_tile: int = 3
) -> List[Tuple[int, int]]:
    """Generate sample cell positions for quick scanning.

    Uses a tile grid approach to sample cells across the range.

    Args:
        start_row, start_col, end_row, end_col: Range bounds (1-based)
        max_tiles: Maximum tiles per dimension
        probes_per_tile: Number of probes per tile

    Returns:
        List of (row, col) positions to probe
    """
    positions = []

    total_rows = end_row - start_row + 1
    total_cols = end_col - start_col + 1

    # Compute tile sizes
    tile_rows = max(1, total_rows // max_tiles)
    tile_cols = max(1, total_cols // max_tiles)

    # Number of tiles
    num_row_tiles = min(max_tiles, total_rows)
    num_col_tiles = min(max_tiles, total_cols)

    for i in range(num_row_tiles):
        for j in range(num_col_tiles):
            tile_start_row = start_row + i * tile_rows
            tile_start_col = start_col + j * tile_cols

            # Sample center of tile
            center_row = tile_start_row + tile_rows // 2
            center_col = tile_start_col + tile_cols // 2

            # Clamp to range
            center_row = min(center_row, end_row)
            center_col = min(center_col, end_col)

            positions.append((center_row, center_col))

            # Add corners if space allows and probes_per_tile > 1
            if probes_per_tile > 1 and tile_rows > 1 and tile_cols > 1:
                # Top-left of tile
                positions.append((tile_start_row, tile_start_col))

            if probes_per_tile > 2:
                # Bottom-right of tile
                br_row = min(tile_start_row + tile_rows - 1, end_row)
                br_col = min(tile_start_col + tile_cols - 1, end_col)
                positions.append((br_row, br_col))

    # Cap total probes
    max_probes = 300
    if len(positions) > max_probes:
        # Take evenly distributed sample
        step = len(positions) // max_probes
        positions = positions[::step][:max_probes]

    return positions
