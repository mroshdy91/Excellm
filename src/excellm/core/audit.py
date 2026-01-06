"""Audit logging for ExceLLM MCP server.

Provides simple logging of write operations for debugging and accountability.
"""

import logging

# Create dedicated audit logger
_audit_logger = logging.getLogger("excellm.audit")

# Ensure audit logger outputs to stderr
if not _audit_logger.handlers:
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(
        "%(asctime)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))
    _audit_logger.addHandler(handler)
    _audit_logger.setLevel(logging.INFO)


def log_write(
    tool: str,
    workbook: str,
    sheet: str,
    range_str: str,
    cells: int,
    dry_run: bool = False,
) -> None:
    """Log a write operation for auditing.

    Args:
        tool: Tool name (e.g., "write_range", "write_cell")
        workbook: Workbook name
        sheet: Sheet name
        range_str: Target range
        cells: Number of cells written
        dry_run: Whether this was a dry run
    """
    status = "DRY_RUN" if dry_run else "WRITE"
    _audit_logger.info(
        f"{status} | {tool} | {workbook} | {sheet} | {range_str} | {cells} cells"
    )


def log_vba(
    workbook: str,
    procedure: str,
    success: bool,
) -> None:
    """Log VBA execution for auditing.

    Args:
        workbook: Workbook name
        procedure: VBA procedure name
        success: Whether execution succeeded
    """
    status = "VBA_OK" if success else "VBA_FAIL"
    _audit_logger.info(f"{status} | {workbook} | {procedure}")
