"""Session management for ExceLLM MCP server.

Provides stateful session tracking for large data transformations.
Sessions allow LLMs to process data in chunks without losing progress.

Best Practice: For datasets > 25 rows, use a subagent with focused context
containing only the current chunk's source data.
"""

import logging
import uuid
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from ..core.connection import (
    _init_com,
    get_excel_app,
    get_workbook,
    get_worksheet,
)
from ..core.errors import ErrorCodes, ToolError
from ..core.utils import (
    column_to_number,
)
from .writers import _verify_against_source, write_range_sync

logger = logging.getLogger(__name__)

# Session storage (in-memory)
_sessions: Dict[str, Dict[str, Any]] = {}

# Session expiry time
SESSION_EXPIRY_HOURS = 1


def _cleanup_expired_sessions():
    """Remove sessions that have expired."""
    now = datetime.now()
    expired = [
        sid for sid, session in _sessions.items()
        if now - session["last_activity"] > timedelta(hours=SESSION_EXPIRY_HOURS)
    ]
    for sid in expired:
        del _sessions[sid]
        logger.info(f"Session {sid} expired and cleaned up")


def create_transform_session_sync(
    workbook_name: str,
    sheet_name: str,
    source_column: str,
    output_columns: str,
    start_row: int = 2,
    end_row: Optional[int] = None,
    chunk_size: int = 25,
    verify_key_index: int = 0,
) -> Dict[str, Any]:
    """Create a new transformation session.

    The session tracks progress and allows processing data in chunks.
    Server maintains state so LLM doesn't need to track row numbers.

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        source_column: Column containing source data (e.g., "A")
        output_columns: Target columns for output (e.g., "B:E")
        start_row: First data row (default: 2, assumes row 1 is header)
        end_row: Last data row (None = auto-detect from UsedRange)
        chunk_size: Rows per chunk (default: 25)
        verify_key_index: Index of output column to verify against source

    Returns:
        Dictionary with session_id, total_rows, chunk_size, first_chunk_data
    """
    _init_com()
    _cleanup_expired_sessions()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    # Auto-detect end row if not provided
    if end_row is None:
        used_range = worksheet.UsedRange
        end_row = used_range.Row + used_range.Rows.Count - 1

    total_rows = end_row - start_row + 1

    # Parse output columns
    if ":" in output_columns:
        start_out_col, end_out_col = output_columns.split(":")
    else:
        start_out_col = output_columns
        end_out_col = output_columns

    num_output_cols = column_to_number(end_out_col) - column_to_number(start_out_col) + 1

    # Generate session ID
    session_id = str(uuid.uuid4())[:8]

    # Read first chunk of source data
    first_chunk_end = min(start_row + chunk_size - 1, end_row)
    source_range = f"{source_column}{start_row}:{source_column}{first_chunk_end}"
    source_data = worksheet.Range(source_range).Value

    # Normalize source data
    if source_data is None:
        first_chunk_source = []
    elif isinstance(source_data, (list, tuple)):
        first_chunk_source = [
            row[0] if isinstance(row, (list, tuple)) else row
            for row in source_data
        ]
    else:
        first_chunk_source = [source_data]

    # Create session
    _sessions[session_id] = {
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "source_column": source_column,
        "output_columns": output_columns,
        "start_out_col": start_out_col,
        "num_output_cols": num_output_cols,
        "start_row": start_row,
        "end_row": end_row,
        "chunk_size": chunk_size,
        "verify_key_index": verify_key_index,
        "current_row": start_row,
        "processed_rows": 0,
        "verified_rows": 0,
        "errors": [],
        "last_activity": datetime.now(),
        "status": "active",
    }

    logger.info(f"Created session {session_id} for {workbook_name}:{sheet_name} ({total_rows} rows)")

    return {
        "success": True,
        "session_id": session_id,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "source_column": source_column,
        "output_columns": output_columns,
        "total_rows": total_rows,
        "chunk_size": chunk_size,
        "status": "active",
        "first_chunk": {
            "start_row": start_row,
            "end_row": first_chunk_end,
            "source_data": first_chunk_source,
            "rows": len(first_chunk_source),
        },
        "instruction": (
            f"Transform the {len(first_chunk_source)} source values above into a 2D array "
            f"with {num_output_cols} columns, then call process_chunk with session_id='{session_id}'"
        ),
    }


def create_parallel_sessions_sync(
    workbook_name: str,
    sheet_name: str,
    source_column: str,
    output_columns: str,
    start_row: int = 2,
    end_row: Optional[int] = None,
    num_sessions: int = 2,
    chunk_size: int = 25,
    verify_key_index: int = 0,
) -> Dict[str, Any]:
    """Create multiple parallel sessions for faster processing of large datasets.

    Use when:
    - Dataset is large (>500 rows)
    - You want to spawn multiple subagents to process in parallel
    - Each subagent handles a different row range independently

    Args:
        workbook_name: Name of open workbook
        sheet_name: Name of worksheet
        source_column: Column containing source data (e.g., "A")
        output_columns: Target columns for output (e.g., "B:E")
        start_row: First data row (default: 2)
        end_row: Last data row (None = auto-detect)
        num_sessions: Number of parallel sessions to create (default: 2)
        chunk_size: Rows per chunk within each session (default: 25)
        verify_key_index: Index of output column to verify against source

    Returns:
        List of sessions with session_id, row range, and first_chunk data
    """
    _init_com()
    _cleanup_expired_sessions()

    app = get_excel_app()
    workbook = get_workbook(app, workbook_name)
    worksheet = get_worksheet(workbook, sheet_name)

    # Auto-detect end row if not provided
    if end_row is None:
        used_range = worksheet.UsedRange
        end_row = used_range.Row + used_range.Rows.Count - 1

    total_rows = end_row - start_row + 1
    total_chunks = (total_rows + chunk_size - 1) // chunk_size  # Ceiling division

    # Ensure we don't have more sessions than chunks
    num_sessions = min(num_sessions, total_chunks)

    # Parse output columns for num_output_cols
    if ":" in output_columns:
        start_out_col, end_out_col = output_columns.split(":")
    else:
        start_out_col = output_columns
        end_out_col = output_columns
    num_output_cols = column_to_number(end_out_col) - column_to_number(start_out_col) + 1

    sessions = []

    # Create sessions with interleaved chunk assignment
    # Session 0 gets chunks 0, num_sessions, 2*num_sessions, ...
    # Session 1 gets chunks 1, num_sessions+1, 2*num_sessions+1, ...
    for session_idx in range(num_sessions):
        # Calculate which chunks this session handles
        session_chunks = []
        chunk_idx = session_idx
        while chunk_idx < total_chunks:
            chunk_start = start_row + chunk_idx * chunk_size
            chunk_end = min(chunk_start + chunk_size - 1, end_row)
            session_chunks.append((chunk_start, chunk_end))
            chunk_idx += num_sessions

        if not session_chunks:
            continue

        # First chunk for this session
        first_chunk_start, first_chunk_end = session_chunks[0]

        # Read first chunk source data
        source_range = f"{source_column}{first_chunk_start}:{source_column}{first_chunk_end}"
        source_data = worksheet.Range(source_range).Value

        # Normalize source data
        if source_data is None:
            first_chunk_source = []
        elif isinstance(source_data, (list, tuple)):
            first_chunk_source = [
                row[0] if isinstance(row, (list, tuple)) else row
                for row in source_data
            ]
        else:
            first_chunk_source = [source_data]

        # Generate session ID
        session_id = str(uuid.uuid4())[:8]

        # Calculate total rows for this session
        session_total_rows = sum(ce - cs + 1 for cs, ce in session_chunks)

        # Store session with chunk list
        _sessions[session_id] = {
            "workbook_name": workbook_name,
            "sheet_name": sheet_name,
            "source_column": source_column,
            "output_columns": output_columns,
            "start_out_col": start_out_col,
            "num_output_cols": num_output_cols,
            "chunks": session_chunks,  # List of (start, end) tuples
            "current_chunk_idx": 0,
            "chunk_size": chunk_size,
            "verify_key_index": verify_key_index,
            "processed_rows": 0,
            "verified_rows": 0,
            "errors": [],
            "last_activity": datetime.now(),
            "status": "active",
        }

        logger.info(f"Created parallel session {session_id} with {len(session_chunks)} chunks")

        sessions.append({
            "session_id": session_id,
            "chunks": session_chunks,
            "total_rows": session_total_rows,
            "first_chunk": {
                "start_row": first_chunk_start,
                "end_row": first_chunk_end,
                "source_data": first_chunk_source,
                "rows": len(first_chunk_source),
            },
        })

    return {
        "success": True,
        "workbook": workbook_name,
        "sheet": sheet_name,
        "total_rows": total_rows,
        "num_sessions": len(sessions),
        "sessions": sessions,
        "instruction": (
            f"Spawn {len(sessions)} subagents, each calling process_chunk with their session_id. "
            f"Sessions handle interleaved chunks so parallel completion fills consecutive rows."
        ),
    }


def process_chunk_sync(
    session_id: str,
    data: List[List[Any]],
) -> Dict[str, Any]:
    """Process a chunk of transformed data.

    The server knows which rows to write based on session state.
    Verifies data against source column before writing.

    Args:
        session_id: Session ID from create_transform_session
        data: LLM's transformed data for current chunk (2D array)

    Returns:
        Dictionary with rows_written, verification, remaining, next_chunk_data
    """
    _init_com()

    if session_id not in _sessions:
        raise ToolError(
            f"Session '{session_id}' not found. Create a new session with create_transform_session.",
            ErrorCodes.VALIDATION_ERROR
        )

    session = _sessions[session_id]
    session["last_activity"] = datetime.now()

    # Determine write range based on session type (interleaved vs contiguous)
    if "chunks" in session:
        # Interleaved mode
        if session["current_chunk_idx"] >= len(session["chunks"]):
            raise ToolError("Session already complete", ErrorCodes.VALIDATION_ERROR)

        current_chunk_start, current_chunk_end = session["chunks"][session["current_chunk_idx"]]
        chunk_size = current_chunk_end - current_chunk_start + 1
        write_range = f"{session['start_out_col']}{current_chunk_start}:{session['output_columns'].split(':')[-1]}{current_chunk_end}"
        current_row_for_verification = current_chunk_start
        # Expected rows for this specific chunk
        rows_to_write = chunk_size
    else:
        # Legacy contiguous mode
        current_row = session["current_row"]
        chunk_size = min(len(data), session["chunk_size"])
        end_row = min(current_row + chunk_size - 1, session["end_row"])
        write_range = f"{session['start_out_col']}{current_row}:{session['output_columns'].split(':')[-1]}{end_row}"
        current_row_for_verification = current_row
        rows_to_write = end_row - current_row + 1

    # Trim data to match expected rows
    write_data = data[:rows_to_write]

    # Get worksheet
    app = get_excel_app()
    workbook = get_workbook(app, session["workbook_name"])
    worksheet = get_worksheet(workbook, session["sheet_name"])

    # Verify against source
    verification = _verify_against_source(
        worksheet=worksheet,
        source_column=session["source_column"],
        data=write_data,
        key_index=session["verify_key_index"],
        start_row=current_row_for_verification,
        match_mode="all_columns"  # Updated to Holographic Verification
    )

    # Check for mismatches (zero tolerance)
    if verification["mismatch_count"] > 0:
        session["errors"].extend(verification["mismatches"])
        raise ToolError(
            f"VERIFICATION FAILED: {verification['mismatch_count']} rows failed. "
            f"First mismatch: {verification['mismatches'][0]}. "
            f"Session paused. Fix data and retry process_chunk.",
            ErrorCodes.VALIDATION_ERROR
        )

    # Write the data
    write_range_sync(
        workbook_name=session["workbook_name"],
        sheet_name=session["sheet_name"],
        range_str=write_range,
        data=write_data,
        force_overwrite=True,
        activate=True,
        max_cells=rows_to_write * session["num_output_cols"],
    )

    # Update session state & determine next step
    session["processed_rows"] += rows_to_write
    session["verified_rows"] += verification["matches"]

    if "chunks" in session:
        # Interleaved mode: Move to next chunk in list
        session["current_chunk_idx"] += 1
        is_complete = session["current_chunk_idx"] >= len(session["chunks"])
        if not is_complete:
            next_start, next_end = session["chunks"][session["current_chunk_idx"]]
    else:
        # Legacy contiguous mode: Increment row counter
        session["current_row"] += rows_to_write
        next_start = session["current_row"]
        next_end = min(next_start + session["chunk_size"] - 1, session["end_row"])
        is_complete = next_start > session["end_row"]

    if is_complete:
        session["status"] = "complete"
        return {
            "success": True,
            "session_id": session_id,
            "rows_written": rows_to_write,
            "verification": verification,
            "processed_total": session["processed_rows"],
            "remaining": 0,
            "status": "complete",
            "message": f"Session complete! Processed {session['processed_rows']} rows with 100% verification.",
        }

    # Read next chunk source data
    next_source_range = f"{session['source_column']}{next_start}:{session['source_column']}{next_end}"
    next_source_data = worksheet.Range(next_source_range).Value

    # Normalize
    if next_source_data is None:
        next_chunk_source = []
    elif isinstance(next_source_data, (list, tuple)):
        next_chunk_source = [
            row[0] if isinstance(row, (list, tuple)) else row
            for row in next_source_data
        ]
    else:
        next_chunk_source = [next_source_data]

    # Calculate remaining rows
    if "chunks" in session:
        remaining = sum(ce - cs + 1 for cs, ce in session["chunks"][session["current_chunk_idx"]:])
    else:
        remaining = session["end_row"] - session["current_row"] + 1

    return {
        "success": True,
        "session_id": session_id,
        "rows_written": rows_to_write,
        "verification": verification,
        "processed_total": session["processed_rows"],
        "remaining": remaining,
        "status": "active",
        "next_chunk": {
            "start_row": next_start,
            "end_row": next_end,
            "source_data": next_chunk_source,
            "rows": len(next_chunk_source),
        },
        "instruction": (
            f"Transform the {len(next_chunk_source)} source values above, "
            f"then call process_chunk with session_id='{session_id}'"
        ),
    }


def get_session_status_sync(session_id: str) -> Dict[str, Any]:
    """Get the status of an existing session.

    Use this to check progress or resume an interrupted session.

    Args:
        session_id: Session ID to check

    Returns:
        Dictionary with session status and progress
    """
    _init_com()

    if session_id not in _sessions:
        raise ToolError(
            f"Session '{session_id}' not found.",
            ErrorCodes.VALIDATION_ERROR
        )

    session = _sessions[session_id]
    session["last_activity"] = datetime.now()

    # Calculate remaining and current row
    if "chunks" in session:
        # Interleaved mode
        if session["status"] == "complete" or session["current_chunk_idx"] >= len(session["chunks"]):
            remaining = 0
            # For current_row, use the end of the last chunk if complete
            current_row_for_display = session["chunks"][-1][1]
        else:
            remaining = sum(ce - cs + 1 for cs, ce in session["chunks"][session["current_chunk_idx"]:])
            current_row_for_display = session["chunks"][session["current_chunk_idx"]][0]

        total_rows = sum(ce - cs + 1 for cs, ce in session["chunks"])
    else:
        # Legacy contiguous mode
        remaining = max(0, session["end_row"] - session["current_row"] + 1)
        current_row_for_display = session["current_row"]
        total_rows = session["end_row"] - session["start_row"] + 1

    result = {
        "success": True,
        "session_id": session_id,
        "workbook": session["workbook_name"],
        "sheet": session["sheet_name"],
        "source_column": session["source_column"],
        "output_columns": session["output_columns"],
        "total_rows": total_rows,
        "processed_rows": session["processed_rows"],
        "verified_rows": session["verified_rows"],
        "remaining_rows": remaining,
        "current_row": current_row_for_display,
        "chunk_size": session["chunk_size"],
        "errors_count": len(session["errors"]),
        "errors": session["errors"][-5:],  # Last 5 errors
        "status": session["status"],
    }

    # If session is active, include next chunk data
    if session["status"] == "active" and remaining > 0:
        app = get_excel_app()
        workbook = get_workbook(app, session["workbook_name"])
        worksheet = get_worksheet(workbook, session["sheet_name"])

        if "chunks" in session:
             next_chunk_start, next_chunk_end = session["chunks"][session["current_chunk_idx"]]
        else:
             next_chunk_start = session["current_row"]
             next_chunk_end = min(next_chunk_start + session["chunk_size"] - 1, session["end_row"])

        next_source_range = f"{session['source_column']}{next_chunk_start}:{session['source_column']}{next_chunk_end}"
        next_source_data = worksheet.Range(next_source_range).Value

        if next_source_data is None:
            next_chunk_source = []
        elif isinstance(next_source_data, (list, tuple)):
            next_chunk_source = [
                row[0] if isinstance(row, (list, tuple)) else row
                for row in next_source_data
            ]
        else:
            next_chunk_source = [next_source_data]

        result["next_chunk"] = {
            "start_row": next_chunk_start,
            "end_row": next_chunk_end,
            "source_data": next_chunk_source,
            "rows": len(next_chunk_source),
        }
        result["instruction"] = (
            f"Transform the {len(next_chunk_source)} source values above, "
            f"then call process_chunk with session_id='{session_id}'"
        )

    return result
