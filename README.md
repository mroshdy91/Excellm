# ExceLLM: Excel Live MCP Server

A Model Context Protocol (MCP) server for Excel automation with dual-engine support. Enables LLMs (Claude, ChatGPT, Cursor, etc.) to interact with Excel files through natural language - live on Windows or file-based cross-platform.

## Features

- ✅ **Dual-Engine Architecture**: Live Excel (Windows COM) or file-based (cross-platform)
- ✅ **34 MCP Tools**: Comprehensive Excel automation toolkit
- ✅ **Real-Time Excel Operations**: Work with files open in Excel (Windows)
- ✅ **Cross-Platform File Mode**: Work with .xlsx files on Windows, Mac, Linux
- ✅ **Charts & Pivot Tables**: Native Excel chart and pivot table creation
- ✅ **VBA Execution**: Execute custom macros (Windows only, disabled by default)
- ✅ **Screen Capture**: Visual verification of changes (Windows only)
- ✅ **Excel Tables**: Create, list, delete table objects
- ✅ **Cell Merging**: Merge, unmerge, and query merged cells
- ✅ **Formula Validation**: Validate syntax before applying
- ✅ **Session Management**: Process large datasets with chunking
- ✅ **Advanced Search & Filtering**: Find and filter data efficiently
- ✅ **LLM-Optimized**: Workflow guidance and structured responses

## Prerequisites

### For Live Excel Mode (Windows)
- **Windows OS** (required for COM automation)
- **Microsoft Excel** installed and running
- **Python 3.10** or higher
- `pywin32` library for Windows COM automation
- **At least one workbook open** in Excel

### For File Mode (Cross-Platform)
- **Any OS**: Windows, Mac, or Linux
- **Python 3.10** or higher
- `openpyxl` library for file operations
- **No Excel required**

## Installation

### From Source

```bash
cd ExceLLM
pip install -e .
```

### Development Installation

```bash
cd ExceLLM
pip install -e ".[dev]"
```

## Usage

### Running the Server

```bash
python -m excellm
```

The server will start and listen for MCP client connections using stdio transport.

## MCP Client Integration

### Claude Desktop

**Windows Configuration:**

Create or edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "ExceLLM": {
      "command": "python",
      "args": [
        "-m",
        "excellm"
      ],
      "description": "Real-time Excel automation for open files"
    }
  }
}
```

**macOS Configuration:**

Create or edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "ExceLLM": {
      "command": "python3",
      "args": [
        "-m",
        "excellm"
      ],
      "description": "Real-time Excel automation for open files"
    }
  }
}
```

**Important:**
- Replace python with full path if needed (e.g., `C:\\Python311\\python.exe`)
- Use double backslashes `\\` for Windows paths
- Restart Claude Desktop after making changes

### Cursor AI Editor

Open Cursor Settings → MCP and add:

```json
{
  "mcpServers": {
    "ExceLLM": {
      "command": "python",
      "args": ["-m", "excellm"]
    }
  }
}
```

### ChatGPT (with MCP support)

Configure in ChatGPT's MCP settings similarly to above.

## Available Tools

### 🔍 Inspection (Start Here)

#### 1. `inspect_workbook()`
Fast workbook-level radar. Returns sheet names, visibility, and "scent" of data without reading cells.
**Usage:** `await inspect_workbook()`

#### 2. `explore(scope, mode="quick")`
Sheet-level analysis. Detects used range, labeled regions, and layout.
**Usage:** `await explore({"sheet": "Sheet1"}, mode="deep")`

---

### 📖 Reading & Navigation

#### 3. `list_open_workbooks()`
List all open workbooks and their sheets.
**Usage:** `await list_open_workbooks()`

#### 4. `read(workbook_name, sheet_name, reference, batch, max_rows)`
Read cells or ranges. Supports single range OR efficient batch reading.
**Usage:** `await read("data.xlsx", "Sheet1", "A1:D10")`
**Batch:** `await read("data.xlsx", batch=[{"sheet": "S1", "range": "A1"}, {"range": "B2"}])`
**Pagination:** `await read("data.xlsx", "Sheet1", max_rows=100)` — limits rows to prevent token explosion

#### 5. `search(workbook_name, filters, ..., max_rows)`
Filter data server-side before returning to LLM.
**Usage:** `await search("data.xlsx", {"Column": "Status", "Value": "Active"})`
**Pagination:** `await search("data.xlsx", "active", max_rows=50)` — limits matching rows

#### 6. `get_unique_values(workbook_name, sheet_name, range)`
Get unique values and counts from a column.
**Usage:** `await get_unique_values("data.xlsx", "Sheet1", "A:A")`

#### 7. `get_current_selection()`
Get the currently selected cell/range in the active window.
**Usage:** `await get_current_selection()`

#### 8. `select_range(workbook_name, sheet_name, reference)`
Visually select a range in the Excel UI.
**Usage:** `await select_range("data.xlsx", "Sheet1", "A1:B10")`

---

### 🕒 History & Changes

#### 9. `get_recent_changes(limit=10)`
Get the history of recent user actions from both Undo and Redo stacks.
**Usage:** `await get_recent_changes(limit=5)`
**Returns:** `{"undo_history": [...], "redo_history": [...]}`

---

### ✍️ Writing & Editing

#### 10. `write(workbook_name, sheet_name, reference, data, ...)`
Write values. Supports single cells or 2D arrays. Includes safety checks.
**Usage:** `await write("data.xlsx", "Sheet1", "A1", "Hello")`

#### 11. `copy_range(source_workbook, source_sheet, source_range, target_sheet, ...)`
Copy data between locations, preserving formatting.
**Usage:** `await copy_range("data.xlsx", "Sheet1", "A1:A10", target_sheet="Sheet2")`

#### 12. `find_replace(workbook_name, find_value, replace_value, ...)`
Find and replace text.
**Usage:** `await find_replace("data.xlsx", "Old", "New")`

#### 13. `sort_range(workbook_name, sheet_name, range, sort_by)`
Sort data by multiple columns.
**Usage:** `await sort_range("data.xlsx", "Sheet1", "A1:D50", [{"column": "A", "order": "asc"}])`

---

### 🎨 Formatting

#### 14. `format(workbook_name, sheet_name, reference, style, format)`
Apply styles (color, bold, number format) to ranges.
**Usage:** `await format("data.xlsx", "Sheet1", "A1", style="header")`

#### 15. `get_format(workbook_name, sheet_name, reference)`
Read formatting properties of a range.
**Usage:** `await get_format("data.xlsx", "Sheet1", "A1")`

#### 16. `merge_cells(workbook_name, sheet_name, start_cell, end_cell)`
Merge a range of cells.
**Usage:** `await merge_cells("data.xlsx", "Sheet1", "A1", "D1")`

#### 17. `unmerge_cells(workbook_name, sheet_name, start_cell, end_cell)`
Unmerge previously merged cells.
**Usage:** `await unmerge_cells("data.xlsx", "Sheet1", "A1", "D1")`

#### 18. `get_merged_cells(workbook_name, sheet_name)`
List all merged cell ranges in a sheet.
**Usage:** `await get_merged_cells("data.xlsx", "Sheet1")`

> **💡 Conditional Formatting:** The `format` tool now supports `conditional_format` parameter:
> - ColorScale: `{type: "colorScale", min_color: "FF0000", max_color: "00FF00"}`
> - DataBar: `{type: "dataBar", bar_color: "638EC6"}`
> - IconSet: `{type: "iconSet", icon_style: "3trafficlights"}`
> - CellIs: `{type: "cellIs", operator: "greaterThan", value: 100, fill_color: "FFEB9C"}`

---

### 📋 Sheet & Structure Management

#### 19. `manage_sheet(workbook_name, action, sheet_name, ...)`
Add, rename, delete, hide, copy, or move worksheets.
**Usage:** `await manage_sheet("data.xlsx", action="add", sheet_name="NewSheet")`

#### 20. `insert(workbook_name, sheet_name, insert_type, position, count)`
Insert rows or columns.
**Usage:** `await insert("data.xlsx", "Sheet1", "row", "5", count=2)`

#### 21. `delete(workbook_name, sheet_name, delete_type, position, count)`
Delete rows or columns.
**Usage:** `await delete("data.xlsx", "Sheet1", "column", "C")`

---

### 📊 Excel Tables

#### 22. `create_table(workbook_name, sheet_name, range_ref, table_name)`
Convert a range into an official Excel Table (ListObject).
**Usage:** `await create_table("data.xlsx", "Sheet1", "A1:D10", "SalesTable")`

#### 23. `list_tables(workbook_name)`
List all tables in the workbook.
**Usage:** `await list_tables("data.xlsx")`

#### 24. `delete_table(workbook_name, sheet_name, table_name, keep_data)`
Remove table structure, optionally keeping data.
**Usage:** `await delete_table("data.xlsx", "Sheet1", "SalesTable")`

---

### 📈 Charts & Pivot Tables

#### 25. `create_chart(workbook_name, sheet_name, data_range, chart_type, target_cell, ...)`
Create charts (line, bar, pie, scatter, area) from data.
- **Live Mode:** Native Excel chart automation.
- **File Mode:** Basic chart creation via openpyxl.
**Usage:** `await create_chart("data.xlsx", "Sheet1", "A1:D10", "bar", "F1", title="Sales")`

#### 26. `create_pivot_table(workbook_name, sheet_name, data_range, rows, values, ...)`
Create pivot tables with aggregation.
- **Live Mode:** Native interactive Excel pivot table.
- **File Mode:** Static summary table (calculated in Python).
**Usage:** `await create_pivot_table("data.xlsx", "Sheet1", "A1:D100", rows=["Category"], values=["Amount"], agg_func="sum")`

---

### ⚙️ Advanced Features

#### 27. `execute_vba(workbook_name, vba_code)`
Run custom VBA macros (Windows only).

**🔒 Security:** Disabled by default. Set `EXCELLM_ENABLE_VBA=true` to enable.

**Usage:** `await execute_vba("data.xlsx", "Range('A1').Value = 'VBA'")`

#### 28. `capture_sheet(workbook_name, sheet_name, range_ref)`
Take a screenshot of a range (Windows only).
**Usage:** `await capture_sheet("data.xlsx", "Sheet1", "A1:H10")`

#### 29. `validate_cell_reference(cell)`
Utility to check if a reference string is valid.
**Usage:** `await validate_cell_reference("A1")`

#### 30. `validate_formula(formula)`
Validate Excel formula syntax, check for errors, and get correction suggestions.
**Usage:** `await validate_formula("=SUM(A1:A10)")`

---

### 🚀 Big Data Sessions (Stateful)

For handling large datasets (>50 rows) safely.

#### 31. `create_transform_session(...)`
Start a session to process data in chunks.
**Usage:** `await create_transform_session("data.xlsx", "Sheet1", "A", "B")`

#### 32. `process_chunk(session_id, data)`
Submit processed data for the current chunk.
**Usage:** `await process_chunk("session_123", [[1, 2], [3, 4]])`

#### 33. `get_session_status(session_id)`
Check progress of a session.
**Usage:** `await get_session_status("session_123")`

#### 34. `create_parallel_sessions(...)`
Split work for parallel sub-agents.
**Usage:** `await create_parallel_sessions("data.xlsx", "Sheet1", "A", "B")`

---


**Example Usage:**
```
User: Replace all "N/A" with "Not Available" in Sheet1
LLM: [Calls find_replace("data.xlsx", "N/A", "Not Available", sheet_name="Sheet1")]
  ✓ Found 15 matches, replaced 15 values

User: How many cells contain "error"? Don't replace yet.
LLM: [Calls find_replace("data.xlsx", "error", "", preview_only=True)]
  Found 7 cells containing "error" (preview mode, no changes made)
```

---

## Common Workflows

### 📊 Recommended Workflow Patterns

ExceLLM tools include workflow guidance markers (🔍, ✍️, ⚙️, 📸, 📊) to help LLMs use tools in the correct order.

---

### Pattern 1: Data Analysis Workflow

**Recommended Order:**
1. 🔍 **Inspect** → Understand workbook/sheet structure
2. 📖 **Read** → Get data from identified regions
3. 📊 **Process** → Analyze data
4. ✍️ **Write** → Write results back
5. 📸 **Verify** (optional) → Visual validation

**Example:**
```
User: Analyze sales data in data.xlsx and write summary

LLM Workflow:
1. inspect_workbook() 
   → Identifies sheets, finds "Sales" sheet
   
2. explore({"sheet": "Sales"}, mode="quick")
   → Detects data in A1:D100, headers present
   
3. read("data.xlsx", "Sales", "A1:D100")
   → Reads all sales data

4. [Analysis] → Calculates totals, averages
   
5. write("data.xlsx", "Summary", "A1:C5", summary_data)
   → Writes analysis to Summary sheet
   
6. capture_sheet("data.xlsx", "Summary") [OPTIONAL]
   → Screenshots for validation
```

---

### Pattern 2: Data Transformation (Large Datasets)

**For datasets > 25 rows:**
1. 🔍 **Explore** → Understand data structure
2. 🎯 **Create Session** → Start stateful processing
3. 🔄 **Process Chunks** → Iterate through chunks
4. ✅ **Verify** → Check session status

**Example:**
```
User: Extract PO numbers from 500 rows of messy text

LLM Workflow:
1. explore({"sheet": "Data"}, mode="deep")
   → Identifies 500 rows in column A
   
2. create_transform_session(
     workbook_name="data.xlsx",
     sheet_name="Data",
     source_column="A",
     output_columns="B:D",
     start_row=2,
     chunk_size=25
   )
   → Session created, first chunk received
   
3. process_chunk(session_id, transformed_data)
   → Process chunk 1 (rows 2-26)
   → Server returns next chunk automatically
   
4. process_chunk(session_id, transformed_data)
   → Process chunk 2 (rows 27-51)
   → Continue until complete
   
5. get_session_status(session_id)
   → Verify all 500 rows processed
```

---

### Pattern 3: Formatting & Styling

**Recommended Order:**
1. 📖 **Read/Explore** → Identify target range
2. ✍️ **Write** → Write data if needed
3. 🎨 **Format** → Apply formatting
4. 📊 **Create Table** (optional) → Convert to Excel table
5. 📸 **Capture** (optional) → Visual verification

**Example:**
```
User: Create a formatted sales table with styling

LLM Workflow:
1. write("data.xlsx", "Sheet1", "A1:D100", sales_data)
   → Write raw data
   
2. format("data.xlsx", "Sheet1", "A1:D1", style="header")
   → Format header row
   
3. format("data.xlsx", "Sheet1", "B2:D100", 
          format={"numberFormat": "$#,##0.00"})
   → Currency formatting for amounts
   
4. create_table("data.xlsx", "Sheet1", "A1:D100",
                "SalesData", table_style="medium9")
   → Convert to Excel table with filters
   
5. capture_sheet("data.xlsx", "Sheet1", "A1:D100")
   → Screenshot for validation
```

---

### Pattern 4: Advanced Operations (VBA)

**⚠️ USE WITH CAUTION - Only when standard tools are insufficient**

**Recommended Order:**
1. 🔍 **Explore** → Understand current state
2. ⚙️ **Execute VBA** → Run custom macro
3. 📸 **Capture** → Verify results visually

**Example:**
```
User: Apply complex conditional formatting that requires VBA

LLM Workflow:
1. explore({"sheet": "Data"})
   → Understand data layout
   
2. execute_vba("data.xlsx", '''
     Dim ws As Worksheet
     Set ws = ActiveSheet
     
     ' Complex conditional formatting
     ws.Range("A1:D100").FormatConditions.Add ...
     ws.Range("A1:D100").FormatConditions(1).Interior.Color = RGB(255,200,200)
   ''')
   → Execute VBA for complex operation
   
3. capture_sheet("data.xlsx", "Data", "A1:D100")
   → Verify formatting applied correctly
```

---

### Workflow 1: Quick Data Update

```
1. List workbooks → identify target file
2. Read current cell value
3. Calculate/update value
4. Write new value back
5. Save workbook
```

**Example:**
```
User: Find all cells in column A that are empty in data.xlsx and fill them with "N/A"
LLM: [Chains multiple tool calls]
  1. list_open_workbooks() - find data.xlsx
  2. read_range("data.xlsx", "Sheet1", "A1:A100") - get values
  3. [Processing] - identify empty cells
  4. write_range() - write "N/A" to empty cells
  5. save_workbook("data.xlsx") - save changes
  ✓ Filled 15 empty cells with "N/A"
```

---

### Workflow 2: Data Analysis

```
1. Read range of data
2. Analyze data patterns
3. Calculate statistics
4. Write summary to Excel
```

**Example:**
```
User: Analyze the sales data in data.xlsx and write a summary to the Summary sheet
LLM: [Chains tools]
  1. read_range("data.xlsx", "Sheet1", "A2:D100") - get sales
  2. [Analysis] - calculates total, average, max/min
  3. write_range("data.xlsx", "Summary", "A1:C5", [summary data])
  4. save_workbook("data.xlsx")
  ✓ Analysis complete. Summary written to Summary sheet.
```

---

### Workflow 3: Cell-by-Cell Validation

```
1. Read a cell
2. Validate format
3. Update status in another cell
```

**Example:**
```
User: Validate all email addresses in column B and mark invalid ones in column C
LLM: [Iterative workflow]
  ✓ Validated 50 email addresses
  ✓ Marked 3 invalid emails in column C
```

---

## 🌍 Cross-Platform Support

ExceLLM now supports two operation modes:

### Live Excel Mode (Windows + COM)
- **Platform:** Windows only
- **Requires:** Excel running with workbook open
- **Features:** Full feature set (VBA, screen capture, tables, etc.)
- **Use when:** Working with open Excel files on Windows

### File Mode (Cross-platform)
- **Platform:** Windows, Mac, Linux
- **Requires:** openpyxl library (no Excel needed)
- **Features:** Core operations + Charts + Pivot Summaries
- **Use when:** Working with closed .xlsx files, or on Mac/Linux

**Auto-Detection:**
- Provide file path (e.g., `C:/data/file.xlsx`) → File mode
- Provide workbook name (e.g., `data.xlsx`) → Live mode (if Excel running)

**Example:**
```python
# File mode (works anywhere)
await read(workbook_path="/Users/me/Documents/report.xlsx", ...)

# Live mode (Windows + Excel)
await read(workbook_name="report.xlsx", ...)
```

---

## Tool Reference Guide

### 🔍 Inspection Tools (STEP 1 - Use First)
- `inspect_workbook()` - Fast workbook overview
- `explore()` - Sheet-level analysis (quick/deep modes)

### 📖 Read Operations & Navigation
- `list_open_workbooks()` - List all open workbooks
- `read()` - Read cells/ranges with filtering and pagination
- `search()` - Find and filter data with pagination
- `get_unique_values()` - Extract unique values
- `get_current_selection()` - Get active cell
- `select_range()` - Visually select a range
- `get_recent_changes()` - Get Undo/Redo history

### ✍️ Write Operations
- `write()` - Write with safety guardrails
- `copy_range()` - Copy with formatting
- `sort_range()` - Multi-column sorting
- `find_replace()` - Find and replace with preview

### 🎨 Formatting & Cells
- `format()` - Apply predefined or custom formats
- `get_format()` - Read formatting details
- `merge_cells()` - Merge a range of cells
- `unmerge_cells()` - Unmerge a range of cells
- `get_merged_cells()` - List all merged cell ranges

### 📋 Sheet Management
- `manage_sheet()` - Add, remove, hide, copy, rename
- `insert()` - Insert rows/columns
- `delete()` - Delete rows/columns

### 📊 Tables & Charts
- `create_table()` - Create Excel table objects
- `list_tables()` - List all tables
- `delete_table()` - Remove tables
- `create_chart()` - Create native charts
- `create_pivot_table()` - Create interactive pivot tables

### ⚙️ Advanced & Utilities
- `execute_vba()` - Execute VBA macros (Windows only)
- `capture_sheet()` - Screenshot capture (Windows only)
- `validate_cell_reference()` - Check reference syntax
- `validate_formula()` - Validate formula syntax

### 🎯 Session Management (For Large Datasets)
- `create_transform_session()` - Start stateful processing
- `process_chunk()` - Process data chunks
- `get_session_status()` - Check progress
- `create_parallel_sessions()` - Multi-threaded processing

---

## 💡 Best Practices

### 1. Always Inspect First
```
❌ DON'T: Immediately read/write without understanding structure
✅ DO: inspect_workbook() or explore() first
```

### 2. Use Safety Features
```
❌ DON'T: write() with force_overwrite=True by default
✅ DO: Use verify_source parameter for data transformations
```

### 3. Chunk Large Datasets
```
❌ DON'T: Process 500+ rows in one write() call
✅ DO: Use create_transform_session() for >25 rows
```

### 4. VBA as Last Resort
```
❌ DON'T: Use execute_vba() for simple operations
✅ DO: Try standard tools first, VBA only when necessary
```

### 5. Verify Visual Changes
```
❌ DON'T: Trust formatting changes blindly
✅ DO: Use capture_sheet() to verify complex formatting
```

### 6. Use Pagination for Large Reads
```
❌ DON'T: Read 10,000 rows without limits
✅ DO: Use max_rows parameter to limit token usage
     await read("data.xlsx", "Sheet1", max_rows=100)
```

## Audit Logging

All write operations are logged for debugging and accountability:

```
2026-01-05 11:40:00 | WRITE | write_range | data.xlsx | Sheet1 | A1:D10 | 40 cells
2026-01-05 11:40:01 | DRY_RUN | write_cell | data.xlsx | Sheet1 | A1 | 1 cells
```

Logs are written to stderr and include:
- Timestamp
- Operation type (WRITE or DRY_RUN)
- Tool name
- Workbook, sheet, and range
- Number of cells affected


## Troubleshooting

### "Could not connect to Excel. Is Excel running?"

**Solution:**
- Open Microsoft Excel
- Open at least one workbook
- Ensure Excel is not in a dialog/macro execution that blocks COM access

### "Worksheet 'SheetName' is protected"

**Solution:**
- Unprotect the worksheet in Excel
- Go to: **Review → Unprotect Sheet**
- Remove password if prompted

### "Workbook 'file.xlsx' is read-only"

**Solution:**
- Close the workbook
- Open it with write permissions
- Right-click file → Properties → Uncheck "Read-only"

### "Invalid cell reference: 'A1B2'"

**Solution:**
- Use proper Excel format: A1, B5, Z100, AA123
- Valid: 1-3 letters + 1-7 digits
- Invalid: 1A, A, A1B2, AAA99999

### "Server not showing up in Claude Desktop"

**Solution:**
1. Verify `claude_desktop_config.json` is valid JSON
2. Use absolute paths in configuration
3. Use double backslashes `\\` for Windows paths
4. Completely quit and restart Claude Desktop (close tray icon too)
5. Check Claude logs: `~/Library/Logs/Claude/mcp.log` (macOS) or `%LOCALAPPDATA%\Claude\mcp.log` (Windows)

### MCP Server Won't Start

**Possible Causes:**
- Python not in PATH
- Dependencies not installed
- Excel not installed

**Solutions:**
```bash
# Check Python version
python --version  # Should be 3.10+

# Install dependencies
pip install -r requirements.txt

# Test Excel connection manually
python -c "import win32com.client; win32com.client.Dispatch('Excel.Application')"
```

## Architecture

```
ExceLLM/
├── src/
│   └── excellm/
│       ├── __init__.py           # Package init
│       ├── __main__.py           # Module entry point
│       ├── server.py              # Main MCP server with 34 tools
│       ├── config.py              # Configuration management
│       ├── excel_session.py       # Excel COM session manager
│       ├── validators.py          # Input validation utilities
│       ├── filters.py             # Filter engine for search
│       ├── core/                  # Shared foundation
│       │   ├── __init__.py
│       │   ├── connection.py     # COM pooling, batch reads
│       │   ├── errors.py         # ToolError, ErrorCodes
│       │   ├── audit.py          # Operation logging and tracking
│       │   ├── utils.py          # Consolidated utilities
│       │   ├── engine.py         # Engine abstraction layer
│       │   ├── com_engine.py     # COM engine implementation
│       │   └── file_engine.py   # File-based engine implementation
│       ├── tools/                 # Tool implementations
│       │   ├── readers.py        # read operations
│       │   ├── writers.py        # write operations
│       │   ├── formatters.py     # format, get_format
│       │   ├── sheet_mgmt.py     # manage_sheet, insert, delete
│       │   ├── range_ops.py      # copy_range, sort_range, find_replace
│       │   ├── search.py         # search with filters
│       │   ├── workbook.py       # list_workbooks, select_range
│       │   ├── chart.py         # chart creation
│       │   ├── pivot.py         # pivot table operations
│       │   ├── history.py        # undo/redo history
│       │   ├── session.py        # transform sessions
│       │   ├── table_ops.py     # Excel table operations
│       │   ├── screen_capture.py # screenshot capture
│       │   └── vba_execution.py # VBA macro execution
│       └── inspection/            # Sheet/workbook inspection
│           ├── __init__.py
│           ├── explore.py        # Sheet-level radar
│           ├── inspect_workbook.py
│           ├── types.py          # Pydantic schemas
│           └── utils.py
├── tests/                         # Unit and integration tests
├── requirements.txt               # Dependencies
├── pyproject.toml                 # Package configuration
└── README.md                      # This file
```

### Key Components

1. **FastMCP Server** (`server.py`):
   - Provides 34 MCP tools
   - Handles tool registration and routing
   - Manages server lifecycle
   - Dual-engine architecture support (COM + File)
   - Centralized error handling

2. **Core Module** (`core/`):
   - **connection.py**: Thread-local COM connection pooling, batch range reads for performance
   - **errors.py**: Centralized error handling (ToolError, ErrorCodes)
   - **audit.py**: Operation logging and tracking for debugging
   - **utils.py**: Consolidated utility functions
   - **engine.py**: Engine abstraction layer for cross-platform support
   - **com_engine.py**: Windows COM engine implementation for live Excel operations
   - **file_engine.py**: Cross-platform file-based engine using openpyxl

3. **Tools Module** (`tools/`):
   - **readers.py**: Read operations with filtering and batch support
   - **writers.py**: Write operations with safety checks
   - **formatters.py**: Format, get_format, merge/unmerge cells
   - **sheet_mgmt.py**: manage_sheet, insert, delete operations
   - **range_ops.py**: copy_range, sort_range, find_replace
   - **search.py**: Server-side search with advanced filters
   - **workbook.py**: list_workbooks, select_range operations
   - **chart.py**: Chart creation (line, bar, pie, scatter, area)
   - **pivot.py**: Pivot table operations with aggregation
   - **history.py**: Undo/redo history tracking
   - **session.py**: Transform sessions for large datasets
   - **table_ops.py**: Excel table operations (create, list, delete)
   - **screen_capture.py**: Screenshot capture (Windows only)
   - **vba_execution.py**: VBA macro execution (Windows only)

4. **Inspection Module** (`inspection/`):
   - **explore.py**: Sheet-level analysis with quick/deep modes
   - **inspect_workbook.py**: Fast workbook-level radar
   - **types.py**: Pydantic schemas for type safety
   - **utils.py**: Inspection utilities

5. **ExcelSessionManager** (`excel_session.py`):
   - Connects to running Excel instance (Windows)
   - Wraps COM operations with async support
   - Provides thread-safe access to Excel
   - Handles Excel COM lifecycle

6. **Configuration** (`config.py`):
   - Centralized configuration management
   - Environment variable support
   - Default settings and overrides

7. **Validators** (`validators.py`):
   - Cell reference format validation
   - Workbook/sheet name validation
   - Range parsing and validation
   - Value type checking

8. **Filters** (`filters.py`):
   - Server-side filtering engine
   - Advanced search capabilities
   - Multiple filter types support

## Development

### Running Tests

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Run with coverage
pytest --cov=src/excellm
```

### Code Quality

```bash
# Format code
black src/

# Lint code
ruff check src/

# Type check
mypy src/
```

## Limitations

### Live Excel Mode (COM Engine)
- **Windows Only**: Requires Windows OS for COM automation
- **Excel Must Be Running**: Cannot open files directly, works with open files only
- **Single Excel Instance**: Connects to first running Excel instance
- **VBA Access**: Requires "Trust access to VBA project object model" enabled for VBA execution

### File Mode (openpyxl Engine)
- **Limited Formatting**: Basic font, fill, and borders only
- **No VBA**: Cannot execute macros
- **No Screen Capture**: Cannot generate screenshots
- **No Conditional Formatting**: Not supported (COM only)
- **Charts**: Supported (5 types)
- **Pivot Tables**: Static summary tables only (no drill-down)

### General
- **Large Datasets**: For best performance, use session management for >100 rows
- **File Corruption**: Always backup important files before automation

## Security Considerations

- Tool calls require user approval in most MCP clients
- No remote API calls - local COM operations only
- Read/write operations limited to opened workbooks
- Cell validation prevents malicious input

### VBA Execution Security

**VBA is disabled by default** for security. VBA code runs with full Excel/COM privileges and can:
- Modify files on disk
- Access system resources
- Execute arbitrary code

To enable VBA execution:

```bash
# Environment variable
set EXCELLM_ENABLE_VBA=true

# Or in .env file
EXCELLM_ENABLE_VBA=true
```

> ⚠️ **Warning:** Only enable VBA in trusted environments where you control the input.

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - See LICENSE file for details

## Support

For issues, questions, or contributions:

- **GitHub Issues**: [Create an issue](https://github.com/mroshdy91/Excellm/issues)
- **Documentation**: [Read the docs](https://github.com/mroshdy91/Excellm/blob/main/README.md)

## Acknowledgments

- **MCP Team**: For the Model Context Protocol
- **pywin32**: For Windows COM automation
- **openpyxl**: For cross-platform Excel file operations  
- **FastMCP**: For the excellent MCP server framework

## Version History

### 1.1.0 (2026-01-05)
- **New Features:**
  - Configuration management system (config.py)
  - Audit logging for operation tracking (core/audit.py)
  - Dual-engine architecture with COM and File engines
  - Engine abstraction layer for cross-platform support
- **Improvements:**
  - Enhanced error handling and validation
  - Performance optimizations with batch reads and connection pooling
  - Updated documentation with complete architecture
  - All 34 MCP tools improved with better error messages
- **Documentation:**
  - Complete Architecture section with all modules documented
  - Updated Key Components with 8 comprehensive sections

### 1.0.0-alpha (2026-01-04)
- Initial alpha release
- 34 MCP tools for Excel automation
- Dual-engine architecture (COM + File-based)
- Cross-platform support (Windows, Mac, Linux)
