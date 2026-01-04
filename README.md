# ExceLLM: Excel Live MCP Server

A Model Context Protocol (MCP) server for Excel automation with dual-engine support. Enables LLMs (Claude, ChatGPT, Cursor, etc.) to interact with Excel files through natural language - live on Windows or file-based cross-platform.

## Features

- ✅ **Dual-Engine Architecture**: Live Excel (Windows COM) or file-based (cross-platform)
- ✅ **25 MCP Tools**: Comprehensive Excel automation toolkit
- ✅ **Real-Time Excel Operations**: Work with files open in Excel (Windows)
- ✅ **Cross-Platform File Mode**: Work with .xlsx files on Windows, Mac, Linux
- ✅ **VBA Execution**: Execute custom macros (Windows only)
- ✅ **Screen Capture**: Visual verification of changes (Windows only)
- ✅ **Excel Tables**: Create, list, delete table objects
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

### 1. `list_open_workbooks()`

List all currently open Excel workbooks with their sheet names.

**Returns:**
```json
{
  "success": true,
  "workbooks": [
    {
      "name": "data.xlsx",
      "sheets": ["Sheet1", "Sheet2", "Summary"]
    },
    {
      "name": "report.xlsx",
      "sheets": ["Dashboard"]
    }
  ],
  "count": 2
}
```

**Example Usage:**
```
User: List all open Excel workbooks
LLM: [Calls list_open_workbooks()]
  Found 2 open workbooks:
  - data.xlsx: Sheet1, Sheet2, Summary
  - report.xlsx: Dashboard
```

---

### 2. `read_cell(workbook_name, sheet_name, cell)`

Read a single cell value from an open Excel workbook.

**Parameters:**
- `workbook_name` (string, required): Name of open workbook
- `sheet_name` (string, required): Name of worksheet
- `cell` (string, required): Cell reference (e.g., "A1", "B5", "Z100")

**Returns:**
```json
{
  "success": true,
  "workbook": "data.xlsx",
  "sheet": "Sheet1",
  "cell": "A1",
  "value": "Approved",
  "type": "string"
}
```

**Example Usage:**
```
User: What's the value in cell A1 of Sheet1 in data.xlsx?
LLM: [Calls read_cell("data.xlsx", "Sheet1", "A1")]
  Cell A1 in data.xlsx!Sheet1 contains: "Approved"
```

---

### 3. `read_range(workbook_name, sheet_name, range_str)`

Read a range of cells from an open Excel workbook.

**Parameters:**
- `workbook_name` (string, required): Name of open workbook
- `sheet_name` (string, required): Name of worksheet
- `range_str` (string, required): Range reference (e.g., "A1:C5", "B2:D10")

**Returns:**
```json
{
  "success": true,
  "workbook": "data.xlsx",
  "sheet": "Sheet1",
  "range": "A1:C3",
  "data": [
    ["Name", "Age", "City"],
    ["Alice", 30, "NYC"],
    ["Bob", 25, "LA"]
  ],
  "rows": 3,
  "cols": 3
}
```

**Example Usage:**
```
User: Read the first 3 rows from Sheet1 in data.xlsx, columns A to C
LLM: [Calls read_range("data.xlsx", "Sheet1", "A1:C3")]
  Here's the data from range A1:C3:
  Row 1: Name, Age, City
  Row 2: Alice, 30, NYC
  Row 3: Bob, 25, LA
```

---

### 4. `write_cell(workbook_name, sheet_name, cell, value, auto_save)`

Write a value to a cell in an open Excel workbook.

**Parameters:**
- `workbook_name` (string, required): Name of open workbook
- `sheet_name` (string, required): Name of worksheet
- `cell` (string, required): Cell reference (e.g., "A1", "B5")
- `value` (string, required): Value to write
- `auto_save` (boolean, optional): Auto-save workbook after writing (default: True)

**Returns:**
```json
{
  "success": true,
  "workbook": "data.xlsx",
  "sheet": "Sheet1",
  "cell": "A1",
  "value": "Approved",
  "saved": true
}
```

**Example Usage:**
```
User: Write "Approved" to cell A1 in Sheet1 of data.xlsx
LLM: [Calls write_cell("data.xlsx", "Sheet1", "A1", "Approved")]
  ✓ Successfully wrote "Approved" to data.xlsx!Sheet1!A1
  Workbook saved automatically.
```

---

### 5. `write_range(workbook_name, sheet_name, range_str, data, auto_save)`

Write a 2D array of values to a range in Excel.

**Parameters:**
- `workbook_name` (string, required): Name of open workbook
- `sheet_name` (string, required): Name of worksheet
- `range_str` (string, required): Range reference (e.g., "A1:C5")
- `data` (array, required): 2D array of values (list of lists)
- `auto_save` (boolean, optional): Auto-save after writing (default: True)

**Returns:**
```json
{
  "success": true,
  "workbook": "data.xlsx",
  "sheet": "Sheet1",
  "range": "A1:C2",
  "cells_written": 4,
  "saved": true
}
```

**Example Usage:**
```
User: Update the data in data.xlsx Sheet1 range A1:C2 with new values
LLM: [Calls write_range with data]
  ✓ Successfully wrote 4 cells to data.xlsx!Sheet1!A1:C2
  Workbook saved automatically.
```

---

### 6. `save_workbook(workbook_name, save_as)`

Save an open Excel workbook.

**Parameters:**
- `workbook_name` (string, required): Name of open workbook
- `save_as` (string, optional): New filepath to save as (optional)

**Returns:**
```json
{
  "success": true,
  "workbook": "data.xlsx",
  "saved_as": "C:\\Users\\User\\Documents\\data.xlsx",
  "message": "Workbook saved successfully"
}
```

**Example Usage:**
```
User: Save the data.xlsx workbook
LLM: [Calls save_workbook("data.xlsx")]
  ✓ Workbook "data.xlsx" saved successfully
```

---

### 7. `validate_cell_reference(cell)`

Validate an Excel cell reference format.

**Parameters:**
- `cell` (string, required): Cell reference to validate

**Returns:**
```json
{
  "valid": true,
  "cell": "A1",
  "message": "Valid cell reference"
}
```

**Example Usage:**
```
User: Is "A1" a valid cell reference?
LLM: [Calls validate_cell_reference("A1")]
  ✓ Yes, "A1" is a valid cell reference
```

---


---

### 8. `inspect_workbook()`


Fast workbook-level radar across ALL sheets without reading cell values. Provides enough information to select which sheet to explore deeply.

**Returns:**
```json
{
  "meta": {
    "tool": "inspect_workbook",
    "version": "v1",
    "timestamp": "2026-01-03T05:37:52.613384+00:00",
    "durationMs": 1399
  },
  "workbook": {
    "name": "data.xlsx",
    "path": "C:\\Users\\User\\Documents",
    "readOnly": false,
    "protected": false
  },
  "activeSheet": "Sheet1",
  "sheetsIndex": [
    {
      "name": "Sheet1",
      "state": "visible",
      "protected": false,
      "usedRangeReported": "A1:M31",
      "dataCellCount": 271,
      "layoutFlags": {
        "hasTableObjects": true,
        "hasAutoFilter": false,
        "hasFreezePanes": true,
        "commentsCount": 0,
        "formulasCount": 54
      },
      "flags": ["HAS_TABLE_OBJECTS", "HAS_FREEZE_PANES", "HAS_FORMULAS"],
      "score": { "priority": 0.9, "density": 0.672 }
    }
  ],
  "recommendations": {
    "primaryCandidateSheets": ["Sheet1", "Summary"],
    "avoidSheets": [],
    "nextExploreSheet": "Sheet1"
  }
}
```

**Flags Detected:**
- `EMPTY_OR_NEAR_EMPTY`, `USED_RANGE_INFLATED`, `EXTREME_USED_RANGE`
- `LIKELY_FORMAT_ONLY`, `HIDDEN_SHEET`
- `HAS_TABLE_OBJECTS`, `HAS_FILTER`, `HAS_FREEZE_PANES`
- `MERGED_CELLS_PRESENT`, `HAS_COMMENTS_NOTES`, `HAS_FORMULAS`

**Example Usage:**
```
User: Scan the workbook and tell me which sheet to explore
LLM: [Calls inspect_workbook()]
  Scanned 5 sheets:
  - "PO Reconciliation" (priority 0.9): 271 data cells, has tables, formulas
  - "Summary" (priority 0.85): 39 cells
  Recommended: Start with "PO Reconciliation"
```

---

### 9. `explore(scope, mode)`

Sheet radar that mimics human first glance: structure/layout without reading full data.

**Parameters:**
- `scope` (object, required): Target sheet - `{"sheet": "ACTIVE"}` or `{"sheet": "SheetName"}`
- `mode` (string, optional): `"quick"` (default) or `"deep"`

**Quick Mode:** Fast sampling-based analysis (~300 cell probes max)
**Deep Mode:** Thorough analysis with region detection and accurate bounds

**Returns:**
```json
{
  "meta": {
    "tool": "explore",
    "version": "v1",
    "durationMs": 505,
    "sheet": "Sheet1",
    "mode": "quick"
  },
  "dataFootprint": {
    "usedRangeReported": "A1:M31",
    "realDataBounds": "A2:I29",
    "dataCellCount": 271,
    "nonEmptyRows": 28,
    "nonEmptyCols": 9
  },
  "regions": [
    {
      "id": "R1",
      "range": "A2:I29",
      "density": 0.82,
      "headerCandidateRows": [2]
    }
  ],
  "outliers": [
    {
      "id": "O1",
      "range": "A30:M31",
      "distanceFromPrimary": { "rows": 1, "cols": 0 }
    }
  ],
  "layout": {
    "tables": { "count": 1, "names": ["Table1"] },
    "freezePanesAt": "A5",
    "autoFilter": false
  },
  "flags": ["OUTLIER_DATA_PRESENT", "HAS_TABLE_OBJECTS", "HAS_FORMULAS"],
  "readHints": {
    "primaryRegionId": "R1",
    "suggestedHeaderScan": "A2:I26",
    "suggestedBodyRead": "A3:I29",
    "suggestedOutlierScans": ["A30:M31"]
  },
  "recommendations": {
    "shouldRunDeep": true,
    "reasons": [{"flag": "OUTLIER_DATA_PRESENT", "severity": "medium"}],
    "nextActions": [{"tool": "explore", "scope": {"sheet": "Sheet1"}, "mode": "deep"}]
  }
}
```

**Example Usage:**
```
User: Explore the active sheet and find its structure
LLM: [Calls explore({"sheet": "ACTIVE"}, mode="quick")]
  Sheet "PO Reconciliation":
  - Primary region: A2:I29 (28 rows, 9 cols, 82% density)
  - Outlier detected at A30:M31 (possibly summary/notes)
  - Has table "Table1", freeze panes at A5
  Recommendation: Run deep mode for better region analysis
```

---

### 10. `delete(workbook_name, sheet_name, delete_type, position, count)` ⭐ NEW

Delete rows or columns at a specific position.

**Parameters:**
- `workbook_name` (string, required): Name of open workbook
- `sheet_name` (string, required): Name of worksheet
- `delete_type` (string, required): `"row"` or `"column"`
- `position` (string, required): Row number (e.g., "5", "5:10") or column (e.g., "C", "C:E")
- `count` (integer, optional): Number to delete (default: 1, ignored if range specified)

**Returns:**
```json
{
  "success": true,
  "action": "rows_deleted",
  "count": 3,
  "at": "5",
  "message": "Deleted 3 row(s) at row 5"
}
```

**Example Usage:**
```
User: Delete rows 10-15 in Sheet1
LLM: [Calls delete("data.xlsx", "Sheet1", "row", "10:15")]
  ✓ Deleted 6 rows at row 10
```

---

### 11. `copy_range(source_workbook, source_sheet, source_range, ...)` ⭐ NEW

Copy data between ranges, sheets, or workbooks.

**Parameters:**
- `source_workbook` (string, required): Name of source workbook
- `source_sheet` (string, required): Name of source worksheet
- `source_range` (string, required): Range to copy (e.g., "A1:D10")
- `target_workbook` (string, optional): Target workbook (defaults to source)
- `target_sheet` (string, optional): Target worksheet (defaults to source)
- `target_cell` (string, optional): Top-left cell of destination (default: "A1")
- `include_formatting` (bool, optional): Copy formatting (default: true)

**Returns:**
```json
{
  "success": true,
  "cells_copied": 40,
  "source": {"workbook": "data.xlsx", "sheet": "Sheet1", "range": "A1:D10"},
  "target": {"workbook": "data.xlsx", "sheet": "Sheet2", "range": "E1:H10"}
}
```

**Example Usage:**
```
User: Copy A1:D10 from Sheet1 to Sheet2 starting at E1
LLM: [Calls copy_range("data.xlsx", "Sheet1", "A1:D10", target_sheet="Sheet2", target_cell="E1")]
  ✓ Copied 40 cells from Sheet1!A1:D10 to Sheet2!E1:H10
```

---

### 12. `sort_range(workbook_name, sheet_name, range, sort_by, has_header)` ⭐ NEW

Sort data in a range by one or more columns.

**Parameters:**
- `workbook_name` (string, required): Name of open workbook
- `sheet_name` (string, required): Name of worksheet
- `range` (string, required): Range to sort (e.g., "A1:D100")
- `sort_by` (array, required): List of sort specs: `[{"column": "B", "order": "asc"}, ...]`
- `has_header` (bool, optional): First row is header (default: true)

**Returns:**
```json
{
  "success": true,
  "range": "A1:D100",
  "rows_sorted": 99,
  "sort_by": [{"column": "B", "order": "asc"}]
}
```

**Example Usage:**
```
User: Sort the data by column B ascending, then column C descending
LLM: [Calls sort_range("data.xlsx", "Sheet1", "A1:D100", 
       [{"column": "B", "order": "asc"}, {"column": "C", "order": "desc"}])]
  ✓ Sorted 99 rows by 2 columns
```

---

### 13. `find_replace(workbook_name, find_value, replace_value, ...)` ⭐ NEW

Find and replace values in a sheet or workbook.

**Parameters:**
- `workbook_name` (string, required): Name of open workbook
- `find_value` (string, required): Value to find
- `replace_value` (string, required): Value to replace with
- `sheet_name` (string, optional): Target worksheet (None = all sheets)
- `match_case` (bool, optional): Match case exactly (default: false)
- `match_entire_cell` (bool, optional): Match entire cell (default: false)
- `range` (string, optional): Specific range (defaults to UsedRange)
- `preview_only` (bool, optional): Count matches without replacing (default: false)

**Returns:**
```json
{
  "success": true,
  "total_matches": 15,
  "total_replacements": 15,
  "results": [{"sheet": "Sheet1", "matches_found": 15, "replacements_made": 15}]
}
```

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
- **Features:** Core operations (read, write, format, sheets)
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

### 📖 Read Operations
- `read()` - Read cells/ranges with filtering
- `search()` - Find and filter data
- `get_unique_values()` - Extract unique values
- `get_current_selection()` - Get active cell

### ✍️ Write Operations
- `write()` - Write with safety guardrails
- `copy_range()` - Copy with formatting
- `sort_range()` - Multi-column sorting
- `find_replace()` - Find and replace with preview

### 🎨 Formatting
- `format()` - Apply predefined or custom formats
- `get_format()` - Read formatting details

### 📋 Sheet Management
- `manage_sheet()` - Add, remove, hide, copy, rename
- `insert()` - Insert rows/columns
- `delete()` - Delete rows/columns

### 📊 Tables (NEW)
- `create_table()` - Create Excel table objects
- `list_tables()` - List all tables
- `delete_table()` - Remove tables

### ⚙️ Advanced Operations (Use with Caution)
- `execute_vba()` - Execute VBA macros (Windows only)
- `capture_sheet()` - Screenshot capture (Windows only)

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
│       ├── server.py              # Main MCP server with 25 tools
│       ├── excel_session.py       # Excel COM session manager
│       ├── validators.py          # Input validation utilities
│       ├── filters.py             # Filter engine for search
│       ├── core/                  # Shared foundation (NEW)
│       │   ├── __init__.py
│       │   ├── connection.py     # COM pooling, batch reads
│       │   ├── errors.py         # ToolError, ErrorCodes
│       │   └── utils.py          # Consolidated utilities
│       ├── tools/                 # Tool implementations (NEW)
│       │   ├── readers.py        # read_cell, read_range
│       │   ├── writers.py        # write_cell, write_range
│       │   ├── formatters.py     # format, get_format
│       │   ├── sheet_mgmt.py     # manage_sheet, insert, delete
│       │   ├── range_ops.py      # copy_range, sort_range, find_replace
│       │   ├── search.py         # search with filters
│       │   └── workbook.py       # list_workbooks, select_range
│       └── inspection/            # Sheet/workbook inspection
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
   - Provides 25 MCP tools
   - Handles tool registration and routing
   - Manages server lifecycle
   - Dual-engine architecture support

2. **Core Module** (`core/`):
   - Thread-local COM connection pooling
   - Batch range reads for performance
   - Centralized error handling
   - Engine abstraction layer (COM + File)

3. **Tools Module** (`tools/`):
   - Modular tool implementations
   - VBA execution, screen capture, table operations
   - Session management for large datasets
   - Clean separation of concerns

4. **ExcelSessionManager** (`excel_session.py`):
   - Connects to running Excel instance
   - Wraps COM operations with async support
   - Provides thread-safe access to Excel

5. **Validators** (`validators.py`):
   - Cell reference format validation
   - Workbook/sheet name validation
   - Range parsing and validation
   - Value type checking

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
- **No Conditional Formatting**: Not supported
- **Charts Limited**: Basic chart support only

### General
- **Large Datasets**: For best performance, use session management for >100 rows
- **File Corruption**: Always backup important files before automation

## Security Considerations

- Tool calls require user approval in most MCP clients
- No remote API calls - local COM operations only
- Read/write operations limited to opened workbooks
- Cell validation prevents malicious input

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

- **GitHub Issues**: [Create an issue](https://github.com/yourusername/ExceLLM/issues)
- **Documentation**: [Read the docs](https://github.com/yourusername/ExceLLM/blob/main/README.md)

## Acknowledgments

- **MCP Team**: For the Model Context Protocol
- **pywin32**: For Windows COM automation
- **openpyxl**: For cross-platform Excel file operations  
- **FastMCP**: For the excellent MCP server framework

## Version History

### 2.0.0-alpha (2026-01-04)
- **Major Update**: Dual-engine architecture (COM + File-based)
- **25 MCP tools** (up from 20)
- **NEW**: VBA execution support
- **NEW**: Screen capture functionality
- **NEW**: Excel table operations
- **NEW**: Cross-platform file mode (Mac/Linux)
- Enhanced session management for large datasets
- Workflow guidance markers
- Comprehensive error handling
