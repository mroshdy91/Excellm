# Dual-Engine Architecture Summary

## Overview

ExceLLM now supports **two engines** for Excel operations:

### 1. COM Engine (Windows + Live Excel)
- **Platform:** Windows only
- **Requires:** Excel running
- **Mode:** Live operations on open workbooks
- **Features:** Full feature set (VBA, screen capture, etc.)

### 2. File Engine (Cross-platform)
- **Platform:** Windows, Mac, Linux
- **Requires:** openpyxl library
- **Mode:** File-based operations (no Excel needed)
- **Features:** Core operations (read, write, format, sheets)

## Architecture

```
ExcelEngine (ABC)
├── COMEngine (Windows/Live)
│   └── Wraps existing ExceLLM COM functionality
└── FileEngine (Cross-platform)
    └── Uses openpyxl for file operations
```

## Auto-Detection

The `EngineFactory` automatically selects the best engine:

1. If `workbook_path` looks like a file path → **FileEngine**
2. If Windows + Excel running → **COMEngine**
3. Otherwise → **FileEngine** (fallback)

##Usage

```python
from core.engine import EngineFactory

# Auto-select engine
engine = EngineFactory.create_engine(workbook_path="C:/data/file.xlsx")

# Manual selection
engine = EngineFactory.create_engine(prefer_live=False)  # Force file mode

# Use engine
data = engine.read_range("file.xlsx", "Sheet1", "A1:C10")
engine.write_range("file.xlsx", "Sheet1", "A1", [[1, 2, 3]])
```

## Feature Matrix

| Feature | COM Engine | File Engine |
|---------|------------|-------------|
| Read cells | ✅ | ✅ |
| Write cells | ✅ | ✅ |
| Format cells | ✅ | ⚠️ Limited |
| Sheet management | ✅ | ✅ |
| VBA execution | ✅ | ❌ |
| Screen capture | ✅ | ❌ |
| Table objects | ✅ | ✅ |
| Live Excel | ✅ | ❌ |

## Files Created

1. `core/engine.py` - Abstract base class + factory
2. `core/com_engine.py` - Windows/COM implementation
3. `core/file_engine.py` - Cross-platform/openpyxl implementation

---

**Status:** ✅ Core architecture complete, ready for tool integration
