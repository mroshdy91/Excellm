import pytest
import os
import tempfile
from pathlib import Path


@pytest.fixture
def filename():
    """Create a temporary Excel file for testing."""
    with tempfile.NamedTemporaryFile(mode="w", suffix=".xlsx", delete=False) as tmp:
        tmp_name = tmp.name
    # Close the file so openpyxl can write to it
    tmp_name = tmp_name.replace("\\", "/")
    return tmp_name


@pytest.fixture
def temp_excel_file():
    """Create a temporary Excel file with data for testing."""
    from openpyxl import Workbook

    with tempfile.NamedTemporaryFile(mode="w", suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    # Add sample data
    data = [
        ["Month", "Sales", "Profit", "Region"],
        ["Jan", 100, 20, "North"],
        ["Feb", 120, 25, "South"],
        ["Mar", 110, 22, "North"],
    ]

    for row in data:
        ws.append(row)

    wb.save(tmp_path)

    yield tmp_path

    # Cleanup
    try:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    except Exception:
        pass
