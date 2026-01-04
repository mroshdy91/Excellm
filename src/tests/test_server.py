"""Test stubs for Excel Live MCP server.

This module contains test functions for unit and integration testing.
Run tests with: pytest
"""

import pytest
import asyncio

# Note: These tests will only work when mcp is installed
# and Excel is running with open workbooks


# ============================================================================
# Validator Tests
# ============================================================================


class TestValidators:
    """Test validation utilities."""

    def test_validate_cell_format_valid(self):
        """Test valid cell reference formats."""
        from excellm.validators import validate_cell_format

        # Valid formats
        assert validate_cell_format("A1") is True
        assert validate_cell_format("Z100") is True
        assert validate_cell_format("AA123") is True
        assert validate_cell_format("a1") is True  # Case insensitive
        assert validate_cell_format("z100") is True

    def test_validate_cell_format_invalid(self):
        """Test invalid cell reference formats."""
        from excellm.validators import validate_cell_format

        # Invalid formats
        assert validate_cell_format("1A") is False
        assert validate_cell_format("A") is False
        assert validate_cell_format("A1B2") is False
        assert validate_cell_format("AAA99999") is False
        assert validate_cell_format("") is False

    def test_validate_workbook_name(self):
        """Test workbook name validation."""
        from excellm.validators import validate_workbook_name

        # Valid names
        assert validate_workbook_name("data.xlsx") is True
        assert validate_workbook_name("report 2025.xlsx") is True
        assert validate_workbook_name("test.xls") is True

        # Invalid names
        assert validate_workbook_name("") is False
        assert validate_workbook_name(None) is False
        assert validate_workbook_name("a" * 256) is False

    def test_validate_sheet_name(self):
        """Test sheet name validation."""
        from excellm.validators import validate_sheet_name

        # Valid names
        assert validate_sheet_name("Sheet1") is True
        assert validate_sheet_name("Data Summary") is True

        # Invalid names (Excel max is 31 chars)
        assert validate_sheet_name("a" * 32) is False
        assert validate_sheet_name("") is False

    def test_validate_range_format(self):
        """Test range format validation."""
        from excellm.validators import validate_range_format

        # Valid ranges
        assert validate_range_format("A1:C5") is True
        assert validate_range_format("B2:D10") is True
        assert validate_range_format("Z100:AA123") is True

        # Invalid ranges
        assert validate_range_format("A1") is False
        assert validate_range_format("A1:C5:D10") is False
        assert validate_range_format("1A:C5") is False


# ============================================================================
# Integration Tests (require running Excel)
# ============================================================================


class TestMCPIntegration:
    """Integration tests for MCP server (require Excel running)."""

    @pytest.mark.asyncio
    @pytest.mark.skipif(not _excel_available(), reason="Excel not running")
    async def test_list_workbooks(self):
        """Test listing open workbooks."""
        from excellm.excel_session import ExcelSessionManager

        session = ExcelSessionManager()
        workbooks = await session.list_workbooks()

        assert isinstance(workbooks, list)
        # Should return at least one workbook if Excel is running
        if workbooks:
            assert "name" in workbooks[0]
            assert "sheets" in workbooks[0]

    @pytest.mark.asyncio
    @pytest.mark.skipif(not _excel_available(), reason="Excel not running")
    async def test_read_write_cell(self):
        """Test reading and writing to a cell."""
        from excellm.excel_session import ExcelSessionManager

        session = ExcelSessionManager()

        # Get list of workbooks
        workbooks = await session.list_workbooks()
        if not workbooks:
            pytest.skip("No workbooks available for testing")

        workbook = workbooks[0]
        if not workbook["sheets"]:
            pytest.skip("No sheets available for testing")

        # Read a cell
        sheet_name = workbook["sheets"][0]
        cell = "A1"

        try:
            read_result = await session.read_cell(workbook["name"], sheet_name, cell)
            assert read_result["success"] is True
            assert "value" in read_result
            assert "type" in read_result

            # Write to the same cell
            test_value = "MCP_TEST_VALUE"
            write_result = await session.write_cell(
                workbook["name"], sheet_name, cell, test_value, auto_save=True
            )
            assert write_result["success"] is True
            assert write_result["value"] == test_value
            assert write_result["saved"] is True

        except Exception as e:
            pytest.fail(f"Read/write test failed: {e}")


# ============================================================================
# Helper Functions
# ============================================================================


def _excel_available() -> bool:
    """Check if Excel is running with open workbooks."""
    try:
        import win32com.client as win32

        excel = win32.GetActiveObject("Excel.Application")
        return excel.Workbooks.Count > 0
    except Exception:
        return False


# ============================================================================
# Test Configuration
# ============================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
