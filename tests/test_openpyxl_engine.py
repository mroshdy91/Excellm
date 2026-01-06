import os
import sys
import logging
from openpyxl import Workbook

# Adjust path to find excellm package
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("test_results.txt", mode="w", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("test_openpyxl")

from excellm.tools.chart import create_chart_sync
from excellm.tools.pivot import create_pivot_table_sync


def setup_test_file(filename):
    """Create a dummy Excel file with data."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    # Add data
    data = [
        ["Month", "Sales", "Profit", "Region"],
        ["Jan", 100, 20, "North"],
        ["Feb", 120, 25, "South"],
        ["Mar", 110, 22, "North"],
        ["Apr", 130, 30, "East"],
        ["May", 140, 35, "West"],
    ]

    for row in data:
        ws.append(row)

    wb.save(filename)
    logger.info(f"Created test file: {filename}")
    return filename


def test_chart_creation(temp_excel_file):
    logger.info("Testing OpenPyXL Chart Creation...")
    result = create_chart_sync(
        workbook_name=temp_excel_file,  # Full path, ensuring it's treated as file if not open
        sheet_name="Data",
        data_range="A1:C6",
        chart_type="bar",
        target_cell="E1",
        title="Sales & Profit by Month",
        x_axis_title="Month",
        y_axis_title="Amount",
    )

    if result.get("success") and result.get("engine") == "openpyxl":
        logger.info("✅ Chart created successfully using openpyxl engine")
        return True
    else:
        logger.error(f"❌ Chart creation failed or used wrong engine: {result}")
        return False


def test_pivot_creation(filename):
    logger.info("Testing OpenPyXL Pivot Table Creation...")
    result = create_pivot_table_sync(
        workbook_name=filename,
        sheet_name="Data",
        data_range="Data!A1:D6",
        rows=["Region"],
        values=["Sales", "Profit"],
        agg_func="sum",
        target_sheet="PivotSummary",
        target_cell="A1",
        table_name="RegionStats",
    )

    if result.get("success") and result.get("engine") == "openpyxl (static)":
        logger.info("✅ Pivot table created successfully using openpyxl engine")
        return True
    else:
        logger.error(f"❌ Pivot creation failed or used wrong engine: {result}")
        return False


if __name__ == "__main__":
    test_file = os.path.abspath("test_openpyxl_engine.xlsx")

    try:
        setup_test_file(test_file)

        # Test Chart
        chart_ok = test_chart_creation(test_file)

        # Test Pivot
        pivot_ok = test_pivot_creation(test_file)

        if chart_ok and pivot_ok:
            logger.info("🎉 All OpenPyXL tests passed!")
        else:
            logger.error("⚠️ Some tests failed.")

    finally:
        # Cleanup
        if os.path.exists(test_file):
            try:
                # os.remove(test_file)
                logger.info(f"Test file kept for inspection: {test_file}")
            except:
                pass
