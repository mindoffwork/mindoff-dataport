import sys
from pathlib import Path

import pytest

# Ensure src is on path when running without editable install
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_template.xlsx"


def _create_fixture():
    import datetime
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, Color

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws.merge_cells("A1:C2")
    cell = ws["A1"]
    cell.value = "Merged Header"
    cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFFFF")
    cell.fill = PatternFill(patternType="solid", fgColor=Color(rgb="FF003366"))
    cell.alignment = Alignment(horizontal="center", vertical="center")

    ws["A3"].value = "Label"
    ws["A3"].font = Font(name="Calibri", size=11)

    ws["B3"].value = 12345.678
    ws["B3"].number_format = "0.00"
    ws["B3"].alignment = Alignment(horizontal="right")

    ws["C3"].value = "=B3*2"

    ws["A4"].value = datetime.datetime(2024, 6, 15, 9, 30)
    ws["A4"].number_format = "DD/MM/YYYY HH:MM"

    ws["B4"].value = "Styled Text"
    ws["B4"].font = Font(bold=True, italic=True, color="FFCC0000", size=12)

    ws["C4"].value = "Yellow Cell"
    ws["C4"].fill = PatternFill(patternType="solid", fgColor=Color(rgb="FFFFFF00"))

    ws["A5"].value = "Bordered"
    ws["A5"].border = Border(
        top=Side(border_style="thick", color=Color(rgb="FF000000")),
        bottom=Side(border_style="thin", color=Color(rgb="FF000000")),
        left=Side(border_style="medium", color=Color(rgb="FF0000FF")),
        right=Side(border_style="dashed", color=Color(rgb="FFFF0000")),
    )

    ws["B5"].value = "This is a long text that should be wrapped inside the cell"
    ws["B5"].alignment = Alignment(wrap_text=True, horizontal="center", vertical="top")

    ws.merge_cells("D1:E3")
    ws["D1"].value = "Second Merge"
    ws["D1"].fill = PatternFill(patternType="solid", fgColor=Color(rgb="FFCCFFCC"))
    ws["D1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 20.0
    ws.column_dimensions["B"].width = 18.0
    ws.column_dimensions["C"].width = 15.0
    ws.row_dimensions[1].height = 30.0
    ws.row_dimensions[2].height = 30.0
    ws.row_dimensions[5].height = 50.0

    FIXTURE_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(FIXTURE_PATH))


def pytest_configure(config):
    if not FIXTURE_PATH.exists():
        _create_fixture()


@pytest.fixture(scope="session")
def fixture_path() -> str:
    return str(FIXTURE_PATH)


@pytest.fixture(scope="session")
def workbook_schema(fixture_path):
    from mindoff_data_export import extract_template
    return extract_template(fixture_path)
