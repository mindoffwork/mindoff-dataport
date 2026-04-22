import json

# §1 Constants & Exceptions

# §2 Classes and Sub Classes

# §3 Private Helper Functions

# §4 Public Functions


def test_extract_returns_workbook_schema(workbook_schema):
    assert "sheets" in workbook_schema
    assert len(workbook_schema["sheets"]) >= 1


def test_sheet_has_required_keys(workbook_schema):
    sheet = workbook_schema["sheets"][0]
    for key in (
        "name",
        "dimensions",
        "merged_regions",
        "column_widths",
        "row_heights",
        "cells",
    ):
        assert key in sheet, f"Missing key: {key}"


def test_merged_regions_captured(workbook_schema):
    sheet = workbook_schema["sheets"][0]
    assert len(sheet["merged_regions"]) >= 2
    assert "A1:C2" in sheet["merged_regions"]
    assert "D1:E3" in sheet["merged_regions"]


def test_merge_anchor_cell_tagged(workbook_schema):
    cells = workbook_schema["sheets"][0]["cells"]
    assert cells["A1"]["merged"] is True
    assert cells["A1"]["merge_anchor"] == "A1"


def test_non_anchor_merged_cell_tagged(workbook_schema):
    cells = workbook_schema["sheets"][0]["cells"]
    assert cells["B1"]["merged"] is True
    assert cells["B1"]["merge_anchor"] == "A1"
    assert cells["B1"]["value"] is None


def test_non_merged_cell_not_tagged(workbook_schema):
    cells = workbook_schema["sheets"][0]["cells"]
    assert cells["A3"]["merged"] is False
    assert cells["A3"]["merge_anchor"] is None


def test_plain_string_cell(workbook_schema):
    cell = workbook_schema["sheets"][0]["cells"]["A3"]
    assert cell["value"] == "Label"
    assert cell["cell_type"] == "string"


def test_number_cell_with_format(workbook_schema):
    cell = workbook_schema["sheets"][0]["cells"]["B3"]
    assert cell["cell_type"] == "number"
    assert cell["number_format"] == "0.00"


def test_formula_cell(workbook_schema):
    cell = workbook_schema["sheets"][0]["cells"]["C3"]
    assert cell["cell_type"] == "formula"
    assert str(cell["value"]).startswith("=")


def test_date_cell(workbook_schema):
    cell = workbook_schema["sheets"][0]["cells"]["A4"]
    assert cell["cell_type"] == "date"
    assert "2024" in str(cell["value"])


def test_font_bold_italic(workbook_schema):
    font = workbook_schema["sheets"][0]["cells"]["B4"]["font"]
    assert font["bold"] is True
    assert font["italic"] is True
    assert font["color"] is not None


def test_fill_color_captured(workbook_schema):
    fill = workbook_schema["sheets"][0]["cells"]["C4"]["fill"]
    assert fill["bg_color"] is not None
    assert len(fill["bg_color"]) == 8


def test_borders_captured(workbook_schema):
    borders = workbook_schema["sheets"][0]["cells"]["A5"]["borders"]
    assert borders["top"]["style"] == "thick"
    assert borders["left"]["style"] == "medium"
    assert borders["right"]["style"] == "dashed"


def test_alignment_captured(workbook_schema):
    alignment = workbook_schema["sheets"][0]["cells"]["B5"]["alignment"]
    assert alignment["wrap_text"] is True
    assert alignment["horizontal"] == "center"


def test_column_widths_captured(workbook_schema):
    widths = workbook_schema["sheets"][0]["column_widths"]
    assert "A" in widths
    assert abs(widths["A"] - 20.0) < 0.5


def test_row_heights_captured(workbook_schema):
    heights = workbook_schema["sheets"][0]["row_heights"]
    assert "1" in heights
    assert abs(heights["1"] - 30.0) < 0.5


def test_schema_is_json_serializable(workbook_schema):
    json.dumps(workbook_schema)


# §5 Entrypoints
