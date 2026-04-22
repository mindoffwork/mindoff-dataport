"""Tests for column_width_mode and row_height_mode."""
import openpyxl
import pytest

from mindoff_data_export import build_template


def _minimal_cell(coord, value="Hello"):
    return {
        "coordinate": coord,
        "value": value,
        "cell_type": "string",
        "number_format": None,
        "font": {"name": "Calibri", "size": 11.0, "bold": False, "italic": False, "underline": None, "color": None},
        "fill": {"bg_color": None},
        "alignment": {"horizontal": None, "vertical": None, "wrap_text": False},
        "borders": {
            "top": {"style": None, "color": None},
            "bottom": {"style": None, "color": None},
            "left": {"style": None, "color": None},
            "right": {"style": None, "color": None},
        },
        "merged": False,
        "merge_anchor": None,
    }


def _minimal_sheet(cells_dict, **extra):
    return {
        "name": "Sheet1",
        "dimensions": "A1:C3",
        "merged_regions": [],
        "column_widths": {},
        "row_heights": {},
        "cells": cells_dict,
        **extra,
    }


def _build_and_reload(schema, tmp_path):
    out = str(tmp_path / "out.xlsx")
    build_template(schema, out)
    return openpyxl.load_workbook(out)


# ---------------------------------------------------------------------------
# even column mode
# ---------------------------------------------------------------------------

def test_even_columns_applies_uniform_width(tmp_path):
    cells = {c: _minimal_cell(c) for c in ["A1", "B1", "C1"]}
    schema = {"sheets": [_minimal_sheet(cells, column_width_mode="even", default_column_width=25.0)]}
    wb = _build_and_reload(schema, tmp_path)
    ws = wb.active
    for col in ["A", "B", "C"]:
        assert ws.column_dimensions[col].width == pytest.approx(25.0)


def test_even_columns_default_width_when_not_specified(tmp_path):
    cells = {"A1": _minimal_cell("A1")}
    schema = {"sheets": [_minimal_sheet(cells, column_width_mode="even")]}
    wb = _build_and_reload(schema, tmp_path)
    ws = wb.active
    assert ws.column_dimensions["A"].width == pytest.approx(15.0)


# ---------------------------------------------------------------------------
# hug column mode
# ---------------------------------------------------------------------------

def test_hug_columns_wider_for_longer_content(tmp_path):
    cells = {
        "A1": _minimal_cell("A1", value="Hi"),
        "B1": _minimal_cell("B1", value="A very long header label here"),
    }
    schema = {"sheets": [_minimal_sheet(cells, column_width_mode="hug")]}
    wb = _build_and_reload(schema, tmp_path)
    ws = wb.active
    assert ws.column_dimensions["B"].width > ws.column_dimensions["A"].width


def test_hug_columns_bold_factor(tmp_path):
    plain = _minimal_cell("A1", value="SameText")
    bold = _minimal_cell("B1", value="SameText")
    bold["font"] = dict(bold["font"])
    bold["font"]["bold"] = True
    cells = {"A1": plain, "B1": bold}
    schema = {"sheets": [_minimal_sheet(cells, column_width_mode="hug")]}
    wb = _build_and_reload(schema, tmp_path)
    ws = wb.active
    assert ws.column_dimensions["B"].width > ws.column_dimensions["A"].width


# ---------------------------------------------------------------------------
# even row mode
# ---------------------------------------------------------------------------

def test_even_rows_applies_uniform_height(tmp_path):
    cells = {f"A{r}": _minimal_cell(f"A{r}") for r in range(1, 4)}
    schema = {"sheets": [_minimal_sheet(cells, row_height_mode="even", default_row_height=30.0)]}
    wb = _build_and_reload(schema, tmp_path)
    ws = wb.active
    for row in range(1, 4):
        assert ws.row_dimensions[row].height == pytest.approx(30.0)


# ---------------------------------------------------------------------------
# hug row mode
# ---------------------------------------------------------------------------

def test_hug_rows_larger_for_bigger_font(tmp_path):
    small = _minimal_cell("A1", value="text")
    large = _minimal_cell("A2", value="text")
    large["font"] = dict(large["font"])
    large["font"]["size"] = 24.0
    cells = {"A1": small, "A2": large}
    schema = {"sheets": [_minimal_sheet(cells, row_height_mode="hug")]}
    wb = _build_and_reload(schema, tmp_path)
    ws = wb.active
    assert ws.row_dimensions[2].height > ws.row_dimensions[1].height


# ---------------------------------------------------------------------------
# backward compat: default "fixed" mode unchanged
# ---------------------------------------------------------------------------

def test_fixed_mode_uses_explicit_widths(tmp_path):
    cells = {"A1": _minimal_cell("A1")}
    schema = {"sheets": [{
        "name": "Sheet1", "dimensions": "A1:A1",
        "merged_regions": [], "cells": cells,
        "column_widths": {"A": 42.0}, "row_heights": {},
    }]}
    wb = _build_and_reload(schema, tmp_path)
    ws = wb.active
    assert ws.column_dimensions["A"].width == pytest.approx(42.0)
