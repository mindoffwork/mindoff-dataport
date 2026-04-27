"""Tests for column_width_mode and row_height_mode."""

import openpyxl
import pytest

from mindoff_dataport import mode

# §1 Constants & Exceptions

# §2 Classes and Sub Classes

# §3 Private Helper Functions


def _minimal_cell(coord, value="Hello"):
    return {
        "coordinate": coord,
        "value": value,
        "cell_type": "string",
        "number_format": None,
        "font": {
            "name": "Calibri",
            "size": 11.0,
            "bold": False,
            "italic": False,
            "underline": None,
            "color": None,
        },
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


def _build_and_reload(schema, managed_tmp_dir):
    out = str(managed_tmp_dir / "out.xlsx")
    bundle = mode.compile(schema, {"Sheet1": {}})
    mode.export(bundle, out)
    return openpyxl.load_workbook(out)


# §4 Public Functions


def test_even_columns_applies_uniform_width(managed_tmp_dir):
    cells = {c: _minimal_cell(c) for c in ["A1", "B1", "C1"]}
    schema = {
        "sheets": [
            _minimal_sheet(cells, column_width_mode="even", default_column_width=25.0)
        ]
    }
    wb = _build_and_reload(schema, managed_tmp_dir)
    ws = wb.active
    for col in ["A", "B", "C"]:
        assert ws.column_dimensions[col].width == pytest.approx(25.0)
    wb.close()


def test_even_columns_default_width_when_not_specified(managed_tmp_dir):
    cells = {"A1": _minimal_cell("A1")}
    schema = {"sheets": [_minimal_sheet(cells, column_width_mode="even")]}
    wb = _build_and_reload(schema, managed_tmp_dir)
    ws = wb.active
    assert ws.column_dimensions["A"].width == pytest.approx(15.0)
    wb.close()


def test_hug_columns_wider_for_longer_content(managed_tmp_dir):
    cells = {
        "A1": _minimal_cell("A1", value="Hi"),
        "B1": _minimal_cell("B1", value="A very long header label here"),
    }
    schema = {"sheets": [_minimal_sheet(cells, column_width_mode="hug")]}
    wb = _build_and_reload(schema, managed_tmp_dir)
    ws = wb.active
    assert ws.column_dimensions["B"].width > ws.column_dimensions["A"].width
    wb.close()


def test_hug_columns_bold_factor(managed_tmp_dir):
    plain = _minimal_cell("A1", value="SameText")
    bold = _minimal_cell("B1", value="SameText")
    bold["font"] = dict(bold["font"])
    bold["font"]["bold"] = True
    cells = {"A1": plain, "B1": bold}
    schema = {"sheets": [_minimal_sheet(cells, column_width_mode="hug")]}
    wb = _build_and_reload(schema, managed_tmp_dir)
    ws = wb.active
    assert ws.column_dimensions["B"].width > ws.column_dimensions["A"].width
    wb.close()


def test_even_rows_applies_uniform_height(managed_tmp_dir):
    cells = {f"A{r}": _minimal_cell(f"A{r}") for r in range(1, 4)}
    schema = {
        "sheets": [
            _minimal_sheet(cells, row_height_mode="even", default_row_height=30.0)
        ]
    }
    wb = _build_and_reload(schema, managed_tmp_dir)
    ws = wb.active
    for row in range(1, 4):
        assert ws.row_dimensions[row].height == pytest.approx(30.0)
    wb.close()


def test_hug_rows_larger_for_bigger_font(managed_tmp_dir):
    small = _minimal_cell("A1", value="text")
    large = _minimal_cell("A2", value="text")
    large["font"] = dict(large["font"])
    large["font"]["size"] = 24.0
    cells = {"A1": small, "A2": large}
    schema = {"sheets": [_minimal_sheet(cells, row_height_mode="hug")]}
    wb = _build_and_reload(schema, managed_tmp_dir)
    ws = wb.active
    assert ws.row_dimensions[2].height > ws.row_dimensions[1].height
    wb.close()


def test_fixed_mode_uses_explicit_widths(managed_tmp_dir):
    cells = {"A1": _minimal_cell("A1")}
    schema = {
        "sheets": [
            {
                "name": "Sheet1",
                "dimensions": "A1:A1",
                "merged_regions": [],
                "cells": cells,
                "column_widths": {"A": 42.0},
                "row_heights": {},
            }
        ]
    }
    wb = _build_and_reload(schema, managed_tmp_dir)
    ws = wb.active
    assert ws.column_dimensions["A"].width == pytest.approx(42.0)
    wb.close()


# §5 Entrypoints
