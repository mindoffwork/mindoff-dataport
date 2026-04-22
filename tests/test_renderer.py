"""Tests for template variable rendering (renderer.py)."""
import datetime
import pytest

from mindoff_data_export import get_template_inputs, render_schema
from mindoff_data_export.renderer import _infer_cell_type, _to_rows


# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------

def _cell(coord, value):
    return {
        "coordinate": coord, "value": value, "cell_type": "string",
        "number_format": None,
        "font": {"name": "Calibri", "size": 11.0, "bold": False, "italic": False, "underline": None, "color": None},
        "fill": {"bg_color": None},
        "alignment": {"horizontal": None, "vertical": None, "wrap_text": False},
        "borders": {
            "top": {"style": None, "color": None}, "bottom": {"style": None, "color": None},
            "left": {"style": None, "color": None}, "right": {"style": None, "color": None},
        },
        "merged": False, "merge_anchor": None,
    }


def _schema(cells_dict, dims="A1:C3"):
    return {
        "sheets": [{
            "name": "Sheet1", "dimensions": dims,
            "merged_regions": [], "column_widths": {}, "row_heights": {},
            "cells": cells_dict,
        }]
    }


# ---------------------------------------------------------------------------
# get_template_inputs
# ---------------------------------------------------------------------------

def test_get_template_inputs_finds_scalar_placeholders():
    schema = _schema({"A1": _cell("A1", "{{name:string}}"), "B1": _cell("B1", "static")})
    result = get_template_inputs(schema)
    assert result == {"name": "string"}


def test_get_template_inputs_finds_multiple_types():
    schema = _schema({
        "A1": _cell("A1", "{{name:string}}"),
        "B1": _cell("B1", "{{count:number}}"),
        "C1": _cell("C1", "{{rows:dataframe-headers}}"),
    })
    result = get_template_inputs(schema)
    assert result == {"name": "string", "count": "number", "rows": "dataframe-headers"}


def test_get_template_inputs_ignores_legacy_markers():
    # {{key}} and {{table:key}} are legacy; should not appear in output
    schema = _schema({"A1": _cell("A1", "{{legacy}}"), "B1": _cell("B1", "{{table:sales}}")})
    result = get_template_inputs(schema)
    assert result == {}


def test_get_template_inputs_ignores_non_string_values():
    schema = _schema({"A1": _cell("A1", 42), "B1": _cell("B1", None)})
    result = get_template_inputs(schema)
    assert result == {}


# ---------------------------------------------------------------------------
# render_schema — scalar substitution
# ---------------------------------------------------------------------------

def test_render_schema_replaces_whole_cell_string():
    schema = _schema({"A1": _cell("A1", "{{name:string}}")})
    result = render_schema(schema, {"name": "Alice"})
    assert result["sheets"][0]["cells"]["A1"]["value"] == "Alice"


def test_render_schema_replaces_embedded_placeholder():
    schema = _schema({"A1": _cell("A1", "Hello {{name:string}}, welcome!")})
    result = render_schema(schema, {"name": "Bob"})
    assert result["sheets"][0]["cells"]["A1"]["value"] == "Hello Bob, welcome!"


def test_render_schema_number_value_preserved():
    schema = _schema({"A1": _cell("A1", "{{count:number}}")})
    result = render_schema(schema, {"count": 42})
    cell = result["sheets"][0]["cells"]["A1"]
    assert cell["value"] == 42
    assert cell["cell_type"] == "number"


def test_render_schema_does_not_mutate_original():
    schema = _schema({"A1": _cell("A1", "{{x:string}}")})
    render_schema(schema, {"x": "changed"})
    assert schema["sheets"][0]["cells"]["A1"]["value"] == "{{x:string}}"


def test_render_schema_leaves_non_placeholder_cells_unchanged():
    schema = _schema({"A1": _cell("A1", "static"), "B1": _cell("B1", 99)})
    result = render_schema(schema, {})
    assert result["sheets"][0]["cells"]["A1"]["value"] == "static"
    assert result["sheets"][0]["cells"]["B1"]["value"] == 99


# ---------------------------------------------------------------------------
# render_schema — validation errors
# ---------------------------------------------------------------------------

def test_render_schema_raises_keyerror_for_missing_key():
    schema = _schema({"A1": _cell("A1", "{{name:string}}")})
    with pytest.raises(KeyError, match="name"):
        render_schema(schema, {})


def test_render_schema_raises_typeerror_for_wrong_type():
    schema = _schema({"A1": _cell("A1", "{{count:number}}")})
    with pytest.raises(TypeError, match="count"):
        render_schema(schema, {"count": "not-a-number"})


def test_render_schema_raises_typeerror_for_non_dataframe():
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-headers}}")})
    with pytest.raises(TypeError, match="rows"):
        render_schema(schema, {"rows": [1, 2, 3]})


# ---------------------------------------------------------------------------
# render_schema — dataframe expansion (polars)
# ---------------------------------------------------------------------------

polars = pytest.importorskip("polars", reason="polars not installed")


def test_render_schema_dataframe_headers_writes_header_and_data():
    df = polars.DataFrame({"Name": ["Alice", "Bob"], "Age": [30, 25]})
    schema = _schema({"A1": _cell("A1", "{{people:dataframe-headers}}")}, dims="A1:B3")
    result = render_schema(schema, {"people": df})
    cells = result["sheets"][0]["cells"]

    # Header row at A1, B1
    assert cells["A1"]["value"] == "Name"
    assert cells["B1"]["value"] == "Age"
    assert cells["A1"]["font"]["bold"] is True

    # Data rows
    assert cells["A2"]["value"] == "Alice"
    assert cells["B2"]["value"] == 30
    assert cells["A3"]["value"] == "Bob"


def test_render_schema_dataframe_headers_not_overwritten_by_empty_template_cells():
    df = polars.DataFrame({"Item": ["A"], "Qty": [2], "Unit Price": [10.0]})
    schema = _schema(
        {
            "A1": _cell("A1", "{{rows:dataframe-headers}}"),
            "B1": _cell("B1", None),  # extracted template stubs that previously overwrote headers
            "C1": _cell("C1", None),
        },
        dims="A1:C2",
    )
    result = render_schema(schema, {"rows": df})
    cells = result["sheets"][0]["cells"]
    assert cells["A1"]["value"] == "Item"
    assert cells["B1"]["value"] == "Qty"
    assert cells["C1"]["value"] == "Unit Price"


def test_render_schema_dataframe_data_no_headers():
    df = polars.DataFrame({"X": [1, 2], "Y": [3, 4]})
    schema = _schema({"A1": _cell("A1", "{{tbl:dataframe-data}}")}, dims="A1:B2")
    result = render_schema(schema, {"tbl": df})
    cells = result["sheets"][0]["cells"]

    # A1 should be first data row, no header
    assert cells["A1"]["value"] == 1
    assert cells["B1"]["value"] == 3
    assert cells["A2"]["value"] == 2


def test_render_schema_lazyframe_is_collected():
    lf = polars.DataFrame({"V": [10, 20]}).lazy()
    schema = _schema({"A1": _cell("A1", "{{vals:dataframe-data}}")}, dims="A1:A2")
    result = render_schema(schema, {"vals": lf})
    cells = result["sheets"][0]["cells"]
    assert cells["A1"]["value"] == 10
    assert cells["A2"]["value"] == 20


def test_render_schema_dataframe_inherits_anchor_style():
    df = polars.DataFrame({"Col": ["x"]})
    anchor = _cell("A1", "{{d:dataframe-data}}")
    anchor["fill"] = {"bg_color": "FFFF0000"}
    schema = _schema({"A1": anchor})
    result = render_schema(schema, {"d": df})
    assert result["sheets"][0]["cells"]["A1"]["fill"]["bg_color"] == "FFFF0000"


def test_render_schema_dimensions_extended_by_dataframe():
    df = polars.DataFrame({"A": list(range(10))})
    schema = _schema({"A1": _cell("A1", "{{big:dataframe-data}}")}, dims="A1:A1")
    result = render_schema(schema, {"big": df})
    assert result["sheets"][0]["dimensions"] == "A1:A10"


# ---------------------------------------------------------------------------
# _infer_cell_type helper
# ---------------------------------------------------------------------------

def test_infer_cell_type_number():
    assert _infer_cell_type(42) == "number"
    assert _infer_cell_type(3.14) == "number"


def test_infer_cell_type_string():
    assert _infer_cell_type("hello") == "string"


def test_infer_cell_type_date():
    assert _infer_cell_type(datetime.datetime(2024, 1, 1)) == "date"


def test_infer_cell_type_none():
    assert _infer_cell_type(None) == "empty"
