"""Tests for template variable rendering (renderer.py)."""
import datetime

import pytest

from mindoff_data_export import get_template_inputs, render_schema
from mindoff_data_export.renderer import _infer_cell_type, _to_rows

# Section 1 Types

# Section 2 Constants


def _cell(coord, value):
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


def _sheet(name: str, cells_dict, dims="A1:C3"):
    return {
        "name": name,
        "dimensions": dims,
        "merged_regions": [],
        "column_widths": {},
        "row_heights": {},
        "cells": cells_dict,
    }


def _schema(*sheets):
    return {"sheets": list(sheets)}


# Section 4 Public API


def test_get_template_inputs_returns_sheet_scoped_contract():
    schema = _schema(
        _sheet("Sheet 1", {"A1": _cell("A1", "{{customer_name:string}}")}),
        _sheet("Sheet 2", {"A1": _cell("A1", "{{customer_name:string}}")}),
    )

    result = get_template_inputs(schema)

    assert result == {
        "Sheet 1": {"customer_name": "string"},
        "Sheet 2": {"customer_name": "string"},
    }


def test_get_template_inputs_returns_dynamic_wildcard_contract():
    schema = _schema(
        _sheet("{{sheet_2}}", {"A1": _cell("A1", "{{customer_name:string}}")}),
    )

    result = get_template_inputs(schema)

    assert result == {
        "sheet_2": {
            "*": {"customer_name": "string"},
        }
    }


def test_get_template_inputs_raises_for_conflicting_placeholder_types_in_same_scope():
    schema = _schema(
        _sheet(
            "Sheet 1",
            {
                "A1": _cell("A1", "{{customer_name:string}}"),
                "B1": _cell("B1", "{{customer_name:number}}"),
            },
        )
    )

    with pytest.raises(ValueError, match="Conflicting placeholder types"):
        get_template_inputs(schema)


def test_render_schema_static_sheets_use_per_sheet_data():
    schema = _schema(
        _sheet("Sheet 1", {"A1": _cell("A1", "{{customer_name:string}}")}),
        _sheet("Sheet 2", {"A1": _cell("A1", "{{customer_name:string}}")}),
    )

    result = render_schema(
        schema,
        {
            "Sheet 1": {"customer_name": "John Doe"},
            "Sheet 2": {"customer_name": "Jane Doe"},
        },
    )

    assert result["sheets"][0]["cells"]["A1"]["value"] == "John Doe"
    assert result["sheets"][1]["cells"]["A1"]["value"] == "Jane Doe"


def test_render_schema_expands_dynamic_sheet_name_in_input_order():
    schema = _schema(_sheet("{{sheet_1}}", {"A1": _cell("A1", "{{customer_name:string}}")}))

    result = render_schema(
        schema,
        {
            "sheet_1": {
                "Sheet Name 1": {"customer_name": "John Doe"},
                "Sheet Name 3": {"customer_name": "Jane Doe"},
            }
        },
    )

    assert [sheet["name"] for sheet in result["sheets"]] == ["Sheet Name 1", "Sheet Name 3"]
    assert result["sheets"][0]["cells"]["A1"]["value"] == "John Doe"
    assert result["sheets"][1]["cells"]["A1"]["value"] == "Jane Doe"


def test_render_schema_mixed_static_and_dynamic_sheets():
    schema = _schema(
        _sheet("Sheet 1", {"A1": _cell("A1", "{{customer_name:string}}")}),
        _sheet("{{sheet_2}}", {"A1": _cell("A1", "{{customer_name:string}}")}),
    )

    result = render_schema(
        schema,
        {
            "Sheet 1": {"customer_name": "John Doe"},
            "sheet_2": {
                "Sheet Name 2": {"customer_name": "Alice"},
                "Sheet Name 3": {"customer_name": "Bob"},
            },
        },
    )

    assert [sheet["name"] for sheet in result["sheets"]] == [
        "Sheet 1",
        "Sheet Name 2",
        "Sheet Name 3",
    ]


def test_render_schema_rejects_legacy_flat_payload():
    schema = _schema(_sheet("Sheet 1", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(KeyError, match="requires sheet 'Sheet 1'"):
        render_schema(schema, {"name": "Alice"})


def test_render_schema_raises_for_missing_static_sheet_key():
    schema = _schema(_sheet("Sheet 1", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(KeyError, match="requires sheet 'Sheet 1'"):
        render_schema(schema, {})


def test_render_schema_raises_for_missing_dynamic_sheet_group():
    schema = _schema(_sheet("{{sheet_2}}", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(KeyError, match="requires dynamic sheet group 'sheet_2'"):
        render_schema(schema, {})


def test_render_schema_raises_for_non_dict_dynamic_group_payload():
    schema = _schema(_sheet("{{sheet_2}}", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(TypeError, match="must be an object/dict"):
        render_schema(schema, {"sheet_2": "bad"})


def test_render_schema_raises_for_non_dict_dynamic_sheet_payload():
    schema = _schema(_sheet("{{sheet_2}}", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(TypeError, match="Data for dynamic sheet"):
        render_schema(schema, {"sheet_2": {"Sheet A": "bad"}})


def test_render_schema_raises_for_duplicate_output_sheet_names():
    schema = _schema(
        _sheet("Sheet 1", {"A1": _cell("A1", "{{name:string}}")}),
        _sheet("{{sheet_2}}", {"A1": _cell("A1", "{{name:string}}")}),
    )

    with pytest.raises(ValueError, match="Duplicate output sheet name"):
        render_schema(
            schema,
            {
                "Sheet 1": {"name": "Alice"},
                "sheet_2": {"Sheet 1": {"name": "Bob"}},
            },
        )


polars = pytest.importorskip("polars", reason="polars not installed")


def test_render_schema_dataframe_headers_accepts_list_input_with_sheet_scope():
    schema = _schema(
        _sheet("Sheet 1", {"A1": _cell("A1", "{{hdr:dataframe-headers}}")}, dims="A1:C1")
    )

    result = render_schema(schema, {"Sheet 1": {"hdr": ["Item", "Qty", "Price"]}})
    cells = result["sheets"][0]["cells"]

    assert cells["A1"]["value"] == "Item"
    assert cells["B1"]["value"] == "Qty"
    assert cells["C1"]["value"] == "Price"


def test_render_schema_dataframe_content_no_headers_with_sheet_scope():
    df = polars.DataFrame({"X": [1, 2], "Y": [3, 4]})
    schema = _schema(
        _sheet("Sheet 1", {"A1": _cell("A1", "{{tbl:dataframe-content}}")}, dims="A1:B2")
    )

    result = render_schema(schema, {"Sheet 1": {"tbl": df}})
    cells = result["sheets"][0]["cells"]

    assert cells["A1"]["value"] == 1
    assert cells["B1"]["value"] == 3
    assert cells["A2"]["value"] == 2


def test_render_schema_does_not_mutate_original():
    schema = _schema(_sheet("Sheet 1", {"A1": _cell("A1", "{{x:string}}")}))

    render_schema(schema, {"Sheet 1": {"x": "changed"}})

    assert schema["sheets"][0]["cells"]["A1"]["value"] == "{{x:string}}"


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


def test_to_rows_pandas_or_polars_like_contract():
    df = polars.DataFrame({"A": [1], "B": [2]})
    cols, rows = _to_rows(df)
    assert cols == ["A", "B"]
    assert rows == [(1, 2)]
