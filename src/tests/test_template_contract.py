"""Tests for template input contracts and payload validation."""

import datetime

import pytest

from mindoff_dataport import mo_dataport
from mindoff_dataport.template_contract import _infer_cell_type, get_template_inputs

# §1. Constants & Exceptions

# §2. Classes and Sub Classes

# §3. Private Helper Functions


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


def _sheet(name: str, cells_dict):
    return {
        "name": name,
        "dimensions": "A1:C3",
        "merged_regions": [],
        "column_widths": {},
        "row_heights": {},
        "cells": cells_dict,
    }


def _schema(*sheets):
    return {"sheets": list(sheets)}


# §4. Public Functions


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


def test_get_template_inputs_returns_repeat_contract():
    schema = _schema(
        _sheet(
            "Sheet 1",
            {
                "A1": _cell("A1", "{{reports:repeat-start}}"),
                "A2": _cell("A2", "{{customer_name:string}}"),
                "A3": _cell("A3", "{{line_items:dataframe}}"),
                "A4": _cell("A4", "{{reports:repeat-end}}"),
            },
        )
    )

    result = get_template_inputs(schema)

    assert result == {
        "Sheet 1": {
            "reports": [
                {"customer_name": "string", "line_items": "dataframe"},
            ]
        }
    }


def test_get_template_inputs_returns_sibling_repeat_contracts():
    schema = _schema(
        _sheet(
            "Sheet 1",
            {
                "A1": _cell("A1", "{{customers:repeat-start}}"),
                "A2": _cell("A2", "{{customer_name:string}}"),
                "A3": _cell("A3", "{{customers:repeat-end}}"),
                "A5": _cell("A5", "{{summaries:repeat-start}}"),
                "A6": _cell("A6", "{{region:string}}"),
                "A7": _cell("A7", "{{totals:dataframe}}"),
                "A8": _cell("A8", "{{summaries:repeat-end}}"),
            },
        )
    )

    result = get_template_inputs(schema)

    assert result == {
        "Sheet 1": {
            "customers": [{"customer_name": "string"}],
            "summaries": [{"region": "string", "totals": "dataframe"}],
        }
    }


def test_get_template_inputs_rejects_mismatched_repeat_markers():
    schema = _schema(
        _sheet(
            "Sheet 1",
            {
                "A1": _cell("A1", "{{reports:repeat-start}}"),
                "A2": _cell("A2", "{{items:repeat-end}}"),
            },
        )
    )

    with pytest.raises(ValueError, match="same key"):
        get_template_inputs(schema)


def test_get_template_inputs_rejects_missing_repeat_end():
    schema = _schema(
        _sheet("Sheet 1", {"A1": _cell("A1", "{{reports:repeat-start}}")})
    )

    with pytest.raises(ValueError, match="repeat-end marker"):
        get_template_inputs(schema)


def test_get_template_inputs_rejects_nested_repeat_sections():
    schema = _schema(
        _sheet(
            "Sheet 1",
            {
                "A1": _cell("A1", "{{outer:repeat-start}}"),
                "A2": _cell("A2", "{{inner:repeat-start}}"),
                "A3": _cell("A3", "{{inner:repeat-end}}"),
                "A4": _cell("A4", "{{outer:repeat-end}}"),
            },
        )
    )

    with pytest.raises(ValueError, match="Nested repeat"):
        get_template_inputs(schema)


def test_get_template_inputs_rejects_duplicate_repeat_keys():
    schema = _schema(
        _sheet(
            "Sheet 1",
            {
                "A1": _cell("A1", "{{reports:repeat-start}}"),
                "A2": _cell("A2", "{{name:string}}"),
                "A3": _cell("A3", "{{reports:repeat-end}}"),
                "A5": _cell("A5", "{{reports:repeat-start}}"),
                "A6": _cell("A6", "{{name:string}}"),
                "A7": _cell("A7", "{{reports:repeat-end}}"),
            },
        )
    )

    with pytest.raises(ValueError, match="Duplicate repeat section key"):
        get_template_inputs(schema)


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


def test_get_template_inputs_raises_for_incompatible_dynamic_scope():
    schema = _schema(
        _sheet("reports", {"A1": _cell("A1", "{{name:string}}")}),
        _sheet("{{reports}}", {"A1": _cell("A1", "{{name:string}}")}),
    )

    with pytest.raises(ValueError, match="incompatible scopes"):
        get_template_inputs(schema)


def test_compile_rejects_non_dict_root_payload():
    schema = _schema(_sheet("Sheet 1", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(TypeError, match="Data must be an object/dict"):
        mo_dataport.compile(schema, ["not", "a", "dict"])


def test_compile_rejects_missing_sheet_payload():
    schema = _schema(_sheet("Sheet 1", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(KeyError, match="requires sheet 'Sheet 1'"):
        mo_dataport.compile(schema, {})


def test_compile_rejects_non_dict_sheet_payload():
    schema = _schema(_sheet("Sheet 1", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(TypeError, match="Data for sheet 'Sheet 1' must be an object"):
        mo_dataport.compile(schema, {"Sheet 1": "Alice"})


def test_compile_rejects_dynamic_sheet_payload_shape_and_names():
    schema = _schema(_sheet("{{reports}}", {"A1": _cell("A1", "{{name:string}}")}))

    with pytest.raises(TypeError, match="dynamic sheet group 'reports'"):
        mo_dataport.compile(schema, {"reports": []})

    with pytest.raises(TypeError, match="Dynamic sheet names under 'reports'"):
        mo_dataport.compile(schema, {"reports": {1: {"name": "Acme"}}})

    with pytest.raises(TypeError, match="Data for dynamic sheet 'North'"):
        mo_dataport.compile(schema, {"reports": {"North": "Acme"}})


def test_compile_rejects_duplicate_output_sheet_names():
    schema = _schema(
        _sheet("Summary", {"A1": _cell("A1", "Static")}),
        _sheet("{{reports}}", {"A1": _cell("A1", "{{name:string}}")}),
    )

    with pytest.raises(ValueError, match="Duplicate output sheet name 'Summary'"):
        mo_dataport.compile(schema, {"Summary": {}, "reports": {"Summary": {"name": "Acme"}}})


def test_compile_validates_scalar_and_dataframe_types():
    schema = _schema(
        _sheet(
            "Sheet 1",
            {
                "A1": _cell("A1", "{{active:boolean}}"),
                "A2": _cell("A2", "{{when:date}}"),
                "A3": _cell("A3", "{{rows:dataframe-content}}"),
            },
        )
    )

    with pytest.raises(TypeError, match="'active' expected type 'boolean'"):
        mo_dataport.compile(
            schema,
            {"Sheet 1": {"active": "yes", "when": "2024-01-01", "rows": []}},
        )

    with pytest.raises(TypeError, match="'when' expected type 'date'"):
        mo_dataport.compile(schema, {"Sheet 1": {"active": True, "when": 42, "rows": []}})

    with pytest.raises(TypeError, match="'rows' expected a polars DataFrame"):
        mo_dataport.compile(
            schema,
            {"Sheet 1": {"active": True, "when": "2024-01-01", "rows": []}},
        )


def test_compile_substitutes_scalar_placeholders_inside_text():
    schema = _schema(
        _sheet(
            "Sheet 1",
            {"A1": _cell("A1", "Customer {{name:string}} owes {{amount:number}}")},
        )
    )

    bundle = mo_dataport.compile(schema, {"Sheet 1": {"name": "Acme", "amount": 42}})

    assert bundle.report["sheets"][0]["cells"]["A1"]["value"] == "Customer Acme owes 42"


def test_compile_accepts_boolean_and_date_scalar_values():
    when = datetime.date(2024, 1, 2)
    schema = _schema(
        _sheet(
            "Sheet 1",
            {
                "A1": _cell("A1", "{{active:boolean}}"),
                "A2": _cell("A2", "{{when:date}}"),
            },
        )
    )

    bundle = mo_dataport.compile(schema, {"Sheet 1": {"active": True, "when": when}})

    cells = bundle.report["sheets"][0]["cells"]
    assert cells["A1"]["value"] is True
    assert cells["A1"]["cell_type"] == "string"
    assert cells["A2"]["value"] == "2024-01-02"
    assert cells["A2"]["cell_type"] == "date"


def test_infer_cell_type_number():
    assert _infer_cell_type(42) == "number"
    assert _infer_cell_type(3.14) == "number"


def test_infer_cell_type_string():
    assert _infer_cell_type("hello") == "string"


def test_infer_cell_type_date():
    assert _infer_cell_type(datetime.datetime(2024, 1, 1)) == "date"


def test_infer_cell_type_none():
    assert _infer_cell_type(None) == "empty"


# §5. Entrypoints
