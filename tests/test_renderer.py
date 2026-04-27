"""Tests for placeholder input discovery helpers."""

import datetime

import pytest

from mindoff_data_export.renderer import _infer_cell_type, get_template_inputs

# §1 Constants & Exceptions

# §2 Classes and Sub Classes

# §3 Private Helper Functions


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


# §4 Public Functions


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


def test_infer_cell_type_number():
    assert _infer_cell_type(42) == "number"
    assert _infer_cell_type(3.14) == "number"


def test_infer_cell_type_string():
    assert _infer_cell_type("hello") == "string"


def test_infer_cell_type_date():
    assert _infer_cell_type(datetime.datetime(2024, 1, 1)) == "date"


def test_infer_cell_type_none():
    assert _infer_cell_type(None) == "empty"


# §5 Entrypoints
