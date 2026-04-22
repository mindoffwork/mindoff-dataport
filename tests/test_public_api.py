from mindoff_data_export import (
    build_template_with_data,
    extract_template,
    get_template_inputs,
    mode,
    render_schema,
)

# §1 Constants & Exceptions

# §2 Classes and Sub Classes

# §3 Private Helper Functions


def _cell(coord: str, value: str) -> dict:
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


def _schema() -> dict:
    return {
        "sheets": [
            {
                "name": "Sheet1",
                "dimensions": "A1:A1",
                "merged_regions": [],
                "column_widths": {},
                "row_heights": {},
                "cells": {"A1": _cell("A1", "{{name:string}}")},
            }
        ]
    }


# §4 Public Functions


def test_mode_namespace_exposes_expected_aliases():
    assert mode.extract is extract_template
    assert mode.build is build_template_with_data
    assert mode.get_inputs is get_template_inputs
    assert mode.alter_schema is render_schema


def test_mode_renderer_aliases_work():
    schema = _schema()

    inputs = mode.get_inputs(schema)
    assert inputs == {"Sheet1": {"name": "string"}}

    rendered = mode.alter_schema(schema, {"Sheet1": {"name": "Alice"}})
    assert rendered["sheets"][0]["cells"]["A1"]["value"] == "Alice"


# §5 Entrypoints
