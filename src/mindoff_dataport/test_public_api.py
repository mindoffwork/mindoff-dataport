from mindoff_dataport import (
    compile_report_bundle,
    export_report_bundle,
    extract_template,
    get_template_inputs,
    mode,
    parquet_source,
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


def test_mode_namespace_exposes_bundle_first_aliases():
    assert mode.extract is extract_template
    assert mode.inputs is get_template_inputs
    assert mode.compile is compile_report_bundle
    assert mode.export is export_report_bundle
    assert mode.parquet_source is parquet_source
    assert not hasattr(mode, "build")
    assert not hasattr(mode, "alter_schema")


def test_mode_bundle_aliases_work(managed_tmp_dir):
    schema = _schema()

    inputs = mode.inputs(schema)
    assert inputs == {"Sheet1": {"name": "string"}}

    bundle = mode.compile(schema, {"Sheet1": {"name": "Alice"}})
    assert bundle.report["sheets"][0]["cells"]["A1"]["value"] == "Alice"

    output = managed_tmp_dir / "out.xlsx"
    mode.export(bundle, str(output))
    assert output.exists()


# §5 Entrypoints
