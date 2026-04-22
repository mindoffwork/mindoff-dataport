from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from mindoff_data_export import build_template_with_data

# §1 Types

# §2 Constants

# §3 Private Helpers


def _cell(coord: str, value):
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


def _schema(cells: dict[str, dict], *, dims: str = "A1:C3", merges: list[str] | None = None):
    return {
        "sheets": [
            {
                "name": "Sheet1",
                "dimensions": dims,
                "merged_regions": merges or [],
                "column_widths": {},
                "row_heights": {},
                "cells": cells,
            }
        ]
    }


polars = pytest.importorskip("polars", reason="polars not installed")


def _temp_output_path(managed_tmp_dir: Path, name: str) -> Path:
    return managed_tmp_dir / name


# §4 Public API


def test_streaming_split_outputs_and_names(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    df = polars.DataFrame({"A": [1, 2, 3, 4, 5]})
    out = _temp_output_path(managed_tmp_dir, "filled.xlsx")

    paths = build_template_with_data(
        schema,
        {"rows": df},
        str(out),
        export_mode="streaming",
        streaming_chunk_rows=2,
        max_rows_per_workbook=3,
    )

    assert isinstance(paths, list)
    assert [Path(p).name for p in paths] == ["filled.part001.xlsx", "filled.part002.xlsx"]

    wb1 = openpyxl.load_workbook(paths[0], data_only=True)
    wb2 = openpyxl.load_workbook(paths[1], data_only=True)
    ws1 = wb1["Sheet1"]
    ws2 = wb2["Sheet1"]
    assert [ws1["A1"].value, ws1["A2"].value, ws1["A3"].value] == [1, 2, 3]
    assert [ws2["A1"].value, ws2["A2"].value] == [4, 5]
    wb1.close()
    wb2.close()


def test_streaming_rejects_hug_mode(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    df = polars.DataFrame({"A": [1]})
    out = _temp_output_path(managed_tmp_dir, "filled.xlsx")

    with pytest.raises(ValueError, match="does not support 'hug'"):
        build_template_with_data(
            schema,
            {"rows": df},
            str(out),
            export_mode="streaming",
            column_width_mode="hug",
        )


def test_streaming_rejects_merged_regions(managed_tmp_dir: Path):
    schema = _schema(
        {"A1": _cell("A1", "{{rows:dataframe-content}}")},
        dims="A1:A1",
        merges=["A1:B1"],
    )
    df = polars.DataFrame({"A": [1]})
    out = _temp_output_path(managed_tmp_dir, "filled.xlsx")

    with pytest.raises(ValueError, match="does not support merged cells"):
        build_template_with_data(schema, {"rows": df}, str(out), export_mode="streaming")


def test_streaming_anchor_style_is_cloned(managed_tmp_dir: Path):
    anchor = _cell("A1", "{{rows:dataframe-content}}")
    anchor["fill"] = {"bg_color": "FFFF0000"}
    anchor["number_format"] = "0.00"
    schema = _schema({"A1": anchor}, dims="A1:A1")
    df = polars.DataFrame({"A": [1.5]})
    out = _temp_output_path(managed_tmp_dir, "filled.xlsx")

    paths = build_template_with_data(schema, {"rows": df}, str(out), export_mode="streaming")
    wb = openpyxl.load_workbook(paths[0])
    cell = wb["Sheet1"]["A1"]
    assert cell.fill.fgColor.type == "rgb"
    assert cell.fill.fgColor.rgb == "FFFF0000"
    assert cell.number_format == "0.00"
    wb.close()


def test_streaming_lazyframe_uses_collect_batches(monkeypatch, managed_tmp_dir: Path):
    lf = polars.DataFrame({"A": [1, 2]}).lazy()
    lazy_cls = type(lf)
    called = {"collect_batches": 0}
    original_batches = lazy_cls.collect_batches
    original_collect = lazy_cls.collect

    def _collect_batches(self, **kwargs):
        called["collect_batches"] += 1
        return original_batches(self, **kwargs)

    def _collect_fail(self, *args, **kwargs):
        raise AssertionError("collect() should not be called in streaming path")

    monkeypatch.setattr(lazy_cls, "collect_batches", _collect_batches)
    monkeypatch.setattr(lazy_cls, "collect", _collect_fail)

    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    out = _temp_output_path(managed_tmp_dir, "filled.xlsx")
    paths = build_template_with_data(
        schema,
        {"rows": lf},
        str(out),
        export_mode="streaming",
        streaming_chunk_rows=1,
        max_rows_per_workbook=10,
    )

    assert called["collect_batches"] >= 1
    wb = openpyxl.load_workbook(paths[0], data_only=True)
    assert [wb["Sheet1"]["A1"].value, wb["Sheet1"]["A2"].value] == [1, 2]
    wb.close()
