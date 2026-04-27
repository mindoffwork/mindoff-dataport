from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import openpyxl
import pyarrow as pa
import pyarrow.parquet as pq
import pytest

from mindoff_dataport import mode, parquet_source

# §1 Constants & Exceptions

polars = pytest.importorskip("polars", reason="polars not installed")

# §2 Classes and Sub Classes

# §3 Private Helper Functions


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


def _schema(
    cells: dict[str, dict], *, dims: str = "A1:C3", merges: list[str] | None = None
):
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


def _export_streaming(schema, data, managed_tmp_dir: Path, **options):
    bundle = mode.compile(schema, data)
    out = managed_tmp_dir / "filled.xlsx"
    return mode.export(bundle, str(out), export_mode="streaming", **options)


# §4 Public Functions


def test_streaming_split_outputs_and_names(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    df = polars.DataFrame({"A": [1, 2, 3, 4, 5]})

    paths = _export_streaming(
        schema,
        {"Sheet1": {"rows": df}},
        managed_tmp_dir,
        max_rows_per_workbook=3,
    )

    assert isinstance(paths, list)
    assert len(paths) == 1
    zip_path = Path(paths[0])
    assert zip_path.name == "filled.zip"
    assert zip_path.exists()

    with ZipFile(zip_path) as zip_file:
        assert zip_file.namelist() == ["filled.part001.xlsx", "filled.part002.xlsx"]
        zip_file.extractall(path=managed_tmp_dir)

    wb1 = openpyxl.load_workbook(managed_tmp_dir / "filled.part001.xlsx", data_only=True)
    wb2 = openpyxl.load_workbook(managed_tmp_dir / "filled.part002.xlsx", data_only=True)
    assert [wb1["Sheet1"]["A1"].value, wb1["Sheet1"]["A2"].value, wb1["Sheet1"]["A3"].value] == [1, 2, 3]
    assert [wb2["Sheet1"]["A1"].value, wb2["Sheet1"]["A2"].value] == [4, 5]
    wb1.close()
    wb2.close()


def test_streaming_single_part_does_not_zip(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    df = polars.DataFrame({"A": [1, 2]})

    paths = _export_streaming(
        schema,
        {"Sheet1": {"rows": df}},
        managed_tmp_dir,
        max_rows_per_workbook=10,
    )

    assert len(paths) == 1
    assert Path(paths[0]).name == "filled.part001.xlsx"
    assert Path(paths[0]).exists()
    assert not (managed_tmp_dir / "filled.zip").exists()


def test_streaming_reads_parquet_source_in_batches(managed_tmp_dir: Path):
    source_path = managed_tmp_dir / "source.parquet"
    pq.write_table(pa.table({"A": [1, 2, 3], "B": [4, 5, 6]}), source_path)
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")

    paths = _export_streaming(
        schema,
        {"Sheet1": {"rows": parquet_source(str(source_path), columns=["B", "A"])}},
        managed_tmp_dir,
        streaming_chunk_rows=1,
        max_rows_per_workbook=10,
    )

    wb = openpyxl.load_workbook(paths[0], data_only=True)
    assert [wb["Sheet1"]["A1"].value, wb["Sheet1"]["B1"].value] == [4, 1]
    assert [wb["Sheet1"]["A3"].value, wb["Sheet1"]["B3"].value] == [6, 3]
    wb.close()


def test_streaming_rejects_hug_mode(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    df = polars.DataFrame({"A": [1]})

    with pytest.raises(ValueError, match="does not support 'hug'"):
        _export_streaming(
            schema,
            {"Sheet1": {"rows": df}},
            managed_tmp_dir,
            column_width_mode="hug",
        )


def test_streaming_preserves_static_merged_regions(managed_tmp_dir: Path):
    title = _cell("A1", "Merged Title")
    title["merged"] = True
    title["merge_anchor"] = "A1"
    title["fill"] = {"bg_color": "FF00FF00"}
    shadow = _cell("B1", "Should not be written")
    shadow["merged"] = True
    shadow["merge_anchor"] = "A1"
    schema = _schema(
        {
            "A1": title,
            "B1": shadow,
            "A3": _cell("A3", "{{rows:dataframe-content}}"),
        },
        dims="A1:B3",
        merges=["A1:B1"],
    )
    df = polars.DataFrame({"A": [10], "B": [20]})

    paths = _export_streaming(schema, {"Sheet1": {"rows": df}}, managed_tmp_dir)

    wb = openpyxl.load_workbook(paths[0], data_only=True)
    ws = wb["Sheet1"]
    assert [str(region) for region in ws.merged_cells.ranges] == ["A1:B1"]
    assert ws["A1"].value == "Merged Title"
    assert ws["B1"].value is None
    assert ws["A1"].fill.fgColor.rgb == "FF00FF00"
    assert [ws["A3"].value, ws["B3"].value] == [10, 20]
    wb.close()


def test_streaming_rejects_merged_regions_that_intersect_content(managed_tmp_dir: Path):
    schema = _schema(
        {"A1": _cell("A1", "{{rows:dataframe-content}}")},
        dims="A1:A1",
        merges=["A1:B1"],
    )
    df = polars.DataFrame({"A": [1]})

    with pytest.raises(ValueError, match="intersect dataframe-content"):
        _export_streaming(schema, {"Sheet1": {"rows": df}}, managed_tmp_dir)


def test_streaming_anchor_style_is_cloned(managed_tmp_dir: Path):
    anchor = _cell("A1", "{{rows:dataframe-content}}")
    anchor["fill"] = {"bg_color": "FFFF0000"}
    anchor["number_format"] = "0.00"
    schema = _schema({"A1": anchor}, dims="A1:A1")
    df = polars.DataFrame({"A": [1.5]})

    paths = _export_streaming(schema, {"Sheet1": {"rows": df}}, managed_tmp_dir)
    wb = openpyxl.load_workbook(paths[0])
    cell = wb["Sheet1"]["A1"]
    assert cell.fill.fgColor.type == "rgb"
    assert cell.fill.fgColor.rgb == "FFFF0000"
    assert cell.number_format == "0.00"
    wb.close()


def test_streaming_expands_dynamic_sheet_names_in_order(managed_tmp_dir: Path):
    schema = {
        "sheets": [
            {
                "name": "{{sheet_1}}",
                "dimensions": "A1:A1",
                "merged_regions": [],
                "column_widths": {},
                "row_heights": {},
                "cells": {"A1": _cell("A1", "{{name:string}}")},
            }
        ]
    }

    paths = _export_streaming(
        schema,
        {
            "sheet_1": {
                "Sheet Name 1": {"name": "Alpha"},
                "Sheet Name 2": {"name": "Beta"},
            }
        },
        managed_tmp_dir,
        max_rows_per_workbook=10,
    )

    wb = openpyxl.load_workbook(paths[0], data_only=True)
    assert wb.sheetnames == ["Sheet Name 1", "Sheet Name 2"]
    assert wb["Sheet Name 1"]["A1"].value == "Alpha"
    assert wb["Sheet Name 2"]["A1"].value == "Beta"
    wb.close()


# §5 Entrypoints
