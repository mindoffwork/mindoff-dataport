from __future__ import annotations

import datetime
import json
from pathlib import Path

import openpyxl
import pytest
import reportlab

from mindoff_dataport import mode
from mindoff_dataport.bundle import load_report_bundle
from mindoff_dataport.pdf_renderer import _FontResolver, _LazyFlowables, _table_style

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


def _schema(cells: dict[str, dict], *, dims: str = "A1:C3") -> dict:
    return {
        "sheets": [
            {
                "name": "Sheet1",
                "dimensions": dims,
                "merged_regions": [],
                "column_widths": {},
                "row_heights": {},
                "cells": cells,
            }
        ]
    }


def _schema_with_options(cells: dict[str, dict], *, dims: str = "A1:C3", **options):
    schema = _schema(cells, dims=dims)
    schema["sheets"][0].update(options)
    return schema


def _has_line_command(commands: list[tuple], command: str, width: float) -> bool:
    return any(
        item[:4] == (command, (0, 0), (1, 1), width)
        for item in commands
    )


def _vera_font_path() -> str:
    return str(Path(reportlab.__file__).parent / "fonts" / "Vera.ttf")


# §4 Public Functions


def test_compile_creates_valid_report_bundle_directory(managed_tmp_dir: Path):
    schema = _schema(
        {
            "A1": _cell("A1", "{{name:string}}"),
            "A2": _cell("A2", "{{rows:dataframe-content}}"),
        },
        dims="A1:B2",
    )
    df = polars.DataFrame({"A": [1, 2], "B": [3, 4]})
    bundle_path = managed_tmp_dir / "report_bundle"

    bundle = mode.compile(
        schema,
        {"Sheet1": {"name": "Alice", "rows": df}},
        bundle_path=str(bundle_path),
    )

    assert bundle_path.exists()
    assert bundle_path.is_dir()
    assert (bundle_path / "manifest.json").exists()
    assert (bundle_path / "report.json").exists()
    assert bundle.report["sheets"][0]["cells"]["A1"]["value"] == "Alice"
    assert "A2" not in bundle.report["sheets"][0]["cells"]
    source_path = bundle.report["sheets"][0]["dataframe_anchors"][0]["source"]
    assert source_path.startswith("data/")
    assert (bundle_path / source_path).exists()

    manifest = json.loads((bundle_path / "manifest.json").read_text(encoding="utf-8"))
    report = json.loads((bundle_path / "report.json").read_text(encoding="utf-8"))

    assert manifest["bundle_format"] == "directory"
    assert manifest["output_capabilities"] == {
        "xlsx": True,
        "pdf": True,
        "image": False,
    }
    assert report["sheets"][0]["dataframe_anchors"][0]["columns"] == ["A", "B"]


def test_compile_creates_compact_repeat_section_bundle(managed_tmp_dir: Path):
    schema = _schema(
        {
            "A1": _cell("A1", "{{reports:repeat-start}}"),
            "A2": _cell("A2", "{{customer_name:string}}"),
            "A3": _cell("A3", "{{line_items:dataframe}}"),
            "A4": _cell("A4", "{{reports:repeat-end}}"),
        },
        dims="A1:B4",
    )
    rows = polars.DataFrame({"sku": ["A", "B"], "qty": [1, 2]})

    bundle = mode.compile(
        schema,
        {
            "Sheet1": {
                "reports": [
                    {"customer_name": "Acme", "line_items": rows},
                    {"customer_name": "Globex", "line_items": rows},
                ]
            }
        },
        bundle_path=str(managed_tmp_dir / "bundle"),
    )

    sheet = bundle.report["sheets"][0]
    assert sheet["cells"] == {}
    assert "repeat_sections" in sheet
    section = sheet["repeat_sections"][0]
    assert section["key"] == "reports"
    assert len(section["records"]) == 2
    assert section["records"][0]["cells"][0]["cell"]["value"] == "Acme"
    assert len(bundle.manifest["dataframe_sources"]) == 1


def test_compile_rejects_repeat_payload_that_is_not_list():
    schema = _schema(
        {
            "A1": _cell("A1", "{{reports:repeat-start}}"),
            "A2": _cell("A2", "{{customer_name:string}}"),
            "A3": _cell("A3", "{{reports:repeat-end}}"),
        },
        dims="A1:A3",
    )

    with pytest.raises(TypeError, match="must be a list"):
        mode.compile(schema, {"Sheet1": {"reports": {"customer_name": "Acme"}}})


def test_compile_rejects_repeat_item_that_is_not_dict():
    schema = _schema(
        {
            "A1": _cell("A1", "{{reports:repeat-start}}"),
            "A2": _cell("A2", "{{customer_name:string}}"),
            "A3": _cell("A3", "{{reports:repeat-end}}"),
        },
        dims="A1:A3",
    )

    with pytest.raises(TypeError, match="item 0.*must be an object/dict"):
        mode.compile(schema, {"Sheet1": {"reports": ["Acme"]}})


def test_compile_rejects_repeat_item_missing_required_value():
    schema = _schema(
        {
            "A1": _cell("A1", "{{reports:repeat-start}}"),
            "A2": _cell("A2", "{{customer_name:string}}"),
            "A3": _cell("A3", "{{reports:repeat-end}}"),
        },
        dims="A1:A3",
    )

    with pytest.raises(KeyError, match="Sheet 'Sheet1.reports\\[0\\]' requires"):
        mode.compile(schema, {"Sheet1": {"reports": [{}]}})


def test_compile_preserves_static_content_after_repeat_end():
    schema = _schema(
        {
            "A1": _cell("A1", "{{reports:repeat-start}}"),
            "A2": _cell("A2", "{{customer_name:string}}"),
            "A3": _cell("A3", "{{reports:repeat-end}}"),
            "A4": _cell("A4", "Footer"),
        },
        dims="A1:A4",
    )

    bundle = mode.compile(schema, {"Sheet1": {"reports": [{"customer_name": "Acme"}]}})

    assert bundle.report["sheets"][0]["cells"]["A4"]["value"] == "Footer"


def test_compile_creates_ordered_sibling_repeat_sections():
    schema = _schema(
        {
            "A1": _cell("A1", "Title"),
            "A2": _cell("A2", "{{customers:repeat-start}}"),
            "A3": _cell("A3", "{{customer_name:string}}"),
            "A4": _cell("A4", "{{customers:repeat-end}}"),
            "A5": _cell("A5", "Between"),
            "A6": _cell("A6", "{{summaries:repeat-start}}"),
            "A7": _cell("A7", "{{region:string}}"),
            "A8": _cell("A8", "{{totals:dataframe-content}}"),
            "A9": _cell("A9", "{{summaries:repeat-end}}"),
            "A10": _cell("A10", "Footer"),
        },
        dims="A1:B10",
    )

    bundle = mode.compile(
        schema,
        {
            "Sheet1": {
                "customers": [{"customer_name": "Acme"}],
                "summaries": [
                    {
                        "region": "North",
                        "totals": polars.DataFrame({"metric": ["Sales"], "value": [10]}),
                    }
                ],
            }
        },
    )

    sheet = bundle.report["sheets"][0]
    assert [section["key"] for section in sheet["repeat_sections"]] == [
        "customers",
        "summaries",
    ]
    assert sheet["cells"]["A1"]["value"] == "Title"
    assert sheet["cells"]["A5"]["value"] == "Between"
    assert sheet["cells"]["A10"]["value"] == "Footer"


def test_compile_stores_repeat_merges_as_relative_metadata():
    title = _cell("A2", "{{customer_name:string}}")
    title["merged"] = True
    title["merge_anchor"] = "A2"
    shadow = _cell("B2", None)
    shadow["merged"] = True
    shadow["merge_anchor"] = "A2"
    schema = _schema(
        {
            "A1": _cell("A1", "{{reports:repeat-start}}"),
            "A2": title,
            "B2": shadow,
            "A3": _cell("A3", "{{reports:repeat-end}}"),
        },
        dims="A1:B3",
    )
    schema["sheets"][0]["merged_regions"] = ["A2:B2"]

    bundle = mode.compile(schema, {"Sheet1": {"reports": [{"customer_name": "Acme"}]}})

    section = bundle.report["sheets"][0]["repeat_sections"][0]
    assert bundle.report["sheets"][0]["merged_regions"] == []
    assert section["merged_regions"] == [
        {"min_row_offset": 0, "max_row_offset": 0, "min_col": 1, "max_col": 2}
    ]


def test_compile_rejects_repeat_merge_over_dataframe_content():
    anchor = _cell("A2", "{{line_items:dataframe-content}}")
    anchor["merged"] = True
    anchor["merge_anchor"] = "A2"
    shadow = _cell("B2", None)
    shadow["merged"] = True
    shadow["merge_anchor"] = "A2"
    schema = _schema(
        {
            "A1": _cell("A1", "{{reports:repeat-start}}"),
            "A2": anchor,
            "B2": shadow,
            "A3": _cell("A3", "{{reports:repeat-end}}"),
        },
        dims="A1:B3",
    )
    schema["sheets"][0]["merged_regions"] = ["A2:B2"]

    with pytest.raises(ValueError, match="dataframe-content rows"):
        mode.compile(
            schema,
            {
                "Sheet1": {
                    "reports": [
                        {"line_items": polars.DataFrame({"A": [1], "B": [2]})}
                    ]
                }
            },
        )


def test_compile_without_bundle_path_creates_temp_directory():
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    df = polars.DataFrame({"A": [1]})

    bundle = mode.compile(schema, {"Sheet1": {"rows": df}})

    path = Path(bundle.path)
    assert path.exists()
    assert path.is_dir()
    assert (path / "manifest.json").exists()
    paths = [source["path"] for source in bundle.manifest["dataframe_sources"]]
    assert paths == ["data/Sheet1__rows__dataframe-content.parquet"]
    assert (path / paths[0]).exists()


def test_compile_accepts_lazyframe_without_collecting_rows(managed_tmp_dir: Path):
    source_path = managed_tmp_dir / "source.parquet"
    polars.DataFrame({"A": [1, 2], "B": [3, 4]}).write_parquet(source_path)
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    bundle_path = managed_tmp_dir / "bundle"

    bundle = mode.compile(
        schema,
        {"Sheet1": {"rows": polars.scan_parquet(source_path).select(["B", "A"])}},
        bundle_path=str(bundle_path),
    )

    source = bundle.manifest["dataframe_sources"][0]
    assert source["columns"] == ["B", "A"]
    assert source["rows"] == 2
    assert source["format"] == "parquet"
    assert (bundle_path / source["path"]).exists()


def test_report_bundle_write_copies_directory(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle = mode.compile(schema, {"Sheet1": {"name": "Alice"}})
    target = managed_tmp_dir / "copied_bundle"

    bundle.write(str(target))
    copied = load_report_bundle(str(target))

    assert copied.report["sheets"][0]["cells"]["A1"]["value"] == "Alice"
    assert copied.manifest["bundle_format"] == "directory"


def test_compile_rejects_bundle_path_that_is_file(managed_tmp_dir: Path):
    target = managed_tmp_dir / "bundle-file"
    target.write_text("not a directory", encoding="utf-8")
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")

    with pytest.raises(ValueError, match="must be a directory"):
        mode.compile(schema, {"Sheet1": {"name": "Alice"}}, bundle_path=str(target))


def test_report_bundle_write_rejects_file_target(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle = mode.compile(schema, {"Sheet1": {"name": "Alice"}})
    target = managed_tmp_dir / "bundle-file"
    target.write_text("not a directory", encoding="utf-8")

    with pytest.raises(ValueError, match="must be a directory"):
        bundle.write(str(target))


def test_load_rejects_missing_file_path_and_unsupported_metadata(managed_tmp_dir: Path):
    missing = managed_tmp_dir / "missing"
    file_path = managed_tmp_dir / "bundle-file"
    file_path.write_text("not a directory", encoding="utf-8")

    with pytest.raises(FileNotFoundError, match="Report bundle not found"):
        load_report_bundle(str(missing))
    with pytest.raises(ValueError, match="must be a directory"):
        load_report_bundle(str(file_path))

    bad_version = managed_tmp_dir / "bad-version"
    bad_version.mkdir()
    (bad_version / "manifest.json").write_text(
        json.dumps({"version": "0.0", "bundle_format": "directory"}),
        encoding="utf-8",
    )
    (bad_version / "report.json").write_text("{}", encoding="utf-8")
    with pytest.raises(ValueError, match="Unsupported report bundle version"):
        load_report_bundle(str(bad_version))

    bad_format = managed_tmp_dir / "bad-format"
    bad_format.mkdir()
    (bad_format / "manifest.json").write_text(
        json.dumps({"version": "1.0", "bundle_format": "zip"}),
        encoding="utf-8",
    )
    (bad_format / "report.json").write_text("{}", encoding="utf-8")
    with pytest.raises(ValueError, match="Unsupported report bundle format"):
        load_report_bundle(str(bad_format))


def test_compile_rejects_non_polars_dataframe_source():
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")

    with pytest.raises(TypeError, match="expected a polars DataFrame or LazyFrame"):
        mode.compile(schema, {"Sheet1": {"rows": [{"A": 1}]}})


def test_compile_sanitizes_and_deduplicates_dataframe_source_ids(managed_tmp_dir: Path):
    schema = {
        "sheets": [
            {
                "name": "{{reports}}",
                "dimensions": "A1:A1",
                "merged_regions": [],
                "column_widths": {},
                "row_heights": {},
                "cells": {"A1": _cell("A1", "{{rows:dataframe-content}}")},
            }
        ]
    }
    df = polars.DataFrame({"A": [1]})

    bundle = mode.compile(
        schema,
        {"reports": {"North/West": {"rows": df}, "North?West": {"rows": df.clone()}}},
        bundle_path=str(managed_tmp_dir / "bundle"),
    )

    paths = [source["path"] for source in bundle.manifest["dataframe_sources"]]
    assert paths == [
        "data/North_West__rows__dataframe-content.parquet",
        "data/North_West__rows__dataframe-content_2.parquet",
    ]


def test_compile_serializes_date_scalar_in_report():
    schema = _schema({"A1": _cell("A1", "{{when:date}}")}, dims="A1:A1")

    bundle = mode.compile(
        schema,
        {"Sheet1": {"when": datetime.datetime(2024, 1, 2, 3, 4, 5)}},
    )

    assert bundle.report["sheets"][0]["cells"]["A1"]["value"] == "2024-01-02T03:04:05"


def test_xlsx_export_uses_bundle_data_without_expanding_report(managed_tmp_dir: Path):
    schema = _schema(
        {
            "A1": _cell("A1", "{{name:string}}"),
            "A2": _cell("A2", "{{headers:dataframe-header}}"),
            "A3": _cell("A3", "{{rows:dataframe-content}}"),
        },
        dims="A1:B3",
    )
    df = polars.DataFrame({"A": [1, 2], "B": [3, 4]})
    bundle = mode.compile(
        schema,
        {"Sheet1": {"name": "Alice", "headers": df, "rows": df}},
    )

    assert "A3" not in bundle.report["sheets"][0]["cells"]
    out = managed_tmp_dir / "out.xlsx"
    paths = mode.export(bundle, str(out), export_mode="streaming")

    wb = openpyxl.load_workbook(paths[0], data_only=True)
    ws = wb["Sheet1"]
    assert ws["A1"].value == "Alice"
    assert [ws["A2"].value, ws["B2"].value] == ["A", "B"]
    assert [ws["A3"].value, ws["B3"].value, ws["A4"].value, ws["B4"].value] == [
        1,
        3,
        2,
        4,
    ]
    wb.close()


def test_dataframe_placeholder_writes_headers_and_content(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe}}")}, dims="A1:B1")
    df = polars.DataFrame({"A": [1, 2], "B": [3, 4]}).lazy()
    bundle = mode.compile(schema, {"Sheet1": {"rows": df}})

    anchors = bundle.report["sheets"][0]["dataframe_anchors"]
    assert [anchor["placeholder_type"] for anchor in anchors] == [
        "dataframe-header",
        "dataframe-content",
    ]
    assert anchors[0]["start_row"] == 1
    assert anchors[1]["start_row"] == 2

    out = managed_tmp_dir / "combined.xlsx"
    paths = mode.export(bundle, str(out), export_mode="streaming")

    wb = openpyxl.load_workbook(paths[0], data_only=True)
    ws = wb["Sheet1"]
    assert [ws["A1"].value, ws["B1"].value] == ["A", "B"]
    assert [ws["A2"].value, ws["B2"].value, ws["A3"].value, ws["B3"].value] == [
        1,
        3,
        2,
        4,
    ]
    wb.close()


def test_dataframe_header_only_does_not_write_source_file(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-header}}")}, dims="A1:B1")
    df = polars.DataFrame({"A": [1], "B": [2]}).lazy()

    bundle = mode.compile(schema, {"Sheet1": {"rows": df}})

    assert bundle.manifest["dataframe_sources"] == []
    anchors = bundle.report["sheets"][0]["dataframe_anchors"]
    assert len(anchors) == 1
    assert anchors[0]["placeholder_type"] == "dataframe-header"
    assert anchors[0]["source"] is None


def test_old_dataframe_headers_spelling_is_not_a_placeholder():
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-headers}}")}, dims="A1:A1")

    assert mode.inputs(schema) == {"Sheet1": {}}


def test_compile_deduplicates_same_lazyframe_object(managed_tmp_dir: Path):
    source_path = managed_tmp_dir / "source.parquet"
    polars.DataFrame({"A": [1], "B": [2]}).write_parquet(source_path)
    rows = polars.scan_parquet(source_path)
    schema = _schema(
        {
            "A1": _cell("A1", "{{headers:dataframe-header}}"),
            "A2": _cell("A2", "{{rows:dataframe-content}}"),
        },
        dims="A1:B2",
    )

    bundle = mode.compile(schema, {"Sheet1": {"headers": rows, "rows": rows}})

    assert len(bundle.manifest["dataframe_sources"]) == 1
    content_anchor = [
        anchor
        for anchor in bundle.report["sheets"][0]["dataframe_anchors"]
        if anchor["placeholder_type"] == "dataframe-content"
    ][0]
    assert content_anchor["source"] == bundle.manifest["dataframe_sources"][0]["path"]


def test_compile_keeps_distinct_lazyframe_objects_separate(managed_tmp_dir: Path):
    source_path = managed_tmp_dir / "source.parquet"
    polars.DataFrame({"A": [1], "B": [2]}).write_parquet(source_path)
    schema = _schema(
        {
            "A1": _cell("A1", "{{first:dataframe-content}}"),
            "A3": _cell("A3", "{{second:dataframe-content}}"),
        },
        dims="A1:B3",
    )

    bundle = mode.compile(
        schema,
        {
            "Sheet1": {
                "first": polars.scan_parquet(source_path),
                "second": polars.scan_parquet(source_path),
            }
        },
    )

    assert len(bundle.manifest["dataframe_sources"]) == 2


def test_export_accepts_bundle_path(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle_path = managed_tmp_dir / "report_bundle"
    mode.compile(schema, {"Sheet1": {"name": "Alice"}}, bundle_path=str(bundle_path))
    out = managed_tmp_dir / "out.xlsx"

    mode.export(str(bundle_path), str(out))

    wb = openpyxl.load_workbook(out, data_only=True)
    assert wb["Sheet1"]["A1"].value == "Alice"
    wb.close()


def test_xlsx_export_preserves_gridline_visibility(managed_tmp_dir: Path):
    schema = _schema_with_options(
        {"A1": _cell("A1", "No gridlines")},
        dims="A1:A1",
        show_gridlines=False,
    )
    bundle = mode.compile(schema, {"Sheet1": {}})
    out = managed_tmp_dir / "no-gridlines.xlsx"

    mode.export(bundle, str(out))

    wb = openpyxl.load_workbook(out)
    assert wb["Sheet1"].sheet_view.showGridLines is False
    wb.close()


def test_xlsx_export_preserves_merged_region_border(managed_tmp_dir: Path):
    title = _cell("A1", "Merged Title")
    title["merged"] = True
    title["merge_anchor"] = "A1"
    title["borders"] = {
        "top": {"style": "thin", "color": "FF000000"},
        "bottom": {"style": "thick", "color": "FF000000"},
        "left": {"style": "medium", "color": "FF000000"},
        "right": {"style": "dashed", "color": "FF000000"},
    }
    shadow = _cell("B2", "")
    shadow["merged"] = True
    shadow["merge_anchor"] = "A1"
    schema = _schema({"A1": title, "B2": shadow}, dims="A1:B2")
    schema["sheets"][0]["merged_regions"] = ["A1:B2"]
    bundle = mode.compile(schema, {"Sheet1": {}})
    out = managed_tmp_dir / "merged-border.xlsx"

    mode.export(bundle, str(out))

    wb = openpyxl.load_workbook(out)
    ws = wb["Sheet1"]
    assert ws["A1"].border.top.style == "thin"
    assert ws["A2"].border.bottom.style == "thick"
    assert ws["A1"].border.left.style == "medium"
    assert ws["B1"].border.right.style == "dashed"
    wb.close()


def test_export_pdf_creates_nonempty_pdf(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle = mode.compile(schema, {"Sheet1": {"name": "Alice"}})
    out = managed_tmp_dir / "out.pdf"

    mode.export(bundle, str(out), format="pdf")

    assert out.exists()
    assert out.stat().st_size > 0
    assert out.read_bytes().startswith(b"%PDF")


def test_export_image_is_reserved(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle = mode.compile(schema, {"Sheet1": {"name": "Alice"}})

    with pytest.raises(NotImplementedError, match="image"):
        mode.export(bundle, str(managed_tmp_dir / "out.png"), format="image")


def test_pdf_export_renders_parquet_backed_dataframe(managed_tmp_dir: Path):
    source_path = managed_tmp_dir / "source.parquet"
    polars.DataFrame({"A": [1, 2], "B": [3, 4]}).write_parquet(source_path)
    schema = _schema(
        {
            "A1": _cell("A1", "{{headers:dataframe-header}}"),
            "A2": _cell("A2", "{{rows:dataframe-content}}"),
        },
        dims="A1:B2",
    )
    bundle = mode.compile(
        schema,
        {
            "Sheet1": {
                "headers": polars.scan_parquet(source_path).select(["B", "A"]),
                "rows": polars.scan_parquet(source_path).select(["B", "A"]),
            }
        },
    )
    out = managed_tmp_dir / "table.pdf"

    mode.export(bundle, str(out), format="pdf", streaming_chunk_rows=1)

    assert out.exists()
    assert out.read_bytes().startswith(b"%PDF")


def test_pdf_export_handles_merges_and_basic_styles(managed_tmp_dir: Path):
    title = _cell("A1", "Merged Title")
    title["merged"] = True
    title["merge_anchor"] = "A1"
    title["font"] = dict(title["font"])
    title["font"]["bold"] = True
    title["fill"] = {"bg_color": "FF003366"}
    title["alignment"] = dict(title["alignment"])
    title["alignment"]["horizontal"] = "center"
    shadow = _cell("B1", "")
    shadow["merged"] = True
    shadow["merge_anchor"] = "A1"
    schema = _schema({"A1": title, "B1": shadow}, dims="A1:B1")
    schema["sheets"][0]["merged_regions"] = ["A1:B1"]
    bundle = mode.compile(schema, {"Sheet1": {}})
    out = managed_tmp_dir / "styled.pdf"

    mode.export(bundle, str(out), format="pdf")

    assert out.exists()
    assert out.stat().st_size > 0


def test_pdf_table_style_applies_merged_region_border_to_span():
    title = _cell("A1", "Merged Title")
    title["merged"] = True
    title["merge_anchor"] = "A1"
    title["borders"] = {
        "top": {"style": "thin", "color": "FF000000"},
        "bottom": {"style": "thick", "color": "FF000000"},
        "left": {"style": "medium", "color": "FF000000"},
        "right": {"style": "dashed", "color": "FF000000"},
    }
    sheet = _schema({"A1": title}, dims="A1:B2")["sheets"][0]
    sheet["merged_regions"] = ["A1:B2"]

    commands = _table_style(
        sheet, {(1, 1): title}, 1, 1, 2, 2, _FontResolver()
    ).getCommands()

    assert _has_line_command(commands, "LINEABOVE", 0.5)
    assert _has_line_command(commands, "LINEBELOW", 1.5)
    assert _has_line_command(commands, "LINEBEFORE", 1.0)
    assert _has_line_command(commands, "LINEAFTER", 0.75)


def test_pdf_table_style_applies_merged_region_fill_to_span():
    title = _cell("A1", "Merged Title")
    title["merged"] = True
    title["merge_anchor"] = "A1"
    title["fill"] = {"bg_color": "FF003366"}
    sheet = _schema({"A1": title}, dims="A1:D1")["sheets"][0]
    sheet["merged_regions"] = ["A1:D1"]

    commands = _table_style(
        sheet, {(1, 1): title}, 1, 1, 4, 1, _FontResolver()
    ).getCommands()

    assert any(item[0] == "SPAN" and item[1] == (0, 0) and item[2] == (3, 0) for item in commands)
    assert any(item[0] == "BACKGROUND" and item[1] == (0, 0) and item[2] == (3, 0) for item in commands)


def test_pdf_table_style_does_not_draw_default_grid_for_empty_cells():
    cell = _cell("A1", "Visible")
    sheet = _schema({"A1": cell}, dims="A1:B2")["sheets"][0]

    commands = _table_style(
        sheet, {(1, 1): cell}, 1, 1, 2, 2, _FontResolver()
    ).getCommands()

    assert not any(item[0] == "GRID" for item in commands)
    assert not any(
        item[0] in {"LINEABOVE", "LINEBELOW", "LINEBEFORE", "LINEAFTER"}
        for item in commands
    )


def test_pdf_table_style_uses_registered_custom_font():
    cell = _cell("A1", "Custom Font")
    cell["font"] = dict(cell["font"])
    cell["font"]["name"] = "Vera"
    sheet = _schema({"A1": cell}, dims="A1:A1")["sheets"][0]
    resolver = _FontResolver({"Vera": _vera_font_path()})

    commands = _table_style(sheet, {(1, 1): cell}, 1, 1, 1, 1, resolver).getCommands()

    font_commands = [item for item in commands if item[0] == "FONTNAME"]
    assert font_commands
    assert str(font_commands[0][3]).startswith("Mindoff-Vera-regular-")


def test_pdf_lazy_flowables_accepts_split_insertion():
    flowables = _LazyFlowables(["first", "second"])

    flowables[0:0] = ["split"]
    flowables.insert(0, "retry")

    assert [flowables[0], flowables[1], flowables[2], flowables[3]] == [
        "retry",
        "split",
        "first",
        "second",
    ]


def test_pdf_export_accepts_custom_fonts(managed_tmp_dir: Path):
    cell = _cell("A1", "Custom Font")
    cell["font"] = dict(cell["font"])
    cell["font"]["name"] = "Vera"
    schema = _schema({"A1": cell}, dims="A1:A1")
    bundle = mode.compile(schema, {"Sheet1": {}})
    out = managed_tmp_dir / "custom-font.pdf"

    mode.export(
        bundle,
        str(out),
        format="pdf",
        fonts={"Vera": {"regular": _vera_font_path()}},
    )

    assert out.exists()
    assert out.read_bytes().startswith(b"%PDF")


def test_pdf_export_accepts_bundle_path(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle_path = managed_tmp_dir / "report_bundle"
    mode.compile(schema, {"Sheet1": {"name": "Alice"}}, bundle_path=str(bundle_path))
    out = managed_tmp_dir / "out.pdf"

    mode.export(str(bundle_path), str(out), format="pdf")

    assert out.exists()
    assert out.read_bytes().startswith(b"%PDF")


def test_pdf_export_auto_delete_removes_bundle_after_success(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle_path = managed_tmp_dir / "bundle"
    bundle = mode.compile(
        schema,
        {"Sheet1": {"name": "Alice"}},
        bundle_path=str(bundle_path),
    )
    out = managed_tmp_dir / "out.pdf"

    mode.export(bundle, str(out), format="pdf", auto_delete_bundle=True)

    assert out.exists()
    assert not bundle_path.exists()


def test_pdf_export_rejects_streaming_export_mode(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle = mode.compile(schema, {"Sheet1": {"name": "Alice"}})

    with pytest.raises(ValueError, match="PDF export does not support export_mode"):
        mode.export(
            bundle,
            str(managed_tmp_dir / "out.pdf"),
            format="pdf",
            export_mode="streaming",
        )


def test_pdf_export_rejects_invalid_page_options(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle = mode.compile(schema, {"Sheet1": {"name": "Alice"}})

    with pytest.raises(ValueError, match="Unsupported PDF page_size"):
        mode.export(
            bundle,
            str(managed_tmp_dir / "bad-page.pdf"),
            format="pdf",
            page_size="tabloid",
        )
    with pytest.raises(ValueError, match="Unsupported PDF orientation"):
        mode.export(
            bundle,
            str(managed_tmp_dir / "bad-orientation.pdf"),
            format="pdf",
            orientation="sideways",
        )
    with pytest.raises(ValueError, match="margin must be non-negative"):
        mode.export(
            bundle,
            str(managed_tmp_dir / "bad-margin.pdf"),
            format="pdf",
            margin=-1,
        )


def test_pdf_export_rejects_invalid_font_configs(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "Custom Font")}, dims="A1:A1")
    bundle = mode.compile(schema, {"Sheet1": {}})

    with pytest.raises(TypeError, match="PDF font config must be a path or dict"):
        mode.export(
            bundle,
            str(managed_tmp_dir / "bad-font-type.pdf"),
            format="pdf",
            fonts={"Broken": 42},
        )
    with pytest.raises(ValueError, match="file does not exist"):
        mode.export(
            bundle,
            str(managed_tmp_dir / "missing-font.pdf"),
            format="pdf",
            fonts={"Missing": str(managed_tmp_dir / "missing.ttf")},
        )
    with pytest.raises(ValueError, match="requires a regular font file"):
        mode.export(
            bundle,
            str(managed_tmp_dir / "missing-regular.pdf"),
            format="pdf",
            fonts={"Broken": {"bold": _vera_font_path()}},
        )


def test_load_rejects_malformed_bundle(managed_tmp_dir: Path):
    bad_path = managed_tmp_dir / "bad"
    bad_path.mkdir()
    (bad_path / "manifest.json").write_text("{}", encoding="utf-8")

    with pytest.raises(ValueError, match="Malformed report bundle"):
        load_report_bundle(str(bad_path))


def test_export_auto_delete_removes_bundle_after_success(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{name:string}}")}, dims="A1:A1")
    bundle_path = managed_tmp_dir / "bundle"
    bundle = mode.compile(
        schema,
        {"Sheet1": {"name": "Alice"}},
        bundle_path=str(bundle_path),
    )
    out = managed_tmp_dir / "out.xlsx"

    mode.export(bundle, str(out), auto_delete_bundle=True)

    assert out.exists()
    assert not bundle_path.exists()


def test_failed_export_preserves_bundle_directory(managed_tmp_dir: Path):
    schema = _schema({"A1": _cell("A1", "{{rows:dataframe-content}}")}, dims="A1:A1")
    df = polars.DataFrame({"A": [1]})
    bundle_path = managed_tmp_dir / "bundle"
    bundle = mode.compile(
        schema,
        {"Sheet1": {"rows": df}},
        bundle_path=str(bundle_path),
    )

    with pytest.raises(ValueError, match="Fidelity XLSX export does not support"):
        mode.export(
            bundle,
            str(managed_tmp_dir / "out.xlsx"),
            auto_delete_bundle=True,
        )

    assert bundle_path.exists()


# §5 Entrypoints
