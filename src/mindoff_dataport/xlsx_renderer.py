from __future__ import annotations

import datetime
import json
from copy import copy
import shutil
from pathlib import Path
from typing import Any, Iterable, Iterator
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
import pyarrow.parquet as pq
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Border
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)

from .xlsx_builder import (
    _apply_cell_styles,
    _apply_cell_value,
    _apply_dimensions,
    _apply_hug_columns,
    _apply_hug_rows,
    _build_alignment,
    _build_border,
    _build_fill,
    _build_font,
)
from .bundle import ReportBundle, load_report_bundle
from .page_breaks import apply_manual_breaks
from .style_conversion import resolve_theme_color
from .template_contract import _infer_cell_type, _parse_dims
from .schema import CellSchema, SheetSchema

__all__ = ["export_report_bundle"]

# §1. Constants & Exceptions

MAX_EXCEL_ROWS = 1_048_576

# §2. Classes and Sub Classes

# §3. Private Helper Functions


def _coerce_bundle(bundle_or_path: ReportBundle | str) -> ReportBundle:
    if isinstance(bundle_or_path, ReportBundle):
        return bundle_or_path
    return load_report_bundle(str(bundle_or_path))


def _data_source_map(bundle: ReportBundle) -> dict[str, dict[str, Any]]:
    return {
        source["path"]: source for source in bundle.manifest.get("dataframe_sources", [])
    }


def _source_rows(
    bundle: ReportBundle, source: dict[str, Any], *, batch_size: int
) -> Iterator[tuple[Any, ...]]:
    path = source["path"]
    if source["format"] == "parquet":
        yield from _parquet_rows(_source_path(bundle, path), source["columns"], batch_size)
        return
    raise ValueError(f"Unsupported dataframe source format: {source['format']!r}")


def _cached_source_rows(
    bundle: ReportBundle, source: dict[str, Any], *, batch_size: int
) -> Iterable[tuple[Any, ...]]:
    cell_count = int(source.get("rows", 0)) * len(source.get("columns", []))
    if cell_count > 100_000:
        return _source_rows(bundle, source, batch_size=batch_size)
    cache = getattr(bundle, "_mindoff_row_cache", None)
    if cache is None:
        cache = {}
        object.__setattr__(bundle, "_mindoff_row_cache", cache)
    key = source["path"]
    rows = cache.get(key)
    if rows is None:
        rows = list(_source_rows(bundle, source, batch_size=batch_size))
        cache[key] = rows
    return rows


def _source_dict_rows(
    bundle: ReportBundle, source: dict[str, Any], *, batch_size: int
) -> Iterator[dict[str, Any]]:
    columns = [str(column) for column in source["columns"]]
    for row in _source_rows(bundle, source, batch_size=batch_size):
        yield dict(zip(columns, row))


def _source_path(bundle: ReportBundle, relative_path: str) -> Path:
    path = Path(bundle.path) / relative_path
    if not path.exists():
        raise ValueError(f"Report bundle is missing dataframe source '{relative_path}'")
    return path


def _parquet_rows(
    path: Path, columns: list[str], batch_size: int
) -> Iterator[tuple[Any, ...]]:
    parquet_file = pq.ParquetFile(path)
    for batch in parquet_file.iter_batches(batch_size=batch_size, columns=columns):
        column_values = [batch.column(index).to_pylist() for index in range(batch.num_columns)]
        yield from zip(*column_values)


def _header_cells(anchor: dict[str, Any]) -> Iterable[tuple[str, CellSchema]]:
    cell = anchor["cell"]
    start_row = anchor["start_row"]
    start_col = anchor["start_col"]
    bold_font = dict(cell["font"])
    bold_font["bold"] = True
    for layout in _anchor_layouts(anchor):
        coord = f"{get_column_letter(start_col + layout['start_col_offset'])}{start_row}"
        header_cell = dict(cell)
        header_cell["coordinate"] = coord
        header_cell["value"] = str(layout["name"])
        header_cell["cell_type"] = "string"
        header_cell["font"] = bold_font
        header_cell["alignment"] = _layout_alignment(cell, layout)
        yield coord, header_cell  # type: ignore[misc]


def _header_cells_at(anchor: dict[str, Any], row_idx: int) -> Iterable[CellSchema]:
    cell = anchor["cell"]
    start_col = anchor["start_col"]
    bold_font = dict(cell["font"])
    bold_font["bold"] = True
    for layout in _anchor_layouts(anchor):
        header_cell = dict(cell)
        header_cell["coordinate"] = f"{get_column_letter(start_col + layout['start_col_offset'])}{row_idx}"
        header_cell["value"] = str(layout["name"])
        header_cell["cell_type"] = "string"
        header_cell["font"] = bold_font
        header_cell["alignment"] = _layout_alignment(cell, layout)
        yield header_cell  # type: ignore[misc]


def _content_cells(
    bundle: ReportBundle,
    source_map: dict[str, dict[str, Any]],
    anchor: dict[str, Any],
    *,
    batch_size: int,
) -> Iterator[tuple[str, CellSchema]]:
    cell = anchor["cell"]
    start_row = anchor["start_row"]
    start_col = anchor["start_col"]
    source = source_map[anchor["source"]]
    layouts = _anchor_layouts(anchor)
    for row_offset, row in enumerate(_source_rows(bundle, source, batch_size=batch_size)):
        for layout, value in zip(layouts, row):
            coord = f"{get_column_letter(start_col + layout['start_col_offset'])}{start_row + row_offset}"
            content_cell = dict(cell)
            content_cell["coordinate"] = coord
            content_cell["value"] = value
            content_cell["cell_type"] = _infer_cell_type(value)
            content_cell["alignment"] = _layout_alignment(cell, layout)
            yield coord, content_cell  # type: ignore[misc]


def _anchor_layouts(anchor: dict[str, Any]) -> list[dict[str, Any]]:
    layouts = anchor.get("column_layouts")
    if layouts:
        return layouts
    return [
        {"name": name, "start_col_offset": offset, "occupation": 1}
        for offset, name in enumerate(anchor["columns"])
    ]


def _dataframe_stream_id(anchor: dict[str, Any]) -> str:
    key = anchor.get("key")
    if key:
        return f"key:{key}"
    source = anchor.get("source")
    if source:
        return str(source)
    return "key:"


def _layout_alignment(cell: CellSchema, layout: dict[str, Any]) -> dict[str, Any]:
    alignment = dict(cell["alignment"])
    if layout.get("alignment") is not None:
        alignment["horizontal"] = layout["alignment"]
    return alignment


def _occupied_width(anchor: dict[str, Any]) -> int:
    layouts = _anchor_layouts(anchor)
    if not layouts:
        return 0
    last = layouts[-1]
    return int(last["start_col_offset"]) + int(last["occupation"])


def _layout_merge_range(
    anchor: dict[str, Any], layout: dict[str, Any], row_idx: int
) -> CellRange | None:
    occupation = int(layout["occupation"])
    if occupation <= 1:
        return None
    start_col = int(anchor["start_col"]) + int(layout["start_col_offset"])
    return CellRange(
        min_col=start_col,
        min_row=row_idx,
        max_col=start_col + occupation - 1,
        max_row=row_idx,
    )


def _anchor_row_merges(anchor: dict[str, Any], row_idx: int) -> list[dict[str, Any]]:
    merges: list[dict[str, Any]] = []
    for layout in _anchor_layouts(anchor):
        occupation = int(layout["occupation"])
        if occupation <= 1:
            continue
        start_col = int(anchor["start_col"]) + int(layout["start_col_offset"])
        merges.append(
            {
                "min_row_offset": 0,
                "max_row_offset": 0,
                "min_col": start_col,
                "max_col": start_col + occupation - 1,
            }
        )
    return merges


def _generated_cell_merges(
    anchor: dict[str, Any], row_idx: int, col_idx: int
) -> list[CellRange]:
    result: list[CellRange] = []
    for layout in _anchor_layouts(anchor):
        if int(anchor["start_col"]) + int(layout["start_col_offset"]) != col_idx:
            continue
        merge_range = _layout_merge_range(anchor, layout, row_idx)
        if merge_range is not None:
            result.append(merge_range)
    return result


def _render_fidelity(
    bundle: ReportBundle,
    output_path: str,
    *,
    column_width_mode: str | None = None,
    row_height_mode: str | None = None,
    default_column_width: float | None = None,
    default_row_height: float | None = None,
) -> None:
    if any(sheet.get("repeat_sections") for sheet in bundle.report["sheets"]):
        raise ValueError(
            "Fidelity XLSX export does not support repeat sections. Use export_mode='streaming'."
        )

    source_map = _data_source_map(bundle)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for raw_sheet in bundle.report["sheets"]:
        sheet = _sheet_with_overrides(
            raw_sheet,
            column_width_mode=column_width_mode,
            row_height_mode=row_height_mode,
            default_column_width=default_column_width,
            default_row_height=default_row_height,
        )
        sheet, cells = _expanded_cells(bundle, sheet, streaming_chunk_rows=50_000)
        ws = wb.create_sheet(title=sheet["name"])
        _apply_sheet_view(ws, sheet)
        _apply_dimensions(ws, sheet)
        for cell_schema in cells.values():
            cell = ws[cell_schema["coordinate"]]
            _apply_cell_value(cell, cell_schema)
            _apply_cell_styles(cell, cell_schema)

        _, _, max_col, max_row = _parse_dims(sheet["dimensions"])

        for region in sheet["merged_regions"]:
            ws.merge_cells(region)
        _apply_merged_region_borders(ws, sheet, cells)

        if sheet.get("column_width_mode", "fixed") == "hug":
            _apply_hug_columns(ws, _dimensioned_sheet(sheet, max_col, max_row))
        if sheet.get("row_height_mode", "fixed") == "fixed":
            _apply_fixed_dataframe_row_heights_fidelity(ws, sheet, source_map)
        if sheet.get("row_height_mode", "fixed") == "hug":
            _apply_hug_rows(ws, _dimensioned_sheet(sheet, max_col, max_row))
        apply_manual_breaks(
            ws,
            row_breaks=sheet.get("resolved_row_page_breaks", sheet.get("row_page_breaks")),
            column_breaks=sheet.get(
                "resolved_column_page_breaks",
                sheet.get("column_page_breaks"),
            ),
        )

    wb.save(output_path)


def _dimensioned_sheet(
    sheet: SheetSchema, max_col: int, max_row: int
) -> SheetSchema:
    result = dict(sheet)
    min_col, min_row, _, _ = _parse_dims(sheet["dimensions"])
    result["dimensions"] = (
        f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    )
    return result  # type: ignore[return-value]


def _expanded_sheet_from_manifest(
    sheet: SheetSchema, source_map: dict[str, dict[str, Any]]
) -> SheetSchema:
    _, _, max_col, max_row = _parse_dims(sheet["dimensions"])
    for anchor in sheet.get("dataframe_anchors", []):
        max_col = max(max_col, anchor["start_col"] + max(_occupied_width(anchor) - 1, 0))
        if anchor["placeholder_type"] == "dataframe-content":
            source = source_map[anchor["source"]]
            max_row = max(max_row, anchor["start_row"] + max(source.get("rows", 0) - 1, 0))
        else:
            max_row = max(max_row, anchor["start_row"])
    return _dimensioned_sheet(sheet, max_col, max_row)


def _expanded_cells(
    bundle: ReportBundle,
    sheet: SheetSchema,
    *,
    streaming_chunk_rows: int,
) -> tuple[SheetSchema, dict[tuple[int, int], CellSchema]]:
    source_map = _data_source_map(bundle)
    sheet = _expanded_sheet_from_manifest(sheet, source_map)
    cells: dict[tuple[int, int], CellSchema] = {}
    generated_regions: list[str] = []

    for cell in sheet["cells"].values():
        if _is_non_anchor_merged_cell(cell):
            continue
        cells[_coord_indexes(cell["coordinate"])] = cell

    for anchor in sheet.get("dataframe_anchors", []):
        anchor_cells = (
            _header_cells(anchor)
            if anchor["placeholder_type"] == "dataframe-header"
            else _content_cells(
                bundle,
                source_map,
                anchor,
                batch_size=streaming_chunk_rows,
            )
        )
        for coord, cell in anchor_cells:
            row_idx, col_idx = _coord_indexes(coord)
            cells[(row_idx, col_idx)] = cell
            for merge_range in _generated_cell_merges(anchor, row_idx, col_idx):
                generated_regions.append(str(merge_range))

    if generated_regions:
        sheet = dict(sheet)
        sheet["merged_regions"] = [*sheet.get("merged_regions", []), *generated_regions]
    return sheet, cells


def _sheet_with_overrides(
    sheet: dict[str, Any],
    *,
    column_width_mode: str | None,
    row_height_mode: str | None,
    default_column_width: float | None,
    default_row_height: float | None,
) -> SheetSchema:
    result = dict(sheet)
    for key, value in {
        "column_width_mode": column_width_mode,
        "row_height_mode": row_height_mode,
        "default_column_width": default_column_width,
        "default_row_height": default_row_height,
    }.items():
        if value is not None:
            result[key] = value
    return result  # type: ignore[return-value]


def _validate_streaming(
    bundle: ReportBundle,
    source_map: dict[str, dict[str, Any]],
    column_width_mode: str | None,
    row_height_mode: str | None,
) -> None:
    for sheet in bundle.report["sheets"]:
        col_mode = column_width_mode or sheet.get("column_width_mode", "fixed")
        row_mode = row_height_mode or sheet.get("row_height_mode", "fixed")
        if col_mode == "hug" or row_mode == "hug":
            raise ValueError(
                "Streaming XLSX export does not support 'hug' sizing. Use fixed/even sizing or export_mode='fidelity'."
            )
        if sheet.get("repeat_sections"):
            continue
        content_anchors = [
            anchor
            for anchor in sheet.get("dataframe_anchors", [])
            if anchor["placeholder_type"] == "dataframe-content"
        ]
        if len(content_anchors) > 1:
            raise ValueError(
                "Streaming XLSX export currently supports one dataframe-content placeholder per sheet."
            )
        _validate_streaming_merges(sheet, source_map, content_anchors)


def _validate_streaming_merges(
    sheet: SheetSchema,
    source_map: dict[str, dict[str, Any]],
    content_anchors: list[dict[str, Any]],
) -> None:
    if not sheet.get("merged_regions") or not content_anchors:
        return

    content_ranges = [
        _dataframe_content_range(anchor, source_map[anchor["source"]])
        for anchor in content_anchors
    ]
    for region in sheet["merged_regions"]:
        merge_range = CellRange(region)
        for content_range in content_ranges:
            if content_range is not None and _ranges_overlap(merge_range, content_range):
                raise ValueError(
                    "Streaming XLSX export does not support merged cells that intersect dataframe-content output. "
                    "Move the merge outside streamed rows or use export_mode='fidelity'."
                )


def _dataframe_content_range(
    anchor: dict[str, Any], source: dict[str, Any]
) -> CellRange | None:
    row_count = int(source.get("rows", 0))
    col_count = _occupied_width(anchor)
    if row_count <= 0 or col_count <= 0:
        return None
    return CellRange(
        min_col=anchor["start_col"],
        min_row=anchor["start_row"],
        max_col=anchor["start_col"] + col_count - 1,
        max_row=anchor["start_row"] + row_count - 1,
    )


def _anchor_row_height(schema: SheetSchema, anchor: dict[str, Any]) -> float | None:
    height = schema.get("row_heights", {}).get(str(anchor["start_row"]))
    if height is None:
        return None
    return float(height)


def _apply_fixed_dataframe_row_heights_fidelity(
    ws, schema: SheetSchema, source_map: dict[str, dict[str, Any]]
) -> None:
    for anchor in schema.get("dataframe_anchors", []):
        if anchor["placeholder_type"] != "dataframe-content":
            continue
        anchor_height = _anchor_row_height(schema, anchor)
        if anchor_height is None:
            continue
        source = source_map[anchor["source"]]
        row_count = int(source.get("rows", 0))
        for row_idx in range(anchor["start_row"], anchor["start_row"] + row_count):
            if str(row_idx) in schema.get("row_heights", {}):
                continue
            ws.row_dimensions[row_idx].height = anchor_height


def _apply_fixed_dataframe_row_height_streaming(
    ws, schema: SheetSchema, anchor: dict[str, Any], row_idx: int
) -> None:
    if str(row_idx) in schema.get("row_heights", {}):
        return
    anchor_height = _anchor_row_height(schema, anchor)
    if anchor_height is None:
        return
    ws.row_dimensions[row_idx].height = anchor_height


def _ranges_overlap(left: CellRange, right: CellRange) -> bool:
    return (
        left.min_col <= right.max_col
        and left.max_col >= right.min_col
        and left.min_row <= right.max_row
        and left.max_row >= right.min_row
    )


def _render_streaming(
    bundle: ReportBundle,
    output_path: str,
    *,
    streaming_engine: str,
    column_width_mode: str | None,
    row_height_mode: str | None,
    default_column_width: float | None,
    default_row_height: float | None,
    streaming_chunk_rows: int,
    max_rows_per_workbook: int,
) -> list[str]:
    if streaming_chunk_rows <= 0:
        raise ValueError(
            f"streaming_chunk_rows must be greater than 0, got {streaming_chunk_rows}"
        )
    if max_rows_per_workbook <= 0 or max_rows_per_workbook > MAX_EXCEL_ROWS:
        raise ValueError(
            f"max_rows_per_workbook must be between 1 and {MAX_EXCEL_ROWS}, got {max_rows_per_workbook}"
        )
    source_map = _data_source_map(bundle)
    _validate_streaming(bundle, source_map, column_width_mode, row_height_mode)
    plans = [
        _streaming_plan(
            bundle,
            source_map,
            _sheet_with_overrides(
                sheet,
                column_width_mode=column_width_mode,
                row_height_mode=row_height_mode,
                default_column_width=default_column_width,
                default_row_height=default_row_height,
            ),
            streaming_chunk_rows=streaming_chunk_rows,
        )
        for sheet in bundle.report["sheets"]
    ]
    if streaming_engine == "xlsxwriter":
        return _render_streaming_xlsxwriter(output_path, plans, max_rows_per_workbook)
    content_iters = [
        plan["content"]
        for plan in plans
        if plan["content"] is not None
    ]

    output_paths: list[str] = []
    part = 1
    while part == 1 or any(not item["exhausted"] for item in content_iters):
        wb = openpyxl.Workbook(write_only=True)
        wrote_any = False
        for plan in plans:
            ws = wb.create_sheet(title=plan["sheet"]["name"])
            _apply_sheet_view(ws, plan["sheet"])
            _apply_streaming_dimensions(ws, plan["sheet"])
            _apply_streaming_merges(ws, plan["sheet"])
            wrote = _write_streaming_sheet(ws, plan, max_rows_per_workbook)
            wrote_any = wrote_any or wrote
        if part > 1 and not wrote_any:
            break
        final_path = _part_path(output_path, part)
        wb.save(final_path)
        output_paths.append(final_path)
        part += 1

    return _bundle_parts_if_needed(output_path, output_paths)


def _streaming_plan(
    bundle: ReportBundle,
    source_map: dict[str, dict[str, Any]],
    sheet: SheetSchema,
    *,
    streaming_chunk_rows: int,
) -> dict[str, Any]:
    static: dict[tuple[int, int], CellSchema] = {}
    generated_merges: dict[int, list[CellRange]] = {}
    min_col, min_row, max_col, max_row = _parse_dims(sheet["dimensions"])
    content = None
    repeat_rows = None

    for cell in sheet["cells"].values():
        if _is_non_anchor_merged_cell(cell):
            continue
        col_letter, row_idx = coordinate_from_string(cell["coordinate"])
        static[(row_idx, column_index_from_string(col_letter))] = cell
    static.update(_merged_region_edge_schemas(sheet))

    for anchor in sheet.get("dataframe_anchors", []):
        if anchor["placeholder_type"] == "dataframe-header":
            for coord, cell in _header_cells(anchor):
                row_idx = int("".join(ch for ch in coord if ch.isdigit()))
                col_idx = column_index_from_string(
                    "".join(ch for ch in coord if ch.isalpha())
                )
                static[(row_idx, col_idx)] = cell
            for layout in _anchor_layouts(anchor):
                merge_range = _layout_merge_range(anchor, layout, anchor["start_row"])
                if merge_range is not None:
                    generated_merges.setdefault(anchor["start_row"], []).append(merge_range)
            max_col = max(max_col, anchor["start_col"] + max(_occupied_width(anchor) - 1, 0))
            continue

        source = source_map[anchor["source"]]
        content = {
            "anchor": anchor,
            "rows": _source_rows(
                bundle,
                source,
                batch_size=streaming_chunk_rows,
            ),
            "exhausted": False,
        }
        max_col = max(max_col, anchor["start_col"] + max(_occupied_width(anchor) - 1, 0))

    if sheet.get("repeat_sections"):
        repeat_rows = _repeat_row_stream(
            bundle,
            source_map,
            sheet,
            streaming_chunk_rows=streaming_chunk_rows,
        )
        content = {"exhausted": False}
        max_col = _repeat_max_col(sheet, max_col)

    return {
        "sheet": sheet,
        "theme_colors": bundle.report.get("theme_colors")
        or bundle.manifest.get("theme_colors"),
        "static": static,
        "generated_merges": generated_merges,
        "content": content,
        "repeat_rows": repeat_rows,
        "min_col": min_col,
        "min_row": min_row,
        "max_col": max_col,
        "max_row": max_row,
    }


def _render_streaming_xlsxwriter(
    output_path: str,
    plans: list[dict[str, Any]],
    max_rows_per_workbook: int,
) -> list[str]:
    import xlsxwriter

    content_iters = [plan["content"] for plan in plans if plan["content"] is not None]
    output_paths: list[str] = []
    part = 1
    while part == 1 or any(not item["exhausted"] for item in content_iters):
        final_path = _part_path(output_path, part)
        workbook = xlsxwriter.Workbook(final_path, {"constant_memory": True})
        wrote_any = False
        for plan in plans:
            worksheet = workbook.add_worksheet(plan["sheet"]["name"])
            _apply_xlsxwriter_sheet_options(workbook, worksheet, plan["sheet"])
            wrote = _write_xlsxwriter_streaming_sheet(
                workbook,
                worksheet,
                plan,
                max_rows_per_workbook,
            )
            wrote_any = wrote_any or wrote
        workbook.close()
        if part > 1 and not wrote_any:
            Path(final_path).unlink(missing_ok=True)
            break
        output_paths.append(final_path)
        part += 1
    return _bundle_parts_if_needed(output_path, output_paths)


def _apply_xlsxwriter_sheet_options(workbook, worksheet, schema: SheetSchema) -> None:
    del workbook
    if not schema.get("show_gridlines", True):
        worksheet.hide_gridlines(2)
    col_mode = schema.get("column_width_mode", "fixed")
    row_mode = schema.get("row_height_mode", "fixed")
    min_col, min_row, max_col, max_row = _parse_dims(schema["dimensions"])
    if col_mode == "fixed":
        for col_letter, width in schema["column_widths"].items():
            if width is not None:
                col_idx = column_index_from_string(col_letter) - 1
                worksheet.set_column(col_idx, col_idx, width)
    elif col_mode == "even":
        width = schema.get("default_column_width", 15.0)
        worksheet.set_column(min_col - 1, max_col - 1, width)
    if row_mode == "fixed":
        for row_str, height in schema["row_heights"].items():
            if height is not None:
                worksheet.set_row(int(row_str) - 1, height)
    elif row_mode == "even":
        height = schema.get("default_row_height", 15.0)
        for row_idx in range(min_row, max_row + 1):
            worksheet.set_row(row_idx - 1, height)


def _apply_xlsxwriter_static_merges(worksheet, schema: SheetSchema) -> None:
    for region in schema.get("merged_regions", []):
        cell_range = CellRange(region)
        _add_xlsxwriter_merge(
            worksheet,
            cell_range.min_row - 1,
            cell_range.min_col - 1,
            cell_range.max_row - 1,
            cell_range.max_col - 1,
        )


def _write_xlsxwriter_streaming_sheet(
    workbook,
    worksheet,
    plan: dict[str, Any],
    max_rows_per_workbook: int,
) -> bool:
    if plan.get("repeat_rows") is not None:
        return _write_xlsxwriter_repeat_sheet(
            workbook,
            worksheet,
            plan,
            max_rows_per_workbook,
        )

    content = plan["content"]
    anchor = content["anchor"] if content is not None else None
    max_col = plan["max_col"]
    if anchor is not None:
        max_col = max(max_col, anchor["start_col"] + max(_occupied_width(anchor) - 1, 0))
    wrote_content = False
    written_rows = 0

    for row_idx in range(plan["min_row"], max(plan["max_row"], plan["min_row"]) + 1):
        if row_idx > max_rows_per_workbook:
            break
        _write_xlsxwriter_static_row(workbook, worksheet, plan, row_idx, max_col)
        if anchor is not None and row_idx >= anchor["start_row"]:
            wrote_content = _fill_xlsxwriter_streaming_row(
                workbook, worksheet, plan, anchor, content, row_idx
            ) or wrote_content
        for merge_range in plan.get("generated_merges", {}).get(row_idx, []):
            _add_xlsxwriter_cell_range_merge(worksheet, merge_range)
        written_rows += 1

    if anchor is not None:
        row_idx = max(plan["max_row"] + 1, anchor["start_row"])
        while row_idx <= max_rows_per_workbook and not content["exhausted"]:
            wrote_content = _fill_xlsxwriter_streaming_row(
                workbook, worksheet, plan, anchor, content, row_idx
            ) or wrote_content
            if content["exhausted"]:
                break
            written_rows += 1
            row_idx += 1
    _apply_xlsxwriter_static_merges(worksheet, plan["sheet"])
    _apply_xlsxwriter_page_breaks(worksheet, plan["sheet"], written_rows)
    return wrote_content


def _write_xlsxwriter_static_row(
    workbook,
    worksheet,
    plan: dict[str, Any],
    row_idx: int,
    max_col: int,
) -> None:
    row_zero = row_idx - 1
    for col_idx in range(plan["min_col"], max_col + 1):
        static = plan["static"].get((row_idx, col_idx))
        if static is not None:
            _write_xlsxwriter_cell(
                workbook,
                worksheet,
                row_zero,
                col_idx - 1,
                static,
                plan.get("theme_colors"),
            )


def _fill_xlsxwriter_streaming_row(
    workbook,
    worksheet,
    plan: dict[str, Any],
    anchor: dict[str, Any],
    content: dict[str, Any],
    row_idx: int,
) -> bool:
    try:
        row_values = next(content["rows"])
    except StopIteration:
        content["exhausted"] = True
        return False

    if plan["sheet"].get("row_height_mode", "fixed") == "fixed":
        _apply_fixed_dataframe_row_height_xlsxwriter(worksheet, plan["sheet"], anchor, row_idx)
    for layout, value in zip(_anchor_layouts(anchor), row_values):
        for col_idx, cell_schema in _layout_row_cells(
            anchor,
            layout,
            row_idx=row_idx,
            value=value,
        ).items():
            if col_idx < plan["min_col"] or col_idx > plan["max_col"]:
                continue
            _write_xlsxwriter_cell(
                workbook,
                worksheet,
                row_idx - 1,
                col_idx - 1,
                cell_schema,
                plan.get("theme_colors"),
            )
        merge_range = _layout_merge_range(anchor, layout, row_idx)
        if merge_range is not None:
            _add_xlsxwriter_cell_range_merge(worksheet, merge_range)
    return True


def _apply_fixed_dataframe_row_height_xlsxwriter(
    worksheet,
    schema: SheetSchema,
    anchor: dict[str, Any],
    row_idx: int,
) -> None:
    if schema.get("row_height_mode", "fixed") != "fixed":
        return
    if str(row_idx) in schema.get("row_heights", {}):
        return
    height = schema.get("row_heights", {}).get(str(anchor["start_row"]))
    if height is not None:
        worksheet.set_row(row_idx - 1, height)


def _write_xlsxwriter_repeat_sheet(
    workbook,
    worksheet,
    plan: dict[str, Any],
    max_rows_per_workbook: int,
) -> bool:
    wrote_any = False
    row_idx = 1
    written_rows = 0
    while row_idx <= max_rows_per_workbook:
        try:
            row_item = next(plan["repeat_rows"])
        except StopIteration:
            plan["content"]["exhausted"] = True
            break
        _write_xlsxwriter_row(
            workbook,
            worksheet,
            row_idx,
            row_item,
            plan.get("theme_colors"),
        )
        wrote_any = True
        written_rows += 1
        row_idx += 1
    _apply_xlsxwriter_page_breaks(worksheet, plan["sheet"], written_rows)
    return wrote_any


def _apply_xlsxwriter_page_breaks(
    worksheet,
    schema: SheetSchema,
    written_rows: int,
) -> None:
    worksheet.set_h_pagebreaks(_row_page_breaks_for_written_rows(schema, written_rows))
    worksheet.set_v_pagebreaks(
        schema.get(
            "resolved_column_page_breaks",
            schema.get("column_page_breaks", []),
        )
        or []
    )


def _row_page_breaks_for_written_rows(
    schema: SheetSchema,
    written_rows: int,
) -> list[int]:
    return [
        break_idx
        for break_idx in schema.get(
            "resolved_row_page_breaks",
            schema.get("row_page_breaks", []),
        )
        if break_idx < written_rows
    ]


def _write_xlsxwriter_row(
    workbook,
    worksheet,
    row_idx: int,
    row_item: dict[str, Any],
    theme_colors: list[str] | None,
) -> None:
    row_zero = row_idx - 1
    row_map = row_item["cells"]
    for col_idx, schema in row_map.items():
        _write_xlsxwriter_cell(
            workbook,
            worksheet,
            row_zero,
            col_idx - 1,
            schema,
            theme_colors,
        )
    for merge in row_item.get("merges", []):
        max_merge_row = row_idx + merge["max_row_offset"] - merge["min_row_offset"]
        first_col = int(merge["min_col"])
        last_col = int(merge["max_col"])
        _add_xlsxwriter_merge(
            worksheet,
            row_zero,
            first_col - 1,
            max_merge_row - 1,
            last_col - 1,
        )


def _add_xlsxwriter_merge(
    worksheet,
    first_row: int,
    first_col: int,
    last_row: int,
    last_col: int,
) -> None:
    from xlsxwriter.utility import xl_range

    cell_range = xl_range(first_row, first_col, last_row, last_col)
    worksheet.merge.append([first_row, first_col, last_row, last_col])
    for row_idx in range(first_row, last_row + 1):
        for col_idx in range(first_col, last_col + 1):
            worksheet.merged_cells[(row_idx, col_idx)] = cell_range


def _add_xlsxwriter_cell_range_merge(worksheet, merge_range: CellRange) -> None:
    _add_xlsxwriter_merge(
        worksheet,
        merge_range.min_row - 1,
        merge_range.min_col - 1,
        merge_range.max_row - 1,
        merge_range.max_col - 1,
    )


def _write_xlsxwriter_cell(
    workbook,
    worksheet,
    row: int,
    col: int,
    schema: CellSchema,
    theme_colors: list[str] | None,
) -> None:
    fmt = _xlsxwriter_format(workbook, schema, theme_colors)
    value = _xlsxwriter_value(schema)
    if schema.get("cell_type") == "formula":
        raise ValueError(
            "XlsxWriter streaming export does not support formula cells. Use streaming_engine='openpyxl'."
        )
    if value is None:
        worksheet.write_blank(row, col, None, fmt)
    else:
        worksheet.write(row, col, value, fmt)


def _xlsxwriter_value(schema: CellSchema | None) -> Any:
    if schema is None:
        return None
    value = schema["value"]
    if schema["cell_type"] == "date" and isinstance(value, str):
        return datetime.datetime.fromisoformat(value)
    return value


def _xlsxwriter_format(
    workbook,
    schema: CellSchema | None,
    theme_colors: list[str] | None,
):
    if schema is None:
        return None
    cache = getattr(workbook, "_mindoff_format_cache", None)
    if cache is None:
        cache = {}
        setattr(workbook, "_mindoff_format_cache", cache)
    key = _style_cache_key(schema)
    fmt = cache.get(key)
    if fmt is None:
        fmt = workbook.add_format(_xlsxwriter_format_props(schema, theme_colors))
        cache[key] = fmt
    return fmt


def _xlsxwriter_format_props(
    schema: CellSchema,
    theme_colors: list[str] | None = None,
) -> dict[str, Any]:
    props: dict[str, Any] = {}
    font = schema["font"]
    if font.get("name"):
        props["font_name"] = font["name"]
    if font.get("size"):
        props["font_size"] = font["size"]
    if font.get("bold"):
        props["bold"] = True
    if font.get("italic"):
        props["italic"] = True
    if font.get("underline"):
        props["underline"] = True
    if font.get("strike"):
        props["font_strikeout"] = True
    if font.get("vert_align") == "superscript":
        props["font_script"] = 1
    elif font.get("vert_align") == "subscript":
        props["font_script"] = 2
    if font.get("color"):
        color = _xlsxwriter_color(font["color"], theme_colors)
        if color is not None:
            props["font_color"] = color

    fill = schema["fill"]
    pattern_type = fill.get("pattern_type") or (fill.get("bg_color") and "solid")
    if pattern_type:
        fill_color = _xlsxwriter_color(
            fill.get("fg_color") or fill.get("bg_color"),
            theme_colors,
        )
        pattern_color = _xlsxwriter_color(fill.get("bg_color"), theme_colors)
        if fill_color is None and pattern_color is None:
            pattern_type = None
    if pattern_type:
        pattern = _xlsxwriter_pattern(pattern_type)
        if pattern is not None:
            props["pattern"] = pattern
        if fill_color is not None:
            props["bg_color"] = fill_color
        if pattern_color is not None and pattern_type != "solid":
            props["fg_color"] = fill_color
            props["bg_color"] = pattern_color

    alignment = schema["alignment"]
    if alignment.get("horizontal"):
        props["align"] = alignment["horizontal"]
    if alignment.get("vertical"):
        props["valign"] = alignment["vertical"]
    if alignment.get("wrap_text"):
        props["text_wrap"] = True
    if alignment.get("indent") is not None:
        props["indent"] = alignment["indent"]
    if alignment.get("shrink_to_fit"):
        props["shrink"] = True
    if alignment.get("text_rotation") is not None:
        props["rotation"] = alignment["text_rotation"]
    if alignment.get("reading_order") is not None:
        props["reading_order"] = alignment["reading_order"]

    _xlsxwriter_border_props(props, schema["borders"], alignment, theme_colors)
    if schema.get("number_format"):
        props["num_format"] = schema["number_format"]
    return props


def _xlsxwriter_pattern(pattern_type: str | None) -> int | None:
    return {
        "solid": 1,
        "darkGray": 2,
        "mediumGray": 3,
        "lightGray": 4,
        "gray125": 17,
    }.get(pattern_type)


def _xlsxwriter_border_props(
    props: dict[str, Any],
    borders: dict[str, Any],
    alignment: dict[str, Any],
    theme_colors: list[str] | None,
) -> None:
    for side in ("top", "bottom", "left", "right"):
        data = borders.get(side) or {}
        style = _xlsxwriter_border_style(data.get("style"))
        if style is None:
            continue
        props[side] = style
        if data.get("color"):
            color = _xlsxwriter_color(data["color"], theme_colors)
            if color is not None:
                props[f"{side}_color"] = color
    for logical_side, physical_side in _xlsxwriter_logical_border_sides(alignment).items():
        data = borders.get(logical_side) or {}
        style = _xlsxwriter_border_style(data.get("style"))
        if style is None:
            continue
        props[physical_side] = style
        if data.get("color"):
            color = _xlsxwriter_color(data["color"], theme_colors)
            if color is not None:
                props[f"{physical_side}_color"] = color
    diagonal = borders.get("diagonal") or {}
    diagonal_style = _xlsxwriter_border_style(diagonal.get("style"))
    if diagonal_style is not None:
        props["diag_border"] = diagonal_style
        if diagonal.get("color"):
            color = _xlsxwriter_color(diagonal["color"], theme_colors)
            if color is not None:
                props["diag_color"] = color
    if borders.get("diagonal_up") and borders.get("diagonal_down"):
        props["diag_type"] = 3
    elif borders.get("diagonal_up"):
        props["diag_type"] = 2
    elif borders.get("diagonal_down"):
        props["diag_type"] = 1


def _xlsxwriter_logical_border_sides(alignment: dict[str, Any]) -> dict[str, str]:
    if alignment.get("reading_order") == 2:
        return {"start": "right", "end": "left"}
    return {"start": "left", "end": "right"}


def _xlsxwriter_border_style(style: str | None) -> int | None:
    return {
        "hair": 7,
        "thin": 1,
        "medium": 2,
        "thick": 5,
        "dashed": 3,
        "dotted": 4,
        "double": 6,
    }.get(style)


def _xlsxwriter_color(
    value: str | None,
    theme_colors: list[str] | None = None,
) -> str | None:
    if not value:
        return None
    resolved = resolve_theme_color(value, theme_colors)
    if resolved is None:
        return None
    raw = resolved[-6:]
    return f"#{raw}" if len(raw) == 6 else None


def _repeat_max_col(sheet: SheetSchema, current: int) -> int:
    max_col = current
    for section in sheet.get("repeat_sections", []):
        if section.get("record_source"):
            for item in section.get("record_bindings", []):
                max_col = max(max_col, item["start_col"])
            for anchor in section.get("dataframe_anchors", []):
                max_col = max(
                    max_col,
                    anchor["start_col"] + max(_occupied_width(anchor) - 1, 0),
                )
            continue
        for record in section["records"]:
            for item in record["cells"]:
                max_col = max(max_col, item["start_col"])
            for anchor in record["dataframe_anchors"]:
                max_col = max(
                    max_col,
                    anchor["start_col"] + max(_occupied_width(anchor) - 1, 0),
                )
    return max_col


def _repeat_row_stream(
    bundle: ReportBundle,
    source_map: dict[str, dict[str, Any]],
    sheet: SheetSchema,
    *,
    streaming_chunk_rows: int,
) -> Iterator[dict[str, Any]]:
    _, min_row, _, max_row = _parse_dims(sheet["dimensions"])
    cursor = min_row
    output_row = 1
    for section in sheet["repeat_sections"]:
        for row_item in _static_repeat_rows(sheet, cursor, section["start_row"] - 1):
            item = dict(row_item)
            item["row_idx"] = output_row
            yield item
            output_row += 1
        records = (
            _source_repeat_records(bundle, source_map, section, streaming_chunk_rows)
            if section.get("record_source")
            else section["records"]
        )
        for record in records:
            block_height = record.get(
                "block_height",
                section.get("record_block_height", section["block_height"]),
            )
            merges = record.get(
                "merged_regions",
                section.get("merged_record_regions", section.get("merged_regions", [])),
            )
            for row_item in _repeat_record_rows(
                bundle,
                source_map,
                record,
                block_height=block_height,
                merges=merges,
                cell_templates=section.get("cell_templates"),
                batch_size=streaming_chunk_rows,
            ):
                item = dict(row_item)
                item["row_idx"] = output_row
                yield item
                output_row += 1
        cursor = section["end_row"] + 1
    for row_item in _static_repeat_rows(sheet, cursor, max_row):
        item = dict(row_item)
        item["row_idx"] = output_row
        yield item
        output_row += 1


def _source_repeat_records(
    bundle: ReportBundle,
    source_map: dict[str, dict[str, Any]],
    section: dict[str, Any],
    streaming_chunk_rows: int,
) -> Iterator[dict[str, Any]]:
    source = source_map[section["record_source"]]
    constants = section.get("constants", {})
    for index, row in enumerate(
        _source_dict_rows(bundle, source, batch_size=streaming_chunk_rows)
    ):
        payload = {**constants, **row}
        yield {
            "index": index,
            "block_height": section.get("record_block_height", section["block_height"]),
            "cells": [
                _source_repeat_binding(binding, payload)
                for binding in section.get("record_bindings", [])
            ],
            "dataframe_anchors": section.get("dataframe_anchors", []),
            "merged_regions": section.get(
                "merged_record_regions",
                section.get("merged_regions", []),
            ),
        }


def _source_repeat_binding(
    binding: dict[str, Any], payload: dict[str, Any]
) -> dict[str, Any]:
    result = {
        "row_offset": binding["row_offset"],
        "start_col": binding["start_col"],
        "cell_template": binding["cell_template"],
    }
    if "value_template" not in binding:
        value = binding.get("value")
    else:
        value = _resolve_repeat_binding_value(binding, payload)
    result["value"] = value
    result["cell_type"] = _infer_cell_type(value)
    return result


def _resolve_repeat_binding_value(
    binding: dict[str, Any], payload: dict[str, Any]
) -> Any:
    value_template = binding.get("value_template")
    if not isinstance(value_template, str):
        return value_template
    if binding.get("full_scalar") and binding.get("scalar_keys"):
        return payload.get(binding["scalar_keys"][0])

    def replacer(match) -> str:
        key = match.group(1)
        if key not in payload:
            return match.group(0)
        return str(payload[key])

    from .template_contract import PLACEHOLDER_RE

    return PLACEHOLDER_RE.sub(replacer, value_template)


def _static_repeat_rows(
    sheet: SheetSchema, start_row: int, end_row: int
) -> Iterator[dict[str, Any]]:
    if end_row < start_row:
        return
    for row_idx in range(start_row, end_row + 1):
        row_cells: dict[int, CellSchema] = {}
        for cell in sheet["cells"].values():
            cell_row, col_idx = _coord_indexes(cell["coordinate"])
            if cell_row == row_idx:
                row_cells[col_idx] = cell
        yield {"cells": row_cells, "merges": _static_row_merges(sheet, row_idx)}


def _static_row_merges(sheet: SheetSchema, row_idx: int) -> list[dict[str, Any]]:
    merges: list[dict[str, Any]] = []
    for raw_region in sheet.get("merged_regions", []):
        region = CellRange(raw_region)
        if region.min_row != row_idx:
            continue
        merges.append(
            {
                "min_row_offset": 0,
                "max_row_offset": region.max_row - region.min_row,
                "min_col": region.min_col,
                "max_col": region.max_col,
            }
        )
    return merges


def _coord_indexes(coord: str) -> tuple[int, int]:
    col_letter, row_idx = coordinate_from_string(coord)
    return row_idx, column_index_from_string(col_letter)


def _repeat_record_rows(
    bundle: ReportBundle,
    source_map: dict[str, dict[str, Any]],
    record: dict[str, Any],
    *,
    block_height: int,
    merges: list[dict[str, Any]] | None = None,
    cell_templates: list[CellSchema] | None = None,
    batch_size: int,
) -> Iterator[dict[str, Any]]:
    if cell_templates is not None:
        for template in cell_templates:
            _style_cache_key(template)
    cells_by_offset: dict[int, dict[int, CellSchema]] = {}
    for item in record["cells"]:
        cell = _repeat_item_cell(item, cell_templates)
        cells_by_offset.setdefault(item["row_offset"], {})[item["start_col"]] = cell
    _apply_repeat_merge_edge_cells(cells_by_offset, merges or [])

    headers_by_offset: dict[int, list[dict[str, Any]]] = {}
    content_by_offset: dict[int, list[dict[str, Any]]] = {}
    header_rows_by_source: dict[str, dict[str, Any]] = {}
    for anchor in record["dataframe_anchors"]:
        stream_id = _dataframe_stream_id(anchor)
        target = (
            headers_by_offset
            if anchor["placeholder_type"] == "dataframe-header"
            else content_by_offset
        )
        target.setdefault(anchor["start_row_offset"], []).append(anchor)
        if anchor["placeholder_type"] == "dataframe-header":
            header_cells: dict[int, CellSchema] = {}
            for cell in _header_cells_at(anchor, 1):
                _, col_idx = _coord_indexes(cell["coordinate"])
                header_cells[col_idx] = cell
            header_rows_by_source[stream_id] = {
                "cells": header_cells,
                "merges": _anchor_row_merges(anchor, 1),
                "is_dataframe_header_row": True,
            }

    offset = 0
    while offset < block_height:
        row_cells = dict(cells_by_offset.get(offset, {}))
        header_merges: list[dict[str, Any]] = []
        header_sources: list[str] = []
        for anchor in headers_by_offset.get(offset, []):
            for cell in _header_cells_at(anchor, offset + 1):
                _, col_idx = _coord_indexes(cell["coordinate"])
                row_cells[col_idx] = cell
            header_merges.extend(_anchor_row_merges(anchor, offset + 1))
            header_sources.append(_dataframe_stream_id(anchor))

        content_anchors = content_by_offset.get(offset, [])
        if not content_anchors:
            yield {
                "cells": row_cells,
                "merges": _repeat_merges_starting_at(merges or [], offset)
                + header_merges,
                "dataframe_header_sources": header_sources,
                "is_dataframe_header_row": bool(header_sources),
            }
            offset += 1
            continue
        yield from _repeat_content_rows(
            bundle,
            source_map,
            row_cells,
            content_anchors,
            base_merges=header_merges,
            header_rows_by_source=header_rows_by_source,
            batch_size=batch_size,
        )
        anchor = content_anchors[0]
        offset += max(int(anchor.get("source_rows") or 0), 1)


def _repeat_item_cell(
    item: dict[str, Any], cell_templates: list[CellSchema] | None
) -> CellSchema:
    cell = item.get("cell")
    if cell is not None:
        return cell
    if cell_templates is None:
        raise ValueError("Compact repeat record is missing cell_templates")
    result = dict(cell_templates[int(item["cell_template"])])
    result["value"] = item.get("value")
    result["cell_type"] = item.get("cell_type", "empty")
    result["coordinate"] = f"{get_column_letter(int(item['start_col']))}{int(item['row_offset']) + 1}"
    return result  # type: ignore[return-value]


def _apply_repeat_merge_edge_cells(
    cells_by_offset: dict[int, dict[int, CellSchema]],
    merges: list[dict[str, Any]],
) -> None:
    for merge in merges:
        anchor = cells_by_offset.get(merge["min_row_offset"], {}).get(merge["min_col"])
        if anchor is None:
            continue
        cell_range = CellRange(
            min_col=merge["min_col"],
            min_row=merge["min_row_offset"] + 1,
            max_col=merge["max_col"],
            max_row=merge["max_row_offset"] + 1,
        )
        for row_offset in range(merge["min_row_offset"], merge["max_row_offset"] + 1):
            for col_idx in range(merge["min_col"], merge["max_col"] + 1):
                border_schema = _edge_border_schema(
                    anchor["borders"], cell_range, row_offset + 1, col_idx
                )
                if border_schema is None and not (
                    row_offset == merge["min_row_offset"]
                    and col_idx == merge["min_col"]
                ):
                    continue
                edge_cell = dict(anchor)
                edge_cell["coordinate"] = f"{get_column_letter(col_idx)}{row_offset + 1}"
                edge_cell["value"] = (
                    anchor["value"]
                    if row_offset == merge["min_row_offset"]
                    and col_idx == merge["min_col"]
                    else None
                )
                edge_cell["cell_type"] = (
                    anchor["cell_type"]
                    if row_offset == merge["min_row_offset"]
                    and col_idx == merge["min_col"]
                    else "empty"
                )
                if border_schema is not None:
                    edge_cell["borders"] = border_schema
                    edge_cell.pop("__style_key", None)
                cells_by_offset.setdefault(row_offset, {})[col_idx] = edge_cell  # type: ignore[assignment]


def _repeat_merges_starting_at(
    merges: list[dict[str, Any]], offset: int
) -> list[dict[str, Any]]:
    return [merge for merge in merges if merge["min_row_offset"] == offset]


def _repeat_content_rows(
    bundle: ReportBundle,
    source_map: dict[str, dict[str, Any]],
    base_cells: dict[int, CellSchema],
    anchors: list[dict[str, Any]],
    *,
    base_merges: list[dict[str, Any]] | None = None,
    header_rows_by_source: dict[str, dict[str, Any]] | None = None,
    batch_size: int,
) -> Iterator[dict[str, Any]]:
    if len(anchors) > 1:
        raise ValueError(
            "Repeat sections support one dataframe-content placeholder per template row in v1."
        )
    anchor = anchors[0]
    source = source_map[anchor["source"]]
    layouts = _anchor_layouts(anchor)
    wrote = False
    for row_values in _cached_source_rows(bundle, source, batch_size=batch_size):
        first_content_row = not wrote
        row_cells = dict(base_cells) if first_content_row else {}
        for layout, value in zip(layouts, row_values):
            row_cells.update(
                _layout_row_cells(
                    anchor,
                    layout,
                    row_idx=1,
                    value=value,
                )
            )
        wrote = True
        yield {
            "cells": row_cells,
            "merges": (base_merges or []) + _anchor_row_merges(anchor, 1),
            "dataframe_content_sources": [_dataframe_stream_id(anchor)],
            "dataframe_content_start_sources": [_dataframe_stream_id(anchor)] if first_content_row else [],
            "dataframe_header_rows_by_source": header_rows_by_source or {},
        }
    if not wrote:
        yield {"cells": base_cells, "merges": base_merges or []}


def _is_non_anchor_merged_cell(cell: CellSchema) -> bool:
    return bool(cell.get("merged")) and cell.get("merge_anchor") not in (
        None,
        cell["coordinate"],
    )


def _write_streaming_sheet(ws, plan: dict[str, Any], max_rows_per_workbook: int) -> bool:
    if plan.get("repeat_rows") is not None:
        return _write_repeat_streaming_sheet(ws, plan, max_rows_per_workbook)

    content = plan["content"]
    anchor = content["anchor"] if content is not None else None
    max_col = plan["max_col"]
    if anchor is not None:
        max_col = max(max_col, anchor["start_col"] + max(_occupied_width(anchor) - 1, 0))
    col_count = max_col - plan["min_col"] + 1
    wrote_content = False
    written_rows = 0

    for row_idx in range(plan["min_row"], max(plan["max_row"], plan["min_row"]) + 1):
        if row_idx > max_rows_per_workbook:
            break
        row_cells = [None] * col_count
        for col_idx in range(plan["min_col"], max_col + 1):
            static = plan["static"].get((row_idx, col_idx))
            if static is not None:
                row_cells[col_idx - plan["min_col"]] = _write_only_cell(ws, static, plan.get("theme_colors"))
        for merge_range in plan.get("generated_merges", {}).get(row_idx, []):
            _add_generated_merge(ws, merge_range)
        if anchor is not None and row_idx >= anchor["start_row"]:
            wrote_content = _fill_streaming_row(
                ws, row_cells, plan, anchor, content, row_idx
            ) or wrote_content
        ws.append(row_cells)
        written_rows += 1

    if anchor is not None:
        row_idx = max(plan["max_row"] + 1, anchor["start_row"])
        while row_idx <= max_rows_per_workbook and not content["exhausted"]:
            row_cells = [None] * col_count
            wrote_content = _fill_streaming_row(
                ws, row_cells, plan, anchor, content, row_idx
            ) or wrote_content
            if content["exhausted"]:
                break
            ws.append(row_cells)
            written_rows += 1
            row_idx += 1
    apply_manual_breaks(
        ws,
        row_breaks=[
            break_idx
            for break_idx in plan["sheet"].get(
                "resolved_row_page_breaks",
                plan["sheet"].get("row_page_breaks", []),
            )
            if break_idx < written_rows
        ],
        column_breaks=plan["sheet"].get(
            "resolved_column_page_breaks",
            plan["sheet"].get("column_page_breaks"),
        ),
    )
    return wrote_content


def _write_repeat_streaming_sheet(
    ws, plan: dict[str, Any], max_rows_per_workbook: int
) -> bool:
    max_col = plan["max_col"]
    col_count = max_col - plan["min_col"] + 1
    wrote_any = False
    row_idx = 1
    written_rows = 0
    while row_idx <= max_rows_per_workbook:
        try:
            row_item = next(plan["repeat_rows"])
        except StopIteration:
            plan["content"]["exhausted"] = True
            break
        row_map = row_item["cells"]
        for merge in row_item["merges"]:
            max_merge_row = row_idx + merge["max_row_offset"] - merge["min_row_offset"]
            if max_merge_row <= max_rows_per_workbook:
                _add_generated_merge(
                    ws,
                    CellRange(
                        min_col=merge["min_col"],
                        min_row=row_idx,
                        max_col=merge["max_col"],
                        max_row=max_merge_row,
                    ),
                )
        row_cells = [None] * col_count
        for col_idx, schema in row_map.items():
            if col_idx < plan["min_col"] or col_idx > max_col:
                continue
            row_cells[col_idx - plan["min_col"]] = _write_only_cell(ws, schema, plan.get("theme_colors"))
        ws.append(row_cells)
        wrote_any = True
        written_rows += 1
        row_idx += 1
    apply_manual_breaks(
        ws,
        row_breaks=[
            break_idx
            for break_idx in plan["sheet"].get(
                "resolved_row_page_breaks",
                plan["sheet"].get("row_page_breaks", []),
            )
            if break_idx < written_rows
        ],
        column_breaks=plan["sheet"].get(
            "resolved_column_page_breaks",
            plan["sheet"].get("column_page_breaks"),
        ),
    )
    return wrote_any


def _fill_streaming_row(
    ws,
    row_cells: list[Any],
    plan: dict[str, Any],
    anchor: dict[str, Any],
    content: dict[str, Any],
    row_idx: int,
) -> bool:
    try:
        row_values = next(content["rows"])
    except StopIteration:
        content["exhausted"] = True
        return False

    layouts = _anchor_layouts(anchor)
    if plan["sheet"].get("row_height_mode", "fixed") == "fixed":
        _apply_fixed_dataframe_row_height_streaming(ws, plan["sheet"], anchor, row_idx)
    for layout, value in zip(layouts, row_values):
        for col_idx, cell_schema in _layout_row_cells(
            anchor,
            layout,
            row_idx=row_idx,
            value=value,
        ).items():
            if col_idx < plan["min_col"] or col_idx > plan["max_col"]:
                continue
            row_cells[col_idx - plan["min_col"]] = _styled_write_only_cell(
                ws,
                cell_schema,
                cell_schema["value"],
                plan.get("theme_colors"),
            )
        merge_range = _layout_merge_range(anchor, layout, row_idx)
        if merge_range is not None:
            _add_generated_merge(ws, merge_range)
    return True


def _layout_row_cells(
    anchor: dict[str, Any],
    layout: dict[str, Any],
    *,
    row_idx: int,
    value: Any,
) -> dict[int, CellSchema]:
    start_col = int(anchor["start_col"]) + int(layout["start_col_offset"])
    schema = _cell_with_layout(anchor["cell"], layout)
    occupation = int(layout["occupation"])
    if occupation <= 1:
        first = dict(schema)
        first["value"] = value
        first["cell_type"] = _infer_cell_type(value)
        return {start_col: first}  # type: ignore[return-value]

    result: dict[int, CellSchema] = {}
    end_col = start_col + occupation - 1
    edge_style_keys = layout.setdefault("__edge_style_keys", {})
    for col_idx in range(start_col, end_col + 1):
        item = dict(schema)
        border_schema = _edge_border_schema_bounds(
            item["borders"],
            min_row=row_idx,
            max_row=row_idx,
            min_col=start_col,
            max_col=end_col,
            row=row_idx,
            col=col_idx,
        )
        if border_schema is not None:
            item["borders"] = border_schema
            style_key = edge_style_keys.get(col_idx - start_col)
            if style_key is None:
                item.pop("__style_key", None)
                style_key = _style_cache_key(item)
                edge_style_keys[col_idx - start_col] = style_key
            else:
                item["__style_key"] = style_key
        if col_idx == start_col:
            item["value"] = value
            item["cell_type"] = _infer_cell_type(value)
        else:
            item["value"] = None
            item["cell_type"] = "empty"
        result[col_idx] = item  # type: ignore[assignment]
    return result


def _cell_with_layout(schema: CellSchema, layout: dict[str, Any]) -> CellSchema:
    cache = layout.get("__cell_schema_cache")
    if cache is not None:
        return cache
    result = dict(schema)
    result["alignment"] = _layout_alignment(schema, layout)
    _style_cache_key(result)
    layout["__cell_schema_cache"] = result
    return result  # type: ignore[return-value]


def _style_cache_key(schema: dict[str, Any]) -> str:
    cached = schema.get("__style_key")
    if cached is not None:
        return cached
    key = json.dumps(
        {
            "font": schema["font"],
            "fill": schema["fill"],
            "alignment": schema["alignment"],
            "borders": schema["borders"],
            "number_format": schema["number_format"],
        },
        sort_keys=True,
        separators=(",", ":"),
        default=str,
    )
    schema["__style_key"] = key
    return key


def _cached_style(ws, kind: str, schema: dict[str, Any], builder):
    cache = getattr(ws, "_mindoff_style_cache", None)
    if cache is None:
        cache = {"font": {}, "fill": {}, "alignment": {}, "border": {}}
        setattr(ws, "_mindoff_style_cache", cache)
    kind_cache = cache[kind]
    key = json.dumps(schema, sort_keys=True, separators=(",", ":"), default=str)
    style = kind_cache.get(key)
    if style is None:
        style = builder(schema)
        kind_cache[key] = style
    return style


def _resolve_color_field(value: str | None, theme_colors: list[str] | None) -> str | None:
    if value is None or theme_colors is None or not value.startswith("theme:"):
        return value
    return resolve_theme_color(value, theme_colors)


def _resolved_fill(fill: dict[str, Any], theme_colors: list[str] | None) -> dict[str, Any]:
    if theme_colors is None:
        return fill
    result = dict(fill)
    for key in ("fg_color", "bg_color"):
        if result.get(key):
            result[key] = _resolve_color_field(result[key], theme_colors)
    return result


def _resolved_font(font: dict[str, Any], theme_colors: list[str] | None) -> dict[str, Any]:
    if theme_colors is None or not font.get("color"):
        return font
    result = dict(font)
    result["color"] = _resolve_color_field(result["color"], theme_colors)
    return result


def _resolved_borders(borders: dict[str, Any], theme_colors: list[str] | None) -> dict[str, Any]:
    if theme_colors is None:
        return borders
    result = {}
    for key, side in borders.items():
        if isinstance(side, dict) and side.get("color"):
            side = dict(side)
            side["color"] = _resolve_color_field(side["color"], theme_colors)
        result[key] = side
    return result


def _cached_style_array(ws, schema: CellSchema, theme_colors: list[str] | None = None):
    cache = getattr(ws, "_mindoff_style_array_cache", None)
    if cache is None:
        cache = {}
        setattr(ws, "_mindoff_style_array_cache", cache)
    key = _style_cache_key(schema)
    style_array = cache.get(key)
    if style_array is not None:
        return style_array

    prototype = WriteOnlyCell(ws, value=None)
    prototype.font = _cached_style(ws, "font", _resolved_font(schema["font"], theme_colors), _build_font)
    prototype.fill = _cached_style(ws, "fill", _resolved_fill(schema["fill"], theme_colors), _build_fill)
    prototype.alignment = _cached_style(ws, "alignment", schema["alignment"], _build_alignment)
    prototype.border = _cached_style(ws, "border", _resolved_borders(schema["borders"], theme_colors), _build_border)
    if schema["number_format"]:
        prototype.number_format = schema["number_format"]
    cache[key] = copy(prototype._style)
    return cache[key]


def _styled_write_only_cell(
    ws, schema: CellSchema, value: Any, theme_colors: list[str] | None = None
) -> WriteOnlyCell:
    cell = WriteOnlyCell(ws, value=value)
    cell._style = _cached_style_array(ws, schema, theme_colors)
    return cell


def _write_only_cell(ws, schema: CellSchema, theme_colors: list[str] | None = None) -> WriteOnlyCell:
    value = schema["value"]
    if schema["cell_type"] == "date" and isinstance(value, str):
        value = datetime.datetime.fromisoformat(value)
    cell = _styled_write_only_cell(ws, schema, value, theme_colors)
    return cell


def _apply_merged_region_borders(
    ws, sheet: SheetSchema, cells: dict[tuple[int, int], CellSchema] | None = None
) -> None:
    for region in sheet.get("merged_regions", []):
        cell_range = CellRange(region)
        anchor = None
        if cells is not None:
            anchor = cells.get((cell_range.min_row, cell_range.min_col))
        if anchor is None:
            anchor = sheet["cells"].get(
                f"{get_column_letter(cell_range.min_col)}{cell_range.min_row}"
            )
        if anchor is None:
            continue
        for row in range(cell_range.min_row, cell_range.max_row + 1):
            for col in range(cell_range.min_col, cell_range.max_col + 1):
                border = _edge_border(anchor["borders"], cell_range, row, col)
                if border is not None:
                    ws.cell(row=row, column=col).border = border


def _merged_region_edge_schemas(
    sheet: SheetSchema,
) -> dict[tuple[int, int], CellSchema]:
    result: dict[tuple[int, int], CellSchema] = {}
    for region in sheet.get("merged_regions", []):
        cell_range = CellRange(region)
        anchor = sheet["cells"].get(
            f"{get_column_letter(cell_range.min_col)}{cell_range.min_row}"
        )
        if anchor is None:
            continue
        for row in range(cell_range.min_row, cell_range.max_row + 1):
            for col in range(cell_range.min_col, cell_range.max_col + 1):
                border_schema = _edge_border_schema(
                    anchor["borders"], cell_range, row, col
                )
                if border_schema is None:
                    continue
                coord = f"{get_column_letter(col)}{row}"
                edge_cell = dict(anchor)
                edge_cell["coordinate"] = coord
                edge_cell["value"] = (
                    anchor["value"] if coord == anchor["coordinate"] else None
                )
                edge_cell["cell_type"] = (
                    anchor["cell_type"] if coord == anchor["coordinate"] else "empty"
                )
                edge_cell["borders"] = border_schema
                result[(row, col)] = edge_cell  # type: ignore[assignment]
    return result


def _edge_border(
    borders: dict[str, dict[str, Any]], cell_range: CellRange, row: int, col: int
) -> Border | None:
    schema = _edge_border_schema(borders, cell_range, row, col)
    return _build_border(schema) if schema is not None else None


def _edge_border_schema(
    borders: dict[str, Any], cell_range: CellRange, row: int, col: int
) -> dict[str, Any] | None:
    return _edge_border_schema_bounds(
        borders,
        min_row=cell_range.min_row,
        max_row=cell_range.max_row,
        min_col=cell_range.min_col,
        max_col=cell_range.max_col,
        row=row,
        col=col,
    )


def _edge_border_schema_bounds(
    borders: dict[str, Any],
    *,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    row: int,
    col: int,
) -> dict[str, Any] | None:
    empty: dict[str, Any] = {"style": None, "color": None}
    # Only materialize the visible outline for synthesized merged cells.
    # Writing interior horizontal/vertical merge borders onto every cell can
    # make Excel hide left/right edges even though openpyxl round-trips them.
    result: dict[str, Any] = {
        "top": dict(empty),
        "bottom": dict(empty),
        "left": dict(empty),
        "right": dict(empty),
        "start": dict(empty),
        "end": dict(empty),
        "horizontal": dict(empty),
        "vertical": dict(empty),
        "diagonal": dict(borders.get("diagonal", empty)),
        "diagonal_up": borders.get("diagonal_up", False),
        "diagonal_down": borders.get("diagonal_down", False),
        "outline": borders.get("outline", True),
    }
    if row == min_row:
        result["top"] = dict(borders["top"])
    if row == max_row:
        result["bottom"] = dict(borders["bottom"])
    if col == min_col:
        result["left"] = dict(borders["left"])
        result["start"] = dict(borders.get("start", empty))
    if col == max_col:
        result["right"] = dict(borders["right"])
        result["end"] = dict(borders.get("end", empty))

    _side_keys = ("top", "bottom", "left", "right", "start", "end",
                  "horizontal", "vertical", "diagonal")
    has_style = any(result[k].get("style") for k in _side_keys)
    has_diagonal = result.get("diagonal_up") or result.get("diagonal_down")
    return result if (has_style or has_diagonal) else None


def _apply_streaming_merges(ws, schema: SheetSchema) -> None:
    for region in schema.get("merged_regions", []):
        ws.merged_cells.add(CellRange(region))


def _add_generated_merge(ws, merge_range: CellRange | str) -> None:
    # Renderer-generated occupation merges are validated by construction and are
    # one row tall, so avoid openpyxl's O(n) overlap scan for every streamed row.
    ws.merged_cells.ranges.add(merge_range)


def _apply_sheet_view(ws, schema: SheetSchema) -> None:
    ws.sheet_view.showGridLines = schema.get("show_gridlines", True)


def _apply_streaming_dimensions(ws, schema: SheetSchema) -> None:
    col_mode = schema.get("column_width_mode", "fixed")
    row_mode = schema.get("row_height_mode", "fixed")
    min_col, min_row, max_col, max_row = _parse_dims(schema["dimensions"])

    if col_mode == "fixed":
        for col_letter, width in schema["column_widths"].items():
            if width is not None:
                ws.column_dimensions[col_letter].width = width
    elif col_mode == "even":
        width = schema.get("default_column_width", 15.0)
        for col_idx in range(min_col, max_col + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = width

    if row_mode == "fixed":
        for row_str, height in schema["row_heights"].items():
            if height is not None:
                ws.row_dimensions[int(row_str)].height = height
    elif row_mode == "even":
        height = schema.get("default_row_height", 15.0)
        for row_idx in range(min_row, max_row + 1):
            ws.row_dimensions[row_idx].height = height


def _part_path(base_output_path: str, part: int) -> str:
    p = Path(base_output_path)
    return str(p.with_name(f"{p.stem}.part{part:03d}{p.suffix}"))


def _bundle_parts_if_needed(base_output_path: str, part_paths: list[str]) -> list[str]:
    if len(part_paths) <= 1:
        return part_paths

    output = Path(base_output_path)
    zip_path = output.with_name(f"{output.stem}.zip")
    with ZipFile(zip_path, mode="w", compression=ZIP_DEFLATED) as zip_file:
        for part_path in part_paths:
            part = Path(part_path)
            zip_file.write(part, arcname=part.name)

    for part_path in part_paths:
        Path(part_path).unlink()
    return [str(zip_path)]


def _delete_bundle_dir(bundle: ReportBundle) -> None:
    path = Path(bundle.path)
    if not path.is_dir():
        return
    if not (path / "manifest.json").exists() or not (path / "report.json").exists():
        raise ValueError(f"Refusing to delete malformed report bundle directory: {path}")
    shutil.rmtree(path)


# §4. Public Functions


def export_report_bundle(
    bundle_or_path: ReportBundle | str,
    output_path: str,
    *,
    format: str = "xlsx",
    export_mode: str = "fidelity",
    column_width_mode: str | None = None,
    row_height_mode: str | None = None,
    default_column_width: float | None = None,
    default_row_height: float | None = None,
    streaming_engine: str | None = None,
    streaming_chunk_rows: int = 50_000,
    max_rows_per_workbook: int = MAX_EXCEL_ROWS,
    auto_delete_bundle: bool = False,
    **options: Any,
) -> None | list[str]:
    """Render a ReportBundle to the requested output format."""
    if format == "image":
        raise NotImplementedError(
            "ReportBundle export format 'image' is reserved but not implemented in v1."
        )
    if format not in {"xlsx", "pdf"}:
        raise ValueError(
            f"Unsupported report export format '{format}'. Expected 'xlsx', 'pdf', or 'image'."
        )

    bundle = _coerce_bundle(bundle_or_path)
    result: None | list[str]
    if format == "pdf":
        if streaming_engine is not None:
            raise ValueError("PDF export does not support streaming_engine.")
        if export_mode != "fidelity":
            raise ValueError(
                f"PDF export does not support export_mode '{export_mode}'. PDF output paginates automatically."
            )
        from .pdf_renderer import export_report_bundle as _export_pdf_report_bundle

        _export_pdf_report_bundle(
            bundle,
            output_path,
            column_width_mode=column_width_mode,
            row_height_mode=row_height_mode,
            default_column_width=default_column_width,
            default_row_height=default_row_height,
            page_size=options.get("page_size", "A4"),
            orientation=options.get("orientation", "portrait"),
            margin=options.get("margin", 36),
            streaming_chunk_rows=streaming_chunk_rows,
            fonts=options.get("fonts"),
            repeat_dataframe_headers=options.get("repeat_dataframe_headers", False),
        )
        result = None
    elif export_mode == "streaming":
        engine = streaming_engine or "openpyxl"
        if engine not in {"openpyxl", "xlsxwriter"}:
            raise ValueError(
                f"Unsupported streaming_engine '{streaming_engine}'. Expected 'openpyxl' or 'xlsxwriter'."
            )
        result = _render_streaming(
            bundle,
            output_path,
            streaming_engine=engine,
            column_width_mode=column_width_mode,
            row_height_mode=row_height_mode,
            default_column_width=default_column_width,
            default_row_height=default_row_height,
            streaming_chunk_rows=streaming_chunk_rows,
            max_rows_per_workbook=max_rows_per_workbook,
        )
    elif export_mode == "fidelity":
        if streaming_engine == "xlsxwriter":
            raise ValueError(
                "Fidelity XLSX export does not support streaming_engine='xlsxwriter'. Use export_mode='streaming'."
            )
        if streaming_engine not in {None, "openpyxl"}:
            raise ValueError(
                f"Unsupported streaming_engine '{streaming_engine}'. Expected 'openpyxl' or 'xlsxwriter'."
            )
        _render_fidelity(
            bundle,
            output_path,
            column_width_mode=column_width_mode,
            row_height_mode=row_height_mode,
            default_column_width=default_column_width,
            default_row_height=default_row_height,
        )
        result = None
    else:
        raise ValueError(
            f"Unsupported export_mode '{export_mode}'. Expected 'fidelity' or 'streaming'."
        )

    if auto_delete_bundle:
        _delete_bundle_dir(bundle)
    return result


# §5. Entrypoints
