from __future__ import annotations

import datetime
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
        for row in batch.to_pylist():
            yield tuple(row.get(col) for col in columns)


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
        merge_range = _layout_merge_range(anchor, layout, row_idx)
        if merge_range is None:
            continue
        merges.append(
            {
                "min_row_offset": 0,
                "max_row_offset": 0,
                "min_col": merge_range.min_col,
                "max_col": merge_range.max_col,
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
        "static": static,
        "generated_merges": generated_merges,
        "content": content,
        "repeat_rows": repeat_rows,
        "min_col": min_col,
        "min_row": min_row,
        "max_col": max_col,
        "max_row": max_row,
    }


def _repeat_max_col(sheet: SheetSchema, current: int) -> int:
    max_col = current
    for section in sheet.get("repeat_sections", []):
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
        for record in section["records"]:
            for row_item in _repeat_record_rows(
                bundle,
                source_map,
                record,
                block_height=record.get("block_height", section["block_height"]),
                merges=record.get("merged_regions", section.get("merged_regions", [])),
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
    batch_size: int,
) -> Iterator[dict[str, Any]]:
    cells_by_offset: dict[int, dict[int, CellSchema]] = {}
    for item in record["cells"]:
        cell = item["cell"]
        cells_by_offset.setdefault(item["row_offset"], {})[item["start_col"]] = cell
    _apply_repeat_merge_edge_cells(cells_by_offset, merges or [])

    headers_by_offset: dict[int, list[dict[str, Any]]] = {}
    content_by_offset: dict[int, list[dict[str, Any]]] = {}
    for anchor in record["dataframe_anchors"]:
        target = (
            headers_by_offset
            if anchor["placeholder_type"] == "dataframe-header"
            else content_by_offset
        )
        target.setdefault(anchor["start_row_offset"], []).append(anchor)

    offset = 0
    while offset < block_height:
        row_cells = dict(cells_by_offset.get(offset, {}))
        header_merges: list[dict[str, Any]] = []
        for anchor in headers_by_offset.get(offset, []):
            for cell in _header_cells_at(anchor, offset + 1):
                _, col_idx = _coord_indexes(cell["coordinate"])
                row_cells[col_idx] = cell
            header_merges.extend(_anchor_row_merges(anchor, offset + 1))

        content_anchors = content_by_offset.get(offset, [])
        if not content_anchors:
            yield {
                "cells": row_cells,
                "merges": _repeat_merges_starting_at(merges or [], offset)
                + header_merges,
            }
            offset += 1
            continue
        yield from _repeat_content_rows(
            bundle,
            source_map,
            row_cells,
            content_anchors,
            base_merges=header_merges,
            batch_size=batch_size,
        )
        anchor = content_anchors[0]
        offset += max(int(anchor.get("source_rows") or 0), 1)


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
    for row_values in _source_rows(bundle, source, batch_size=batch_size):
        row_cells = dict(base_cells) if not wrote else {}
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
        yield {"cells": row_cells, "merges": (base_merges or []) + _anchor_row_merges(anchor, 1)}
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
                row_cells[col_idx - plan["min_col"]] = _write_only_cell(ws, static)
        for merge_range in plan.get("generated_merges", {}).get(row_idx, []):
            ws.merged_cells.add(merge_range)
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
                ws.merged_cells.add(
                    CellRange(
                        min_col=merge["min_col"],
                        min_row=row_idx,
                        max_col=merge["max_col"],
                        max_row=max_merge_row,
                    )
                )
        row_cells = [None] * col_count
        for col_idx, schema in row_map.items():
            if col_idx < plan["min_col"] or col_idx > max_col:
                continue
            row_cells[col_idx - plan["min_col"]] = _write_only_cell(ws, schema)
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
            )
        merge_range = _layout_merge_range(anchor, layout, row_idx)
        if merge_range is not None:
            ws.merged_cells.add(merge_range)
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

    merge_range = _layout_merge_range(anchor, layout, row_idx)
    if merge_range is None:
        first = dict(schema)
        first["value"] = value
        first["cell_type"] = _infer_cell_type(value)
        return {start_col: first}  # type: ignore[return-value]

    result: dict[int, CellSchema] = {}
    end_col = start_col + occupation - 1
    for col_idx in range(start_col, end_col + 1):
        item = dict(schema)
        border_schema = _edge_border_schema(item["borders"], merge_range, row_idx, col_idx)
        if border_schema is not None:
            item["borders"] = border_schema
        if col_idx == start_col:
            item["value"] = value
            item["cell_type"] = _infer_cell_type(value)
        else:
            item["value"] = None
            item["cell_type"] = "empty"
        result[col_idx] = item  # type: ignore[assignment]
    return result


def _cell_with_layout(schema: CellSchema, layout: dict[str, Any]) -> CellSchema:
    result = dict(schema)
    result["alignment"] = _layout_alignment(schema, layout)
    return result  # type: ignore[return-value]


def _styled_write_only_cell(ws, schema: CellSchema, value: Any) -> WriteOnlyCell:
    cell = WriteOnlyCell(ws, value=value)
    cell.font = _build_font(schema["font"])
    cell.fill = _build_fill(schema["fill"])
    cell.alignment = _build_alignment(schema["alignment"])
    cell.border = _build_border(schema["borders"])
    if schema["number_format"]:
        cell.number_format = schema["number_format"]
    return cell


def _write_only_cell(ws, schema: CellSchema) -> WriteOnlyCell:
    value = schema["value"]
    if schema["cell_type"] == "date" and isinstance(value, str):
        value = datetime.datetime.fromisoformat(value)
    cell = _styled_write_only_cell(ws, schema, value)
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
    if row == cell_range.min_row:
        result["top"] = dict(borders["top"])
    if row == cell_range.max_row:
        result["bottom"] = dict(borders["bottom"])
    if col == cell_range.min_col:
        result["left"] = dict(borders["left"])
        result["start"] = dict(borders.get("start", empty))
    if col == cell_range.max_col:
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
        )
        result = None
    elif export_mode == "streaming":
        result = _render_streaming(
            bundle,
            output_path,
            column_width_mode=column_width_mode,
            row_height_mode=row_height_mode,
            default_column_width=default_column_width,
            default_row_height=default_row_height,
            streaming_chunk_rows=streaming_chunk_rows,
            max_rows_per_workbook=max_rows_per_workbook,
        )
    elif export_mode == "fidelity":
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
