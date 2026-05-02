from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from openpyxl.worksheet.pagebreak import Break, ColBreak, RowBreak

from .template_contract import _parse_dims

__all__ = [
    "apply_manual_breaks",
    "extract_manual_breaks",
    "normalize_breaks",
    "resolve_compiled_sheet_page_breaks",
]

# §1. Constants & Exceptions

# §2. Classes and Sub Classes

# §3. Private Helper Functions


def _anchor_width(anchor: dict[str, Any]) -> int:
    layouts = anchor.get("column_layouts", [])
    if not layouts:
        return 0
    last = layouts[-1]
    return int(last["start_col_offset"]) + int(last["occupation"])


def _grouped_sheet_footprints(anchors: list[dict[str, Any]]) -> list[dict[str, int]]:
    grouped: dict[tuple[str, str, int], dict[str, int]] = {}
    for anchor in anchors:
        width = _anchor_width(anchor)
        row_count = int(anchor.get("source_rows") or 0)
        if anchor["placeholder_type"] != "dataframe-content":
            row_count = 1
        if width <= 0 or row_count <= 0:
            continue
        key = (str(anchor["coordinate"]), str(anchor["key"]), int(anchor["start_col"]))
        item = grouped.get(key)
        max_col = int(anchor["start_col"]) + width - 1
        max_row = int(anchor["start_row"]) + row_count - 1
        if item is None:
            grouped[key] = {
                "start_row": int(anchor["start_row"]),
                "start_col": int(anchor["start_col"]),
                "max_row": max_row,
                "max_col": max_col,
                "row_delta": max_row - int(anchor["start_row"]),
                "col_delta": max_col - int(anchor["start_col"]),
            }
            continue
        item["start_row"] = min(item["start_row"], int(anchor["start_row"]))
        item["max_row"] = max(item["max_row"], max_row)
        item["max_col"] = max(item["max_col"], max_col)
        item["row_delta"] = item["max_row"] - item["start_row"]
        item["col_delta"] = max(item["col_delta"], max_col - item["start_col"])
    return list(grouped.values())


def _grouped_record_footprints(anchors: list[dict[str, Any]]) -> list[dict[str, int]]:
    grouped: dict[tuple[str, str, int], dict[str, int]] = {}
    for anchor in anchors:
        width = _anchor_width(anchor)
        row_count = int(anchor.get("source_rows") or 0)
        if anchor["placeholder_type"] != "dataframe-content":
            row_count = 1
        if width <= 0 or row_count <= 0:
            continue
        start_row = int(anchor["start_row_offset"]) + 1
        start_col = int(anchor["start_col"])
        key = (str(anchor["coordinate"]), str(anchor["key"]), start_col)
        max_col = start_col + width - 1
        max_row = start_row + row_count - 1
        item = grouped.get(key)
        if item is None:
            grouped[key] = {
                "start_row": start_row,
                "start_col": start_col,
                "max_row": max_row,
                "max_col": max_col,
                "row_delta": max_row - start_row,
                "col_delta": max_col - start_col,
            }
            continue
        item["start_row"] = min(item["start_row"], start_row)
        item["max_row"] = max(item["max_row"], max_row)
        item["max_col"] = max(item["max_col"], max_col)
        item["row_delta"] = item["max_row"] - item["start_row"]
        item["col_delta"] = max(item["col_delta"], max_col - item["start_col"])
    return list(grouped.values())


def _column_shift_groups(footprints: list[dict[str, int]]) -> list[dict[str, int]]:
    grouped: dict[int, int] = {}
    for footprint in footprints:
        start_col = int(footprint["start_col"])
        grouped[start_col] = max(grouped.get(start_col, 0), int(footprint["col_delta"]))
    return [
        {"start_col": start_col, "col_delta": col_delta}
        for start_col, col_delta in sorted(grouped.items())
    ]


def _resolve_row_breaks(
    row_breaks: list[int],
    footprints: list[dict[str, int]],
    content_start_rows: set[int],
) -> list[int]:
    resolved: list[int] = []
    for break_idx in row_breaks:
        shift = 0
        for footprint in footprints:
            start_row = int(footprint["start_row"])
            if break_idx > start_row or (
                break_idx == start_row and start_row in content_start_rows
            ):
                shift += int(footprint["row_delta"])
        resolved.append(break_idx + shift)
    return normalize_breaks(resolved)


def _resolve_column_breaks(
    column_breaks: list[int], footprints: list[dict[str, int]]
) -> list[int]:
    resolved: list[int] = []
    groups = _column_shift_groups(footprints)
    for break_idx in column_breaks:
        shift = sum(
            int(group["col_delta"])
            for group in groups
            if break_idx >= int(group["start_col"])
        )
        resolved.append(break_idx + shift)
    return normalize_breaks(resolved)


def _resolve_repeat_record_row_breaks(
    record: dict[str, Any],
    local_breaks: list[int],
    output_start_row: int,
) -> list[int]:
    content_start_rows = {
        int(anchor["start_row_offset"]) + 1
        for anchor in record.get("dataframe_anchors", [])
        if anchor["placeholder_type"] == "dataframe-content"
    }
    resolved_local = _resolve_row_breaks(
        local_breaks,
        _grouped_record_footprints(record.get("dataframe_anchors", [])),
        content_start_rows,
    )
    return [output_start_row + local_break - 1 for local_break in resolved_local]


def _resolve_repeat_row_breaks(sheet: dict[str, Any], row_breaks: list[int]) -> list[int]:
    _, min_row, _, max_row = _parse_dims(sheet["dimensions"])
    resolved: list[int] = []
    cursor_template = min_row
    cursor_output = min_row

    for section in sheet.get("repeat_sections", []):
        for break_idx in row_breaks:
            if cursor_template <= break_idx < int(section["start_row"]):
                resolved.append(cursor_output + (break_idx - cursor_template))
        cursor_output += max(int(section["start_row"]) - cursor_template, 0)

        block_start = int(section["template_start_row"])
        block_end = int(section["template_end_row"])
        local_breaks = [
            break_idx - block_start + 1
            for break_idx in row_breaks
            if block_start <= break_idx <= block_end
        ]
        for record in section["records"]:
            resolved.extend(
                _resolve_repeat_record_row_breaks(record, local_breaks, cursor_output)
            )
            cursor_output += int(record.get("block_height", section["block_height"]))
        cursor_template = int(section["end_row"]) + 1

    for break_idx in row_breaks:
        if cursor_template <= break_idx <= max_row:
            resolved.append(cursor_output + (break_idx - cursor_template))
    return normalize_breaks(resolved)


def _resolve_repeat_column_breaks(
    sheet: dict[str, Any], column_breaks: list[int]
) -> list[int]:
    footprints: list[dict[str, int]] = []
    for section in sheet.get("repeat_sections", []):
        widest_by_start_col: dict[int, dict[str, int]] = {}
        for record in section["records"]:
            for footprint in _grouped_record_footprints(record.get("dataframe_anchors", [])):
                start_col = int(footprint["start_col"])
                existing = widest_by_start_col.get(start_col)
                if existing is None or int(footprint["col_delta"]) > int(existing["col_delta"]):
                    widest_by_start_col[start_col] = footprint
        footprints.extend(widest_by_start_col.values())
    return _resolve_column_breaks(column_breaks, footprints)


def _repeat_rendered_max_row(sheet: dict[str, Any]) -> int:
    _, min_row, _, max_row = _parse_dims(sheet["dimensions"])
    cursor_template = min_row
    cursor_output = min_row
    for section in sheet.get("repeat_sections", []):
        cursor_output += max(int(section["start_row"]) - cursor_template, 0)
        for record in section["records"]:
            cursor_output += int(record.get("block_height", section["block_height"]))
        cursor_template = int(section["end_row"]) + 1
    cursor_output += max(max_row - cursor_template + 1, 0)
    return max(cursor_output - 1, min_row)


def _sheet_max_row(sheet: dict[str, Any]) -> int:
    if sheet.get("repeat_sections"):
        return _repeat_rendered_max_row(sheet)
    return _parse_dims(sheet["dimensions"])[3]


# §4. Public Functions


def normalize_breaks(values: Iterable[Any] | None) -> list[int]:
    normalized: set[int] = set()
    if values is None:
        return []
    for value in values:
        try:
            break_idx = int(value)
        except (TypeError, ValueError):
            continue
        if break_idx <= 0:
            continue
        normalized.add(break_idx)
    return sorted(normalized)


def extract_manual_breaks(page_breaks: Any) -> list[int]:
    breaks: list[int] = []
    for page_break in getattr(page_breaks, "brk", ()):
        if not bool(getattr(page_break, "man", False)):
            continue
        try:
            break_idx = int(getattr(page_break, "id"))
        except (TypeError, ValueError):
            continue
        if break_idx <= 0:
            continue
        breaks.append(break_idx)
    return normalize_breaks(breaks)


def resolve_compiled_sheet_page_breaks(sheet: dict[str, Any]) -> None:
    row_breaks = normalize_breaks(sheet.get("row_page_breaks"))
    column_breaks = normalize_breaks(sheet.get("column_page_breaks"))
    sheet["row_page_breaks"] = row_breaks
    sheet["column_page_breaks"] = column_breaks

    if sheet.get("repeat_sections"):
        resolved_rows = _resolve_repeat_row_breaks(sheet, row_breaks)
        resolved_cols = _resolve_repeat_column_breaks(sheet, column_breaks)
    else:
        footprints = _grouped_sheet_footprints(sheet.get("dataframe_anchors", []))
        content_start_rows = {
            int(anchor["start_row"])
            for anchor in sheet.get("dataframe_anchors", [])
            if anchor["placeholder_type"] == "dataframe-content"
        }
        resolved_rows = _resolve_row_breaks(row_breaks, footprints, content_start_rows)
        resolved_cols = _resolve_column_breaks(column_breaks, footprints)

    min_col, min_row, max_col, _ = _parse_dims(sheet["dimensions"])
    max_row = _sheet_max_row(sheet)
    sheet["resolved_row_page_breaks"] = [
        break_idx for break_idx in resolved_rows if min_row <= break_idx < max_row
    ]
    sheet["resolved_column_page_breaks"] = [
        break_idx
        for break_idx in resolved_cols
        if min_col <= break_idx < max_col
    ]


def apply_manual_breaks(
    ws: Any, *, row_breaks: Iterable[Any] | None, column_breaks: Iterable[Any] | None
) -> None:
    ws.row_breaks = RowBreak()
    ws.col_breaks = ColBreak()
    for break_idx in normalize_breaks(row_breaks):
        ws.row_breaks.append(Break(id=break_idx))
    for break_idx in normalize_breaks(column_breaks):
        ws.col_breaks.append(Break(id=break_idx))


# §5. Entrypoints
