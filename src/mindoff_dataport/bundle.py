from __future__ import annotations

import datetime
import json
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pyarrow.parquet as pq
from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)

from .template_contract import (
    PLACEHOLDER_RE,
    _DATAFRAME_TYPES,
    _SCALAR_TYPES,
    _infer_cell_type,
    _parse_dims,
    _repeat_sections,
    _resolve_sheet_payloads,
    _substitute_scalars,
    _to_headers,
    get_template_inputs,
)
from .schema import CellSchema, SheetSchema, WorkbookSchema

__all__ = [
    "ReportBundle",
    "compile_report_bundle",
    "load_report_bundle",
]

# §1. Constants & Exceptions

BUNDLE_VERSION = "1.0"

# §2. Classes and Sub Classes


@dataclass(frozen=True)
class ReportBundle:
    """Canonical intermediate report artifact, backed by a directory."""

    manifest: dict[str, Any]
    report: dict[str, Any]
    path: str

    def write(self, bundle_path: str) -> None:
        _copy_bundle_dir(self, bundle_path)

    @classmethod
    def load(cls, bundle_path: str) -> "ReportBundle":
        return load_report_bundle(bundle_path)


# §3. Private Helper Functions


def _copy_bundle_dir(bundle: ReportBundle, bundle_path: str) -> None:
    source = Path(bundle.path).resolve()
    target = Path(bundle_path).resolve()
    if source == target:
        _write_bundle_metadata(bundle)
        return
    if target.exists() and not target.is_dir():
        raise ValueError(f"Report bundle path must be a directory: {bundle_path}")
    shutil.copytree(source, target, dirs_exist_ok=True)


def _write_bundle_metadata(bundle: ReportBundle) -> None:
    path = Path(bundle.path)
    path.mkdir(parents=True, exist_ok=True)
    (path / "manifest.json").write_text(
        json.dumps(bundle.manifest, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    (path / "report.json").write_text(
        json.dumps(bundle.report, indent=2, sort_keys=True),
        encoding="utf-8",
    )


def _prepare_bundle_dir(bundle_path: str | None) -> Path:
    if bundle_path is None:
        path = Path(tempfile.mkdtemp(prefix="mindoff_report_bundle_"))
    else:
        path = Path(bundle_path)
    if path.exists() and not path.is_dir():
        raise ValueError(f"Report bundle path must be a directory: {path}")
    path.mkdir(parents=True, exist_ok=True)
    (path / "data").mkdir(exist_ok=True)
    return path


def _parquet_metadata(path: Path) -> tuple[list[str], int]:
    parquet_file = pq.ParquetFile(path)
    columns = [str(name) for name in parquet_file.schema_arrow.names]
    return columns, int(parquet_file.metadata.num_rows)


def _write_source_file(
    *,
    value: Any,
    source_id: str,
    bundle_dir: Path,
    source_cache: dict[int, dict[str, Any]],
) -> dict[str, Any]:
    cache_key = id(value)
    cached = source_cache.get(cache_key)
    if cached is not None:
        return cached

    record = _write_dataframe_source(
        value=value,
        source_id=source_id,
        bundle_dir=bundle_dir,
    )
    source_cache[cache_key] = record
    return record


def _write_dataframe_source(
    *, value: Any, source_id: str, bundle_dir: Path
) -> dict[str, Any]:
    rel_path = f"data/{source_id}.parquet"
    output_path = bundle_dir / rel_path
    module = getattr(type(value), "__module__", "") or ""
    qualname = type(value).__qualname__

    if "polars" in module and qualname == "LazyFrame":
        value.sink_parquet(output_path)
    elif "polars" in module and qualname == "DataFrame":
        value.write_parquet(output_path)
    else:
        raise TypeError(
            f"Expected a polars DataFrame or LazyFrame, got {type(value).__name__}"
        )

    columns, row_count = _parquet_metadata(output_path)
    record = {
        "id": source_id,
        "path": rel_path,
        "format": "parquet",
        "columns": columns,
        "rows": row_count,
        "file_backed": True,
    }
    return record


def _schema_value(value: Any) -> Any:
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.isoformat()
    return value


def _unique_id(used: set[str], raw: str) -> str:
    safe = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in raw)
    safe = safe.strip("._-") or "source"
    candidate = safe
    idx = 2
    while candidate in used:
        candidate = f"{safe}_{idx}"
        idx += 1
    used.add(candidate)
    return candidate


def _coord_indexes(coord: str) -> tuple[int, int]:
    col_letter, row_idx = coordinate_from_string(coord)
    return row_idx, column_index_from_string(col_letter)


def _compile_sheet(
    *,
    sheet: SheetSchema,
    output_name: str,
    sheet_data: dict[str, Any],
    bundle_dir: Path,
    data_sources: list[dict[str, Any]],
    used_ids: set[str],
    source_cache: dict[int, dict[str, Any]],
) -> dict[str, Any]:
    repeats = _repeat_sections(sheet)
    if repeats:
        return _compile_repeat_sheet(
            sheet=sheet,
            repeats=repeats,
            output_name=output_name,
            sheet_data=sheet_data,
            bundle_dir=bundle_dir,
            data_sources=data_sources,
            used_ids=used_ids,
            source_cache=source_cache,
        )

    min_col, min_row, max_col, max_row = _parse_dims(sheet["dimensions"])
    static_cells: dict[str, CellSchema] = {}
    anchors: list[dict[str, Any]] = []

    for coord, cell in sheet["cells"].items():
        value = cell.get("value")
        if not isinstance(value, str):
            static_cells[coord] = cell
            continue

        full_match = PLACEHOLDER_RE.fullmatch(value.strip())
        if (
            full_match
            and full_match.group(2) in _DATAFRAME_TYPES
            and full_match.group(1) in sheet_data
        ):
            key = full_match.group(1)
            anchor_type = full_match.group(2)
            columns = _to_headers(sheet_data[key])
            col_letter, row_idx = coordinate_from_string(coord)
            col_idx = column_index_from_string(col_letter)
            max_col = max(max_col, col_idx + max(len(columns) - 1, 0))

            if anchor_type == "dataframe-header":
                anchor_id = _unique_id(used_ids, f"{output_name}__{key}__{anchor_type}")
                anchors.append(
                    _dataframe_anchor(
                        anchor_id=anchor_id,
                        key=key,
                        placeholder_type=anchor_type,
                        coord=coord,
                        row_idx=row_idx,
                        col_idx=col_idx,
                        columns=columns,
                        source_record=None,
                        cell=cell,
                    )
                )
                continue

            source_id_type = (
                "dataframe-content" if anchor_type == "dataframe" else anchor_type
            )
            anchor_id = _unique_id(used_ids, f"{output_name}__{key}__{source_id_type}")
            source_record = _write_source_file(
                value=sheet_data[key],
                source_id=anchor_id,
                bundle_dir=bundle_dir,
                source_cache=source_cache,
            )
            if source_record not in data_sources:
                data_sources.append(source_record)
            if anchor_type == "dataframe":
                header_id = _unique_id(
                    used_ids, f"{output_name}__{key}__dataframe-header"
                )
                anchors.append(
                    _dataframe_anchor(
                        anchor_id=header_id,
                        key=key,
                        placeholder_type="dataframe-header",
                        coord=coord,
                        row_idx=row_idx,
                        col_idx=col_idx,
                        columns=columns,
                        source_record=None,
                        cell=cell,
                    )
                )
                row_idx += 1
            anchors.append(
                _dataframe_anchor(
                    anchor_id=anchor_id,
                    key=key,
                    placeholder_type="dataframe-content",
                    coord=coord,
                    row_idx=row_idx,
                    col_idx=col_idx,
                    columns=columns,
                    source_record=source_record,
                    cell=cell,
                )
            )
            continue

        new_value = _substitute_scalars(value, sheet_data)
        new_cell: dict[str, Any] = dict(cell)
        new_cell["value"] = _schema_value(new_value)
        if full_match and full_match.group(2) in _SCALAR_TYPES:
            new_cell["cell_type"] = _infer_cell_type(new_value)
        static_cells[coord] = new_cell  # type: ignore[assignment]

    dimensions = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    result = dict(sheet)
    result["name"] = output_name
    result["dimensions"] = dimensions
    result["cells"] = static_cells
    result["dataframe_anchors"] = anchors
    return result


def _compile_repeat_sheet(
    *,
    sheet: SheetSchema,
    repeats: list[dict[str, Any]],
    output_name: str,
    sheet_data: dict[str, Any],
    bundle_dir: Path,
    data_sources: list[dict[str, Any]],
    used_ids: set[str],
    source_cache: dict[int, dict[str, Any]],
) -> dict[str, Any]:
    for repeat in repeats:
        _validate_repeat_layout(sheet, repeat)
    min_col, min_row, max_col, _ = _parse_dims(sheet["dimensions"])
    static_cells: dict[str, CellSchema] = {}
    repeat_sections: list[dict[str, Any]] = []
    repeat_ranges = [(repeat["start_row"], repeat["end_row"]) for repeat in repeats]

    for coord, cell in sheet["cells"].items():
        row_idx, _ = _coord_indexes(coord)
        if _row_in_ranges(row_idx, repeat_ranges):
            continue
        if _is_non_anchor_merged_cell(cell):
            continue
        compiled = _compile_cell(
            cell=cell,
            payload=sheet_data,
            output_name=output_name,
            bundle_dir=bundle_dir,
            data_sources=data_sources,
            used_ids=used_ids,
            source_cache=source_cache,
        )
        static_cells.update(compiled["cells"])

    for repeat in repeats:
        section, section_max_col = _compile_repeat_section(
            sheet=sheet,
            repeat=repeat,
            output_name=output_name,
            sheet_data=sheet_data,
            bundle_dir=bundle_dir,
            data_sources=data_sources,
            used_ids=used_ids,
            source_cache=source_cache,
        )
        repeat_sections.append(section)
        max_col = max(max_col, section_max_col)

    result = dict(sheet)
    result["name"] = output_name
    result["dimensions"] = (
        f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{_last_static_row(sheet, repeats)}"
    )
    result["merged_regions"] = _static_merged_regions(sheet, repeats)
    result["cells"] = static_cells
    result["dataframe_anchors"] = []
    result["repeat_sections"] = repeat_sections
    return result


def _compile_repeat_section(
    *,
    sheet: SheetSchema,
    repeat: dict[str, Any],
    output_name: str,
    sheet_data: dict[str, Any],
    bundle_dir: Path,
    data_sources: list[dict[str, Any]],
    used_ids: set[str],
    source_cache: dict[int, dict[str, Any]],
) -> tuple[dict[str, Any], int]:
    repeat_key = repeat["key"]
    records = sheet_data[repeat_key]
    block_start = repeat["start_row"] + 1
    block_end = repeat["end_row"] - 1
    repeat_merges = _repeat_merged_regions(sheet, repeat, block_start, block_end)
    compiled_records: list[dict[str, Any]] = []
    block_height = max(block_end - block_start + 1, 0)
    max_col = _parse_dims(sheet["dimensions"])[2]
    for index, record in enumerate(records):
        record_cells: list[dict[str, Any]] = []
        record_anchors: list[dict[str, Any]] = []
        for coord, cell in sheet["cells"].items():
            row_idx, col_idx = _coord_indexes(coord)
            if row_idx < block_start or row_idx > block_end:
                continue
            if _is_non_anchor_merged_cell(cell):
                continue
            compiled = _compile_cell(
                cell=cell,
                payload=record,
                output_name=f"{output_name}__{repeat_key}_{index + 1}",
                bundle_dir=bundle_dir,
                data_sources=data_sources,
                used_ids=used_ids,
                source_cache=source_cache,
            )
            row_offset = row_idx - block_start
            for item in compiled["cells"].values():
                record_cells.append(
                    {
                        "row_offset": row_offset,
                        "start_col": col_idx,
                        "cell": item,
                    }
                )
            for anchor in compiled["anchors"]:
                anchor = dict(anchor)
                anchor["start_row_offset"] = anchor.pop("start_row") - block_start
                record_anchors.append(anchor)
                max_col = max(
                    max_col,
                    anchor["start_col"] + max(len(anchor["columns"]) - 1, 0),
                )
        compiled_records.append(
            {
                "index": index,
                "cells": record_cells,
                "dataframe_anchors": record_anchors,
            }
        )
    return (
        {
            "key": repeat_key,
            "start_row": repeat["start_row"],
            "end_row": repeat["end_row"],
            "template_start_row": block_start,
            "template_end_row": block_end,
            "block_height": block_height,
            "merged_regions": repeat_merges,
            "records": compiled_records,
        },
        max_col,
    )


def _validate_repeat_layout(sheet: SheetSchema, repeat: dict[str, Any]) -> None:
    block_start = repeat["start_row"] + 1
    block_end = repeat["end_row"] - 1
    content_rows = _repeat_content_rows(sheet, block_start, block_end)
    for region in sheet.get("merged_regions", []):
        start, end = region.split(":")
        start_row, _ = _coord_indexes(start)
        end_row, _ = _coord_indexes(end)
        if start_row < block_start or end_row > block_end:
            continue
        if any(start_row <= row_idx <= end_row for row_idx in content_rows):
            raise ValueError(
                f"Repeat section '{repeat['key']}' does not support merged cells that overlap dataframe-content rows"
            )


def _row_in_ranges(row_idx: int, ranges: list[tuple[int, int]]) -> bool:
    return any(start <= row_idx <= end for start, end in ranges)


def _is_non_anchor_merged_cell(cell: CellSchema) -> bool:
    return bool(cell.get("merged")) and cell.get("merge_anchor") not in (
        None,
        cell["coordinate"],
    )


def _repeat_content_rows(
    sheet: SheetSchema, block_start: int, block_end: int
) -> set[int]:
    rows: set[int] = set()
    for cell in sheet["cells"].values():
        row_idx, _ = _coord_indexes(cell["coordinate"])
        if row_idx < block_start or row_idx > block_end:
            continue
        value = cell.get("value")
        if not isinstance(value, str):
            continue
        match = PLACEHOLDER_RE.fullmatch(value.strip())
        if not match:
            continue
        if match.group(2) == "dataframe-content":
            rows.add(row_idx)
        elif match.group(2) == "dataframe":
            rows.add(row_idx + 1)
    return rows


def _static_merged_regions(sheet: SheetSchema, repeats: list[dict[str, Any]]) -> list[str]:
    regions: list[str] = []
    repeat_ranges = [(repeat["start_row"], repeat["end_row"]) for repeat in repeats]
    for region in sheet.get("merged_regions", []):
        start, end = region.split(":")
        start_row, _ = _coord_indexes(start)
        end_row, _ = _coord_indexes(end)
        if not any(start_row <= repeat_end and end_row >= repeat_start for repeat_start, repeat_end in repeat_ranges):
            regions.append(region)
    return regions


def _last_static_row(sheet: SheetSchema, repeats: list[dict[str, Any]]) -> int:
    _, min_row, _, max_row = _parse_dims(sheet["dimensions"])
    repeat_ranges = [(repeat["start_row"], repeat["end_row"]) for repeat in repeats]
    last = min_row
    for cell in sheet["cells"].values():
        row_idx, _ = _coord_indexes(cell["coordinate"])
        value = cell.get("value")
        if _row_in_ranges(row_idx, repeat_ranges):
            continue
        if value not in (None, ""):
            last = max(last, row_idx)
    return max(last, max(repeat["start_row"] for repeat in repeats), min(max_row, last))


def _repeat_merged_regions(
    sheet: SheetSchema, repeat: dict[str, Any], block_start: int, block_end: int
) -> list[dict[str, Any]]:
    regions: list[dict[str, Any]] = []
    for region in sheet.get("merged_regions", []):
        start, end = region.split(":")
        start_row, start_col = _coord_indexes(start)
        end_row, end_col = _coord_indexes(end)
        if start_row < block_start or end_row > block_end:
            continue
        regions.append(
            {
                "min_row_offset": start_row - block_start,
                "max_row_offset": end_row - block_start,
                "min_col": start_col,
                "max_col": end_col,
            }
        )
    return regions


def _compile_cell(
    *,
    cell: CellSchema,
    payload: dict[str, Any],
    output_name: str,
    bundle_dir: Path,
    data_sources: list[dict[str, Any]],
    used_ids: set[str],
    source_cache: dict[int, dict[str, Any]],
) -> dict[str, Any]:
    value = cell.get("value")
    if not isinstance(value, str):
        return {"cells": {cell["coordinate"]: cell}, "anchors": []}

    full_match = PLACEHOLDER_RE.fullmatch(value.strip())
    if (
        full_match
        and full_match.group(2) in _DATAFRAME_TYPES
        and full_match.group(1) in payload
    ):
        key = full_match.group(1)
        anchor_type = full_match.group(2)
        row_idx, col_idx = _coord_indexes(cell["coordinate"])
        return {
            "cells": {},
            "anchors": _compile_dataframe_anchors(
                key=key,
                anchor_type=anchor_type,
                cell=cell,
                row_idx=row_idx,
                col_idx=col_idx,
                value=payload[key],
                output_name=output_name,
                bundle_dir=bundle_dir,
                data_sources=data_sources,
                used_ids=used_ids,
                source_cache=source_cache,
            ),
        }

    new_value = _substitute_scalars(value, payload)
    new_cell: dict[str, Any] = dict(cell)
    new_cell["value"] = _schema_value(new_value)
    if full_match and full_match.group(2) in _SCALAR_TYPES:
        new_cell["cell_type"] = _infer_cell_type(new_value)
    return {"cells": {cell["coordinate"]: new_cell}, "anchors": []}


def _compile_dataframe_anchors(
    *,
    key: str,
    anchor_type: str,
    cell: CellSchema,
    row_idx: int,
    col_idx: int,
    value: Any,
    output_name: str,
    bundle_dir: Path,
    data_sources: list[dict[str, Any]],
    used_ids: set[str],
    source_cache: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    columns = _to_headers(value)
    if anchor_type == "dataframe-header":
        anchor_id = _unique_id(used_ids, f"{output_name}__{key}__{anchor_type}")
        return [
            _dataframe_anchor(
                anchor_id=anchor_id,
                key=key,
                placeholder_type=anchor_type,
                coord=cell["coordinate"],
                row_idx=row_idx,
                col_idx=col_idx,
                columns=columns,
                source_record=None,
                cell=cell,
            )
        ]

    source_id_type = "dataframe-content" if anchor_type == "dataframe" else anchor_type
    anchor_id = _unique_id(used_ids, f"{output_name}__{key}__{source_id_type}")
    source_record = _write_source_file(
        value=value,
        source_id=anchor_id,
        bundle_dir=bundle_dir,
        source_cache=source_cache,
    )
    if source_record not in data_sources:
        data_sources.append(source_record)

    anchors: list[dict[str, Any]] = []
    if anchor_type == "dataframe":
        header_id = _unique_id(used_ids, f"{output_name}__{key}__dataframe-header")
        anchors.append(
            _dataframe_anchor(
                anchor_id=header_id,
                key=key,
                placeholder_type="dataframe-header",
                coord=cell["coordinate"],
                row_idx=row_idx,
                col_idx=col_idx,
                columns=columns,
                source_record=None,
                cell=cell,
            )
        )
        row_idx += 1
    anchors.append(
        _dataframe_anchor(
            anchor_id=anchor_id,
            key=key,
            placeholder_type="dataframe-content",
            coord=cell["coordinate"],
            row_idx=row_idx,
            col_idx=col_idx,
            columns=columns,
            source_record=source_record,
            cell=cell,
        )
    )
    return anchors


def _dataframe_anchor(
    *,
    anchor_id: str,
    key: str,
    placeholder_type: str,
    coord: str,
    row_idx: int,
    col_idx: int,
    columns: list[str],
    source_record: dict[str, Any] | None,
    cell: CellSchema,
) -> dict[str, Any]:
    return {
        "id": anchor_id,
        "key": key,
        "placeholder_type": placeholder_type,
        "coordinate": coord,
        "start_row": row_idx,
        "start_col": col_idx,
        "columns": columns,
        "source": source_record["path"] if source_record else None,
        "source_format": source_record["format"] if source_record else None,
        "cell": cell,
    }


# §4. Public Functions


def compile_report_bundle(
    template: WorkbookSchema,
    data: dict[str, Any],
    bundle_path: str | None = None,
) -> ReportBundle:
    """Validate *data* against *template* and produce a directory ReportBundle."""
    bundle_dir = _prepare_bundle_dir(bundle_path)
    data_sources: list[dict[str, Any]] = []
    used_ids: set[str] = set()
    source_cache: dict[int, dict[str, Any]] = {}

    sheets = [
        _compile_sheet(
            sheet=sheet,
            output_name=output_name,
            sheet_data=sheet_data,
            bundle_dir=bundle_dir,
            data_sources=data_sources,
            used_ids=used_ids,
            source_cache=source_cache,
        )
        for sheet, output_name, sheet_data, _ in _resolve_sheet_payloads(template, data)
    ]

    report = {"version": BUNDLE_VERSION, "sheets": sheets, "assets": []}
    manifest = {
        "version": BUNDLE_VERSION,
        "bundle_format": "directory",
        "created_at_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "inputs": get_template_inputs(template),
        "sheets": [
            {
                "name": sheet["name"],
                "dataframe_anchors": [
                    anchor["id"] for anchor in sheet.get("dataframe_anchors", [])
                ],
            }
            for sheet in sheets
        ],
        "placeholder_resolution": {
            "scalars": "resolved in report.json",
            "dataframes": "stored as dataframe anchors pointing to data/*.parquet",
        },
        "dataframe_sources": data_sources,
        "assets": [],
        "output_capabilities": {"xlsx": True, "pdf": True, "image": False},
    }
    bundle = ReportBundle(manifest=manifest, report=report, path=str(bundle_dir))
    _write_bundle_metadata(bundle)
    return bundle


def load_report_bundle(bundle_path: str) -> ReportBundle:
    path = Path(bundle_path)
    if not path.exists():
        raise FileNotFoundError(f"Report bundle not found: {bundle_path}")
    if not path.is_dir():
        raise ValueError(f"Report bundle must be a directory: {bundle_path}")

    manifest_path = path / "manifest.json"
    report_path = path / "report.json"
    if not manifest_path.exists() or not report_path.exists():
        raise ValueError(
            "Malformed report bundle: expected manifest.json and report.json"
        )

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    report = json.loads(report_path.read_text(encoding="utf-8"))
    if manifest.get("version") != BUNDLE_VERSION:
        raise ValueError(
            f"Unsupported report bundle version: {manifest.get('version')!r}"
        )
    if manifest.get("bundle_format") != "directory":
        raise ValueError(
            f"Unsupported report bundle format: {manifest.get('bundle_format')!r}"
        )
    return ReportBundle(manifest=manifest, report=report, path=str(path))


# §5. Entrypoints
