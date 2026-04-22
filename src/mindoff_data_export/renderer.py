from __future__ import annotations

import datetime
import re
from typing import Any

from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)

from .schema import (
    AlignmentSchema,
    CellBorders,
    CellSchema,
    FillSchema,
    FontSchema,
    SheetSchema,
    WorkbookSchema,
)

# §1 Types

# §2 Constants

# Matches {{key:type}} where type may include hyphens (for example dataframe-headers).
PLACEHOLDER_RE = re.compile(r"\{\{(\w+):([\w-]+)\}\}")

_DATAFRAME_TYPES = frozenset(["dataframe-headers", "dataframe-content"])
_SCALAR_TYPES = frozenset(["string", "number", "int", "float", "date", "boolean"])
_ALL_TYPES = _SCALAR_TYPES | _DATAFRAME_TYPES

_EMPTY_BORDER_SIDE = {"style": None, "color": None}
_DEFAULT_FONT: FontSchema = {
    "name": "Calibri",
    "size": 11.0,
    "bold": False,
    "italic": False,
    "underline": None,
    "color": None,
}
_DEFAULT_FILL: FillSchema = {"bg_color": None}
_DEFAULT_ALIGNMENT: AlignmentSchema = {
    "horizontal": None,
    "vertical": None,
    "wrap_text": False,
}
_DEFAULT_BORDERS: CellBorders = {
    "top": _EMPTY_BORDER_SIDE,
    "bottom": _EMPTY_BORDER_SIDE,
    "left": _EMPTY_BORDER_SIDE,
    "right": _EMPTY_BORDER_SIDE,
}

# §3 Private Helpers


def _validate_data(placeholders: dict[str, str], data: dict[str, Any]) -> None:
    for key, expected_type in placeholders.items():
        if key not in data:
            raise KeyError(
                f"Template requires '{key}' (type: {expected_type}) but it was not provided in data"
            )
        _check_type(key, data[key], expected_type)


def _check_type(key: str, value: Any, expected: str) -> None:
    if expected == "dataframe-content":
        _assert_dataframe(key, value)
        return
    if expected == "dataframe-headers":
        _assert_headers_input(key, value)
        return

    type_map: dict[str, tuple] = {
        "string": (str,),
        "number": (int, float),
        "int": (int,),
        "float": (float,),
        "boolean": (bool,),
        "date": (str, datetime.datetime, datetime.date),
    }
    allowed = type_map.get(expected)
    if allowed and not isinstance(value, allowed):
        raise TypeError(f"'{key}' expected type '{expected}', got {type(value).__name__}")


def _assert_dataframe(key: str, value: Any) -> None:
    module = getattr(type(value), "__module__", "") or ""
    if "polars" in module or "pandas" in module:
        return
    raise TypeError(
        f"'{key}' expected a polars/pandas DataFrame or LazyFrame, got {type(value).__name__}"
    )


def _assert_headers_input(key: str, value: Any) -> None:
    module = getattr(type(value), "__module__", "") or ""
    if "polars" in module or "pandas" in module:
        return
    if isinstance(value, list):
        return
    raise TypeError(
        f"'{key}' expected a polars/pandas DataFrame/LazyFrame or list of header values, got {type(value).__name__}"
    )


def _render_sheet(sheet: SheetSchema, data: dict[str, Any]) -> SheetSchema:
    min_col, min_row, max_col, max_row = _parse_dims(sheet["dimensions"])

    new_cells: dict[str, CellSchema] = {}
    expanded_coords: set[str] = set()

    for coord, cell_schema in sheet["cells"].items():
        if coord in expanded_coords:
            # Preserve generated cells from dataframe expansion instead of template stubs.
            continue

        value = cell_schema.get("value")
        if not isinstance(value, str):
            new_cells[coord] = cell_schema
            continue

        full_match = PLACEHOLDER_RE.fullmatch(value.strip())
        if (
            full_match
            and full_match.group(2) in _DATAFRAME_TYPES
            and full_match.group(1) in data
        ):
            key = full_match.group(1)
            placeholder_type = full_match.group(2)
            write_headers = placeholder_type == "dataframe-headers"
            write_rows = placeholder_type == "dataframe-content"
            df_cells, df_max_row, df_max_col = _expand_dataframe(
                cell_schema,
                coord,
                data[key],
                write_headers,
                write_rows,
            )
            new_cells.update(df_cells)
            max_row = max(max_row, df_max_row)
            max_col = max(max_col, df_max_col)

            anchor_col_letter, anchor_row = coordinate_from_string(coord)
            anchor_col = column_index_from_string(anchor_col_letter)
            for r in range(anchor_row, df_max_row + 1):
                for c in range(anchor_col, df_max_col + 1):
                    expanded_coords.add(f"{get_column_letter(c)}{r}")
            continue

        new_value = _substitute_scalars(value, data)
        new_cell: dict[str, Any] = dict(cell_schema)
        new_cell["value"] = new_value
        if full_match and full_match.group(2) in _SCALAR_TYPES:
            new_cell["cell_type"] = _infer_cell_type(new_value)
        new_cells[coord] = new_cell  # type: ignore[assignment]

    new_dims = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    result: dict[str, Any] = dict(sheet)
    result["cells"] = new_cells
    result["dimensions"] = new_dims
    return result  # type: ignore[return-value]


def _substitute_scalars(value: str, data: dict[str, Any]) -> Any:
    """Replace scalar placeholders; whole-cell placeholders return raw values."""
    full_match = PLACEHOLDER_RE.fullmatch(value.strip())
    if full_match:
        key, placeholder_type = full_match.group(1), full_match.group(2)
        if placeholder_type in _SCALAR_TYPES and key in data:
            return data[key]

    def replacer(match: re.Match[str]) -> str:
        key, placeholder_type = match.group(1), match.group(2)
        if placeholder_type not in _SCALAR_TYPES or key not in data:
            return match.group(0)
        return str(data[key])

    return PLACEHOLDER_RE.sub(replacer, value)


def _infer_cell_type(value: Any) -> str:
    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "string"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, (datetime.datetime, datetime.date)):
        return "date"
    return "string"


def _expand_dataframe(
    anchor: CellSchema,
    coord: str,
    df: Any,
    write_headers: bool,
    write_rows: bool,
) -> tuple[dict[str, CellSchema], int, int]:
    """Build expanded dataframe cells starting at *coord*.

    Returns (cells_dict, max_row, max_col_index).
    """
    col_letter, start_row = coordinate_from_string(coord)
    start_col = column_index_from_string(col_letter)

    columns: list[str] = []
    rows: list[tuple[Any, ...]] = []
    if write_rows:
        columns, rows = _to_rows(df)
    elif write_headers:
        columns = _to_headers(df)

    if not columns and rows:
        columns = [f"col_{i + 1}" for i in range(len(rows[0]))]
    n_cols = len(columns)
    if n_cols == 0:
        return {}, start_row, start_col

    font: FontSchema = anchor.get("font", _DEFAULT_FONT)  # type: ignore[assignment]
    fill: FillSchema = anchor.get("fill", _DEFAULT_FILL)  # type: ignore[assignment]
    alignment: AlignmentSchema = anchor.get("alignment", _DEFAULT_ALIGNMENT)  # type: ignore[assignment]
    borders: CellBorders = anchor.get("borders", _DEFAULT_BORDERS)  # type: ignore[assignment]

    cells: dict[str, CellSchema] = {}
    current_row = start_row

    if write_headers:
        bold_font: FontSchema = dict(font)  # type: ignore[assignment]
        bold_font["bold"] = True
        for col_offset, col_name in enumerate(columns):
            c_letter = get_column_letter(start_col + col_offset)
            c_coord = f"{c_letter}{current_row}"
            cells[c_coord] = {
                "coordinate": c_coord,
                "value": str(col_name),
                "cell_type": "string",
                "number_format": None,
                "font": bold_font,
                "fill": fill,
                "alignment": alignment,
                "borders": borders,
                "merged": False,
                "merge_anchor": None,
            }
        current_row += 1

    for row_values in rows:
        for col_offset, val in enumerate(row_values):
            c_letter = get_column_letter(start_col + col_offset)
            c_coord = f"{c_letter}{current_row}"
            cells[c_coord] = {
                "coordinate": c_coord,
                "value": val,
                "cell_type": _infer_cell_type(val),
                "number_format": None,
                "font": font,
                "fill": fill,
                "alignment": alignment,
                "borders": borders,
                "merged": False,
                "merge_anchor": None,
            }
        current_row += 1

    max_row = current_row - 1
    max_col = start_col + n_cols - 1
    return cells, max_row, max_col


def _to_rows(df: Any) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Return (column_names, row_tuples) from polars or pandas frames."""
    module = getattr(type(df), "__module__", "") or ""
    qualname = type(df).__qualname__

    if "polars" in module:
        if qualname == "LazyFrame":
            df = df.collect()
        return list(df.columns), list(df.rows())

    if "pandas" in module and "DataFrame" in qualname:
        return list(df.columns), [tuple(row) for row in df.itertuples(index=False)]

    raise TypeError(f"Expected a polars or pandas DataFrame/LazyFrame, got {type(df).__name__}")


def _to_headers(df_or_headers: Any) -> list[str]:
    """Return header names from DataFrame/LazyFrame or list input."""
    module = getattr(type(df_or_headers), "__module__", "") or ""
    qualname = type(df_or_headers).__qualname__

    if "polars" in module:
        if qualname == "LazyFrame":
            df_or_headers = df_or_headers.collect()
        return [str(col) for col in df_or_headers.columns]

    if "pandas" in module and "DataFrame" in qualname:
        return [str(col) for col in df_or_headers.columns]

    if isinstance(df_or_headers, list):
        return [str(col) for col in df_or_headers]

    raise TypeError(
        f"Expected a polars/pandas DataFrame/LazyFrame or list of headers, got {type(df_or_headers).__name__}"
    )


def _parse_dims(dimensions: str) -> tuple[int, int, int, int]:
    """Parse 'A1:F9' to (min_col, min_row, max_col, max_row)."""
    parts = dimensions.split(":")
    start = parts[0]
    end = parts[1] if len(parts) > 1 else parts[0]
    start_col_letter, start_row = coordinate_from_string(start)
    end_col_letter, end_row = coordinate_from_string(end)
    return (
        column_index_from_string(start_col_letter),
        start_row,
        column_index_from_string(end_col_letter),
        end_row,
    )


# §4 Public API


def get_template_inputs(schema: WorkbookSchema) -> dict[str, str]:
    """Return {key: type} for every valid {{key:type}} placeholder in cell values."""
    found: dict[str, str] = {}
    for sheet in schema["sheets"]:
        for cell_schema in sheet["cells"].values():
            value = cell_schema.get("value")
            if isinstance(value, str):
                for match in PLACEHOLDER_RE.finditer(value):
                    key, placeholder_type = match.group(1), match.group(2)
                    if placeholder_type in _ALL_TYPES:
                        found[key] = placeholder_type
    return found


def render_schema(schema: WorkbookSchema, data: dict[str, Any]) -> WorkbookSchema:
    """Validate data against placeholders and return a resolved workbook schema."""
    placeholders = get_template_inputs(schema)
    _validate_data(placeholders, data)
    return {"sheets": [_render_sheet(sheet, data) for sheet in schema["sheets"]]}
