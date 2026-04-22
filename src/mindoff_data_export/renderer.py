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

# Matches {{key:type}} where type may include hyphens (e.g. "dataframe-headers")
PLACEHOLDER_RE = re.compile(r"\{\{(\w+):([\w-]+)\}\}")

_DATAFRAME_TYPES = frozenset(["dataframe-headers", "dataframe-data"])
_SCALAR_TYPES = frozenset(["string", "number", "int", "float", "date", "boolean"])
_ALL_TYPES = _SCALAR_TYPES | _DATAFRAME_TYPES

_EMPTY_BORDER_SIDE = {"style": None, "color": None}
_DEFAULT_FONT: FontSchema = {
    "name": "Calibri", "size": 11.0, "bold": False,
    "italic": False, "underline": None, "color": None,
}
_DEFAULT_FILL: FillSchema = {"bg_color": None}
_DEFAULT_ALIGNMENT: AlignmentSchema = {"horizontal": None, "vertical": None, "wrap_text": False}
_DEFAULT_BORDERS: CellBorders = {
    "top": _EMPTY_BORDER_SIDE, "bottom": _EMPTY_BORDER_SIDE,
    "left": _EMPTY_BORDER_SIDE, "right": _EMPTY_BORDER_SIDE,
}


def get_template_inputs(schema: WorkbookSchema) -> dict[str, str]:
    """Return {key: type} for every {{key:type}} placeholder found in cell values.

    Only known types are returned; legacy {{key}} / {{table:key}} markers are ignored.
    """
    found: dict[str, str] = {}
    for sheet in schema["sheets"]:
        for cell_schema in sheet["cells"].values():
            value = cell_schema.get("value")
            if isinstance(value, str):
                for match in PLACEHOLDER_RE.finditer(value):
                    key, type_ = match.group(1), match.group(2)
                    if type_ in _ALL_TYPES:
                        found[key] = type_
    return found


def render_schema(schema: WorkbookSchema, data: dict[str, Any]) -> WorkbookSchema:
    """Validate *data* against template placeholders and return a resolved WorkbookSchema.

    The returned schema has all {{key:type}} tokens replaced with actual values and
    dataframe placeholders expanded into individual CellSchema entries.
    """
    placeholders = get_template_inputs(schema)
    _validate_data(placeholders, data)
    return {"sheets": [_render_sheet(sheet, data) for sheet in schema["sheets"]]}


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def _validate_data(placeholders: dict[str, str], data: dict[str, Any]) -> None:
    for key, expected_type in placeholders.items():
        if key not in data:
            raise KeyError(
                f"Template requires '{key}' (type: {expected_type}) but it was not provided in data"
            )
        _check_type(key, data[key], expected_type)


def _check_type(key: str, value: Any, expected: str) -> None:
    if expected in _DATAFRAME_TYPES:
        _assert_dataframe(key, value)
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
        raise TypeError(
            f"'{key}' expected type '{expected}', got {type(value).__name__}"
        )


def _assert_dataframe(key: str, value: Any) -> None:
    module = getattr(type(value), "__module__", "") or ""
    if "polars" in module or "pandas" in module:
        return
    raise TypeError(
        f"'{key}' expected a polars/pandas DataFrame or LazyFrame, got {type(value).__name__}"
    )


# ---------------------------------------------------------------------------
# Sheet rendering
# ---------------------------------------------------------------------------

def _render_sheet(sheet: SheetSchema, data: dict[str, Any]) -> SheetSchema:
    dims = sheet["dimensions"]
    min_col, min_row, max_col, max_row = _parse_dims(dims)

    new_cells: dict[str, CellSchema] = {}
    expanded_coords: set[str] = set()

    for coord, cell_schema in sheet["cells"].items():
        if coord in expanded_coords:
            # This coordinate is already produced by a dataframe expansion.
            # Keep expanded data instead of overwriting it with original template stubs.
            continue

        value = cell_schema.get("value")
        if not isinstance(value, str):
            new_cells[coord] = cell_schema
            continue

        full_match = PLACEHOLDER_RE.fullmatch(value.strip())
        if full_match and full_match.group(2) in _DATAFRAME_TYPES and full_match.group(1) in data:
            key = full_match.group(1)
            write_headers = full_match.group(2) == "dataframe-headers"
            df_cells, df_max_row, df_max_col = _expand_dataframe(
                cell_schema, coord, data[key], write_headers
            )
            new_cells.update(df_cells)
            max_row = max(max_row, df_max_row)
            max_col = max(max_col, df_max_col)

            anchor_col_letter, anchor_row = coordinate_from_string(coord)
            anchor_col = column_index_from_string(anchor_col_letter)
            for r in range(anchor_row, df_max_row + 1):
                for c in range(anchor_col, df_max_col + 1):
                    expanded_coords.add(f"{get_column_letter(c)}{r}")
        else:
            new_value = _substitute_scalars(value, data)
            new_cell: dict = dict(cell_schema)
            new_cell["value"] = new_value
            if full_match and full_match.group(2) in _SCALAR_TYPES:
                new_cell["cell_type"] = _infer_cell_type(new_value)
            new_cells[coord] = new_cell  # type: ignore[arg-type]

    min_col_letter = get_column_letter(min_col)
    max_col_letter = get_column_letter(max_col)
    new_dims = f"{min_col_letter}{min_row}:{max_col_letter}{max_row}"

    result: dict = dict(sheet)
    result["cells"] = new_cells
    result["dimensions"] = new_dims
    return result  # type: ignore[return-value]


def _substitute_scalars(value: str, data: dict[str, Any]) -> Any:
    """Replace scalar placeholders. Whole-cell placeholder returns raw value; embedded → string."""
    full_match = PLACEHOLDER_RE.fullmatch(value.strip())
    if full_match:
        key, type_ = full_match.group(1), full_match.group(2)
        if type_ in _SCALAR_TYPES and key in data:
            return data[key]

    def replacer(m: re.Match) -> str:
        key, type_ = m.group(1), m.group(2)
        if type_ not in _SCALAR_TYPES or key not in data:
            return m.group(0)
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


# ---------------------------------------------------------------------------
# Dataframe expansion
# ---------------------------------------------------------------------------

def _expand_dataframe(
    anchor: CellSchema,
    coord: str,
    df: Any,
    write_headers: bool,
) -> tuple[dict[str, CellSchema], int, int]:
    """Build cells for the expanded dataframe starting at *coord*.

    Returns (cells_dict, max_row, max_col_index).
    """
    col_letter, start_row = coordinate_from_string(coord)
    start_col = column_index_from_string(col_letter)

    columns, rows = _to_rows(df)
    n_cols = len(columns)

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


def _to_rows(df: Any) -> tuple[list[str], list[tuple]]:
    """Return (column_names, list_of_row_tuples) from a polars or pandas DataFrame/LazyFrame."""
    module = getattr(type(df), "__module__", "") or ""
    qualname = type(df).__qualname__

    if "polars" in module:
        if qualname == "LazyFrame":
            df = df.collect()
        return list(df.columns), list(df.rows())

    if "pandas" in module and "DataFrame" in qualname:
        return list(df.columns), [tuple(row) for row in df.itertuples(index=False)]

    raise TypeError(
        f"Expected a polars or pandas DataFrame/LazyFrame, got {type(df).__name__}"
    )


# ---------------------------------------------------------------------------
# Dimension helpers
# ---------------------------------------------------------------------------

def _parse_dims(dimensions: str) -> tuple[int, int, int, int]:
    """Parse "A1:F9" → (min_col, min_row, max_col, max_row) as integers."""
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
