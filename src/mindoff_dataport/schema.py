from __future__ import annotations

from typing import Literal, Optional, TypedDict

__all__ = [
    "CellType",
    "BorderSide",
    "CellBorders",
    "FontSchema",
    "FillSchema",
    "AlignmentSchema",
    "CellSchema",
    "SheetSchema",
    "WorkbookSchema",
]

# §1. Constants & Exceptions

# §2. Classes and Sub Classes

CellType = Literal["string", "number", "date", "formula", "empty"]


class BorderSide(TypedDict):
    style: Optional[str]  # e.g. "thin", "medium", "thick", "dashed", "dotted"
    color: Optional[str]  # ARGB hex string e.g. "FF000000"


class CellBorders(TypedDict):
    top: BorderSide
    bottom: BorderSide
    left: BorderSide
    right: BorderSide
    start: BorderSide      # LTR/RTL-aware: start of text direction (left in LTR)
    end: BorderSide        # LTR/RTL-aware: end of text direction (right in LTR)
    horizontal: BorderSide  # interior horizontal divider (used in merged ranges)
    vertical: BorderSide    # interior vertical divider (used in merged ranges)
    diagonal: BorderSide   # diagonal line style
    diagonal_up: bool      # draw diagonal from bottom-left to top-right
    diagonal_down: bool    # draw diagonal from top-left to bottom-right
    outline: bool          # apply border as outline of selection


class FontSchema(TypedDict):
    name: Optional[str]
    size: Optional[float]
    bold: bool
    italic: bool
    underline: Optional[str]  # None, "single", "double"
    strike: bool              # strikethrough
    vert_align: Optional[str]  # None, "superscript", "subscript", "baseline"
    color: Optional[str]  # ARGB hex string


class FillSchema(TypedDict):
    pattern_type: Optional[str]  # "solid", "gray125", "darkGray", etc.; None = no fill
    fg_color: Optional[str]      # foreground/pattern color (the visible color for solid fills)
    bg_color: Optional[str]      # background color (used with patterned fills)


class AlignmentSchema(TypedDict):
    horizontal: Optional[str]        # "left", "center", "right", "fill", "justify", "general"
    vertical: Optional[str]          # "top", "center", "bottom", "justify"
    wrap_text: bool
    indent: Optional[int]            # indent level (1 unit ≈ 1 char width)
    relative_indent: Optional[int]   # relative indent adjustment
    text_rotation: Optional[int]     # degrees 0–180 (255 = vertical text)
    shrink_to_fit: bool              # shrink font to fit cell width
    reading_order: Optional[int]


class CellSchema(TypedDict):
    coordinate: str
    value: object  # str | int | float | None; dates as ISO 8601 string
    cell_type: CellType
    number_format: Optional[str]
    font: FontSchema
    fill: FillSchema
    alignment: AlignmentSchema
    borders: CellBorders
    merged: bool
    merge_anchor: Optional[str]  # coordinate of top-left cell in the merge region


class _SheetSchemaRequired(TypedDict):
    name: str
    dimensions: str  # e.g. "A1:Z100"
    merged_regions: list[str]  # ["A1:C3", "D5:E6"]
    column_widths: dict[str, float]  # {"A": 12.5, "B": 8.0}
    row_heights: dict[str, float]  # {"1": 20.0, "2": 15.0} str keys for JSON compat
    cells: dict[str, CellSchema]  # keyed by coordinate e.g. "A1"


class SheetSchema(_SheetSchemaRequired, total=False):
    # Sizing modes - optional; default is "fixed" (use explicit column_widths / row_heights)
    column_width_mode: Literal["fixed", "even", "hug"]
    default_column_width: float  # width applied to every column in "even" mode
    row_height_mode: Literal["fixed", "even", "hug"]
    default_row_height: float  # height applied to every row in "even" mode
    row_page_breaks: list[int]
    column_page_breaks: list[int]
    show_gridlines: bool
    repeat_sections: list[dict]


class WorkbookSchema(TypedDict):
    sheets: list[SheetSchema]


# §3. Private Helper Functions

# §4. Public Functions

# §5. Entrypoints
