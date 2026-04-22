from typing import Any

from .builder import build_template
from .extractor import extract_template
from .renderer import get_template_inputs, render_schema
from .schema import WorkbookSchema


def build_template_with_data(
    schema: WorkbookSchema,
    data: dict[str, Any],
    output_path: str,
    *,
    column_width_mode: str | None = None,
    row_height_mode: str | None = None,
    default_column_width: float | None = None,
    default_row_height: float | None = None,
) -> None:
    """Render all {{key:type}} placeholders in *schema* with *data* then build the .xlsx.

    Parameters
    ----------
    schema      : WorkbookSchema — may contain {{key:type}} placeholders in cell values
    data        : mapping of placeholder key → value (scalars, DataFrames, LazyFrames)
    output_path : destination .xlsx path

    Sizing kwargs override per-sheet values when provided.
    """
    resolved = render_schema(schema, data)
    build_template(
        resolved, output_path,
        column_width_mode=column_width_mode,
        row_height_mode=row_height_mode,
        default_column_width=default_column_width,
        default_row_height=default_row_height,
    )


__all__ = [
    "extract_template",
    "build_template",
    "build_template_with_data",
    "get_template_inputs",
    "render_schema",
]
