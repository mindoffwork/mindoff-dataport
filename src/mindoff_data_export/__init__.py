from typing import Any, Literal, overload

from .builder import build_template
from .extractor import extract_template
from .renderer import get_template_inputs, render_schema
from .schema import WorkbookSchema
from .streaming import build_template_streaming_with_data

# §1 Types

# §2 Constants

# §3 Private Helpers

# §4 Public API


@overload
def build_template_with_data(
    schema: WorkbookSchema,
    data: dict[str, Any],
    output_path: str,
    *,
    column_width_mode: str | None = None,
    row_height_mode: str | None = None,
    default_column_width: float | None = None,
    default_row_height: float | None = None,
    export_mode: Literal["fidelity"] = "fidelity",
    streaming_chunk_rows: int = 50_000,
    max_rows_per_workbook: int = 1_048_576,
) -> None: ...


@overload
def build_template_with_data(
    schema: WorkbookSchema,
    data: dict[str, Any],
    output_path: str,
    *,
    column_width_mode: str | None = None,
    row_height_mode: str | None = None,
    default_column_width: float | None = None,
    default_row_height: float | None = None,
    export_mode: Literal["streaming"],
    streaming_chunk_rows: int = 50_000,
    max_rows_per_workbook: int = 1_048_576,
) -> list[str]: ...


def build_template_with_data(
    schema: WorkbookSchema,
    data: dict[str, Any],
    output_path: str,
    *,
    column_width_mode: str | None = None,
    row_height_mode: str | None = None,
    default_column_width: float | None = None,
    default_row_height: float | None = None,
    export_mode: Literal["fidelity", "streaming"] = "fidelity",
    streaming_chunk_rows: int = 50_000,
    max_rows_per_workbook: int = 1_048_576,
) -> None | list[str]:
    """Render placeholders in *schema* with *data* then build output workbook(s)."""
    if export_mode == "streaming":
        return build_template_streaming_with_data(
            schema=schema,
            data=data,
            output_path=output_path,
            column_width_mode=column_width_mode,
            row_height_mode=row_height_mode,
            default_column_width=default_column_width,
            default_row_height=default_row_height,
            streaming_chunk_rows=streaming_chunk_rows,
            max_rows_per_workbook=max_rows_per_workbook,
        )

    if export_mode != "fidelity":
        raise ValueError(
            f"Unsupported export_mode '{export_mode}'. Expected 'fidelity' or 'streaming'."
        )

    resolved = render_schema(schema, data)
    build_template(
        resolved,
        output_path,
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
