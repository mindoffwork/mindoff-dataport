from typing import Any, Literal, overload

from .builder import build_template as _build_template
from .extractor import extract_template
from .renderer import get_template_inputs, render_schema
from .schema import WorkbookSchema
from .streaming import build_template_streaming_with_data

__all__ = [
    "extract_template",
    "build_template_with_data",
    "get_template_inputs",
    "render_schema",
    "mode",
]

# §1 Constants & Exceptions

# §2 Classes and Sub Classes


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
    streaming_progress: bool = False,
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
    streaming_progress: bool = False,
) -> list[str]: ...


# §3 Private Helper Functions


def _build_with_fidelity(
    schema: WorkbookSchema,
    data: dict[str, Any],
    output_path: str,
    *,
    column_width_mode: str | None = None,
    row_height_mode: str | None = None,
    default_column_width: float | None = None,
    default_row_height: float | None = None,
) -> None:
    resolved = render_schema(schema, data)
    _build_template(
        resolved,
        output_path,
        column_width_mode=column_width_mode,
        row_height_mode=row_height_mode,
        default_column_width=default_column_width,
        default_row_height=default_row_height,
    )


# §4 Public Functions


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
    streaming_progress: bool = False,
) -> None | list[str]:
    """Render placeholders in *schema* with sheet-scoped *data* then build output workbook(s)."""
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
            streaming_progress=streaming_progress,
        )
    if export_mode != "fidelity":
        raise ValueError(
            f"Unsupported export_mode '{export_mode}'. Expected 'fidelity' or 'streaming'."
        )

    _build_with_fidelity(
        schema=schema,
        data=data,
        output_path=output_path,
        column_width_mode=column_width_mode,
        row_height_mode=row_height_mode,
        default_column_width=default_column_width,
        default_row_height=default_row_height,
    )
    return None


class _ModeAPI:
    """Convenience namespace for a single-import public API."""

    extract = staticmethod(extract_template)
    build = staticmethod(build_template_with_data)
    get_inputs = staticmethod(get_template_inputs)
    alter_schema = staticmethod(render_schema)


mode = _ModeAPI()

# §5 Entrypoints
