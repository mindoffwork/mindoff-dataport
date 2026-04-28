from __future__ import annotations

from typing import Any, Literal

from .bundle import (
    ReportBundle,
    compile_report_bundle as _compile_report_bundle_impl,
)
from .extractor import extract_template as _extract_template_impl
from .template_contract import get_template_inputs as _get_template_inputs_impl
from .schema import WorkbookSchema
from .xlsx_renderer import export_report_bundle as _export_report_bundle_impl

__all__ = [
    "ReportBundle",
    "extract_template",
    "get_template_inputs",
    "compile_report_bundle",
    "export_report_bundle",
    "mode",
]

# §1 Constants & Exceptions

# §2 Classes and Sub Classes

# §3 Private Helper Functions

# §4 Public Functions


def extract_template(path: str) -> WorkbookSchema:
    return _extract_template_impl(path)


def get_template_inputs(template: WorkbookSchema) -> dict[str, Any]:
    return _get_template_inputs_impl(template)


def compile_report_bundle(
    template: WorkbookSchema,
    data: dict[str, Any],
    bundle_path: str | None = None,
) -> ReportBundle:
    return _compile_report_bundle_impl(template, data, bundle_path=bundle_path)


def export_report_bundle(
    bundle_or_path: ReportBundle | str,
    output_path: str,
    format: Literal["xlsx", "pdf", "image"] = "xlsx",
    **options: Any,
) -> None | list[str]:
    return _export_report_bundle_impl(
        bundle_or_path,
        output_path,
        format=format,
        **options,
    )


class _ModeAPI:
    """Bundle-first public API namespace."""

    extract = staticmethod(extract_template)
    inputs = staticmethod(get_template_inputs)
    compile = staticmethod(compile_report_bundle)
    export = staticmethod(export_report_bundle)


mode = _ModeAPI()

# §5 Entrypoints
