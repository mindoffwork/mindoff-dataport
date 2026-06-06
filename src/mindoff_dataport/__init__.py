from __future__ import annotations

from types import SimpleNamespace
from typing import Any, Literal

from .bundle import (
    ReportBundle,
    compile_report_bundle as _compile_report_bundle_impl,
)
from .extractor import extract_template as _extract_template_impl
from .repeat import RepeatRecords, repeat_records
from .schema import WorkbookSchema
from .template_contract import get_template_inputs as _get_template_inputs_impl
from .xlsx_renderer import export_report_bundle as _export_report_bundle_impl

__all__ = [
    "ReportBundle",
    "RepeatRecords",
    "extract_template",
    "get_template_inputs",
    "compile_report_bundle",
    "export_report_bundle",
    "repeat_records",
    "extract",
    "inputs",
    "compile",
    "export",
    "mo_dataport",
]

# §1. Constants & Exceptions

# §2. Classes and Sub Classes

# §3. Private Helper Functions

# §4. Public Functions


def extract_template(path: str) -> WorkbookSchema:
    """
    Read an `.xlsx` template and return a `WorkbookSchema`.

    Extraction captures cell values, styles (font, fill, alignment, borders),
    column widths and row heights, merged regions, manual print breaks, theme
    colors, and every `{{key:type}}` placeholder it finds. The schema is the
    in-memory blueprint everything downstream is built from.

    **Usage**

    ```python
    from mindoff_dataport import mo_dataport

    schema = mo_dataport.extract("invoice_template.xlsx")

    # or the explicit name:
    from mindoff_dataport import extract_template
    schema = extract_template("invoice_template.xlsx")
    ```

    | Parameter | Type  | Required | Description                       |
    |-----------|-------|----------|-----------------------------------|
    | `path`    | `str` | Yes      | Path to the `.xlsx` template file |

    **Returns:** `WorkbookSchema` — the extracted template blueprint.
    """
    return _extract_template_impl(path)


def get_template_inputs(template: WorkbookSchema) -> dict[str, Any]:
    """
    Inspect a schema and report every input the template expects.

    Returns a sheet-scoped dictionary keyed by sheet name, then by placeholder
    key, with the value being the placeholder type (`"string"`, `"number"`,
    `"date"`, `"dataframe"`, and so on). Call this before building a payload so
    you know exactly what `compile()` will ask for — no guessing, no trial runs.

    **Usage**

    ```python
    contract = mo_dataport.inputs(schema)

    # or the explicit name:
    from mindoff_dataport import get_template_inputs
    contract = get_template_inputs(schema)
    ```

    | Parameter  | Type             | Required | Description                    |
    |------------|------------------|----------|--------------------------------|
    | `template` | `WorkbookSchema` | Yes      | Schema produced by `extract()` |

    **Returns:** `dict[str, dict[str, str | list]]` — the per-sheet input contract.

    **Example output**

    ```python
    {
        "Sales Summary": {
            "report_title": "string",
            "generated_on": "date",
            "sales_rows": "dataframe",
        }
    }
    ```
    """
    return _get_template_inputs_impl(template)


def compile_report_bundle(
    template: WorkbookSchema,
    data: dict[str, Any],
    bundle_path: str | None = None,
    dataframe_options: dict[str, Any] | None = None,
    dataframe_shift: Literal["both", "horizontal", "vertical", "none"] = "both",
) -> ReportBundle:
    """
    Bind runtime data to a template and produce a `ReportBundle`.

    Compilation validates the payload against the sheet contract, resolves
    scalar cells in place, materialises Polars DataFrames / LazyFrames to
    Parquet, stores compact dataframe anchors and repeat plans, and (optionally)
    shifts template content out of the way of expanding dataframes. The result
    is a portable bundle you can export now, or persist on disk and export
    later from any process.

    **Usage**

    ```python
    bundle = mo_dataport.compile(
        template=schema,
        data=payload,
        bundle_path="out_bundle",      # omit to keep the bundle in memory
        dataframe_options=None,
        dataframe_shift="both",
    )
    ```

    | Parameter           | Type                     | Required | Description |
    |---------------------|--------------------------|----------|-------------|
    | `template`          | `WorkbookSchema`         | Yes      | Schema from `extract()` |
    | `data`              | `dict[str, Any]`         | Yes      | Sheet-scoped payload (see the Data Contract guide) |
    | `bundle_path`       | `str | None`            | No       | Write the bundle to this directory; omit for in-memory only |
    | `dataframe_options` | `dict[str, Any] | None` | No       | Per-sheet, per-placeholder column occupation and alignment overrides |
    | `dataframe_shift`   | `str`                    | No       | How surrounding cells/merges move around dataframe output: `"both"`, `"horizontal"`, `"vertical"`, or `"none"` |

    **Returns:** `ReportBundle` — the compiled, exportable bundle.

    **Raises:** `KeyError` if a required placeholder key is missing from the payload.
    """
    return _compile_report_bundle_impl(
        template,
        data,
        bundle_path=bundle_path,
        dataframe_options=dataframe_options,
        dataframe_shift=dataframe_shift,
    )


def export_report_bundle(
    bundle_or_path: ReportBundle | str,
    output_path: str,
    format: Literal["xlsx", "pdf", "image"] = "xlsx",
    **options: Any,
) -> None | list[str]:
    """
    Render a compiled bundle to a file on disk.

    Accepts either an in-memory `ReportBundle` or a path to a persisted bundle
    directory, and writes `.xlsx` or `.pdf` output. XLSX supports a full-fidelity
    mode and a low-memory streaming mode; PDF always paginates automatically.
    Format-specific keyword options (export mode, sizing, fonts, page size, and
    so on) are passed through `**options`.

    **Usage**

    ```python
    mo_dataport.export(bundle, "report.xlsx", format="xlsx")
    mo_dataport.export("out_bundle", "report.pdf", format="pdf")
    ```

    | Parameter        | Type                  | Required | Default  | Description |
    |------------------|-----------------------|----------|----------|-------------|
    | `bundle_or_path` | `ReportBundle | str` | Yes      | —        | In-memory bundle or path to a bundle directory |
    | `output_path`    | `str`                 | Yes      | —        | Destination file path (`.xlsx` or `.pdf`) |
    | `format`         | `str`                 | No       | `"xlsx"` | `"xlsx"` or `"pdf"`. `"image"` is reserved and raises `NotImplementedError` |
    | `**options`      | —                     | No       | —        | Sizing and format-specific options (see the Exporting guides) |

    **Returns:** `None` for fidelity XLSX and all PDF exports. For streaming
    XLSX, a `list[str]`: one workbook path when no split is needed, or a single
    `.zip` path when the export is split across multiple workbooks.
    """
    return _export_report_bundle_impl(
        bundle_or_path,
        output_path,
        format=format,
        **options,
    )


extract = extract_template
inputs = get_template_inputs
compile = compile_report_bundle
export = export_report_bundle

mo_dataport = SimpleNamespace(
    extract=extract_template,
    inputs=get_template_inputs,
    compile=compile_report_bundle,
    export=export_report_bundle,
    repeat_records=repeat_records,
)


# §5. Entrypoints
