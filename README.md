# Mindoff Dataport

Extract Excel templates, compile runtime data into a canonical `ReportBundle`, and export production `.xlsx` or `.pdf` files with layout and styling preserved.

Recommended entrypoint: `from mindoff_dataport import mode as mo_dataport`

## What This Library Does

- Extract an `.xlsx` workbook into a JSON-like template schema
- Discover typed template inputs from placeholders such as `{{key:type}}`
- Compile template + runtime data into a portable `ReportBundle` directory
- Store dataframe sources under `data/*` without expanding rows into `report.json`
- Render ordered repeated sections into one sheet without materializing dataframe rows
- Export the bundle to `.xlsx` or styled `.pdf`, with a reserved `image` interface

## Install

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -e .
```

Optional demo/dataframe dependency:

```bash
pip install polars
```

## Quick Start

```python
from mindoff_dataport import mode as mo_dataport

template = mo_dataport.extract("template.xlsx")
required_inputs = mo_dataport.inputs(template)

bundle = mo_dataport.compile(
    template,
    data={
        "Sheet1": {
            "customer_name": "Acme Industries",
            "invoice_number": 1024,
        },
    },
    bundle_path="report_bundle",
)

mo_dataport.export(
    bundle,
    "filled.xlsx",
    format="xlsx",
    column_width_mode="fixed",
    row_height_mode="fixed",
)
```

## API Surface

Import `mode` directly or alias it for readability:

```python
from mindoff_dataport import mode as mo_dataport
```

- `mo_dataport.extract(path)`
- `mo_dataport.inputs(template)`
- `mo_dataport.compile(template, data, bundle_path=None)`
- `mo_dataport.export(bundle_or_path, output_path, format="xlsx", **options)`

Top-level exports mirror the namespace:

- `extract_template`
- `get_template_inputs`
- `compile_report_bundle`
- `export_report_bundle`
- `mode`

## Workflow

```python
from mindoff_dataport import mode as mo_dataport

template = mo_dataport.extract("invoice_template.xlsx")
inputs = mo_dataport.inputs(template)
bundle = mo_dataport.compile(template, {"Sheet1": {"customer_name": "Acme"}})
mo_dataport.export(bundle, "filled.xlsx")
```

`mo_dataport.compile(...)` returns an in-memory `ReportBundle`. When `bundle_path` is provided, the same artifact is also written as a directory containing:

- `manifest.json`: bundle version, inputs, sheet/page metadata, dataframe sources, assets, and output capabilities
- `report.json`: resolved scalar/static cells plus dataframe anchors
- `data/*.parquet`: dataframe sources materialized from Polars DataFrame/LazyFrame inputs
- `assets/*`: reserved for future image/logo payloads

## Export Options

`mo_dataport.export(..., format="xlsx")` supports the current sizing controls:

- `column_width_mode`: `"fixed"`, `"even"`, or `"hug"`
- `row_height_mode`: `"fixed"`, `"even"`, or `"hug"`
- `default_column_width`
- `default_row_height`
- `export_mode`: `"fidelity"` or `"streaming"`
- `max_rows_per_workbook`

`mo_dataport.export(..., format="pdf")` renders each workbook sheet as a styled report page
that can continue vertically onto additional PDF pages. It supports the same sizing
overrides plus PDF options:

- `page_size`: `"A4"`, `"LETTER"`, or `"LEGAL"`
- `orientation`: `"portrait"` or `"landscape"`
- `margin`
- `streaming_chunk_rows`

`format="image"` is reserved and raises `NotImplementedError` in v1.

## Data Contract

Runtime data is sheet-scoped. Dynamic sheet groups preserve payload order.

```python
{
    "Sheet1": {
        "customer_name": "Acme Industries",
        "line_items": polars_dataframe_or_lazyframe,
    },
    "region_sheet": {
        "North": {"region_name": "North"},
        "South": {"region_name": "South"},
    },
}
```

Supported placeholder types:

- Scalars: `string`, `number`, `int`, `float`, `date`, `boolean`
- Dataframes: `dataframe`, `dataframe-header`, `dataframe-content`
- Repeats: `repeat-start`, `repeat-end`

Use `polars.scan_parquet(path)` when the source data starts as Parquet and should remain lazy until compilation.

### Single-Sheet Repeated Sections

Use repeat markers when one template block should be rendered many times in the
same sheet. The rows between the markers are repeated; marker rows are control
rows and are not rendered. A sheet may contain multiple non-overlapping sibling
repeat sections, processed in template row order.

```text
{{reports:repeat-start}}
Customer: {{customer_name:string}}
{{line_items:dataframe-header}}
{{line_items:dataframe-content}}
{{reports:repeat-end}}
```

```python
bundle = mo_dataport.compile(
    template,
    {
        "Sheet1": {
            "reports": [
                {"customer_name": "Acme", "line_items": acme_rows},
                {"customer_name": "Globex", "line_items": globex_rows},
                {"customer_name": "Initech", "line_items": initech_rows},
            ]
        }
    },
)

mo_dataport.export(bundle, "combined.xlsx", export_mode="streaming")
mo_dataport.export(bundle, "combined.pdf", format="pdf", streaming_chunk_rows=5000)
```

Repeat v1 supports ordered list payloads, one or more sibling vertical sections
per sheet, static rows before/between/after sections, unique repeat keys, and no
nested repeats. Merged cells are supported in fixed repeat/static rows, but not
over streamed `dataframe-content` rows. Dataframe content remains parquet-backed
and is streamed in batches for XLSX and PDF output.

## Demo

The examples use the tracked workbook fixture at `examples/template.xlsx` and the tracked
Parquet fixture at `examples/data.parquet`. Generated files are written under
`examples/output/`, which is ignored by git except for `examples/output/.gitkeep`.

```bash
python examples/xlsx_output.py
python examples/pdf_output.py
```

Typical generated outputs:

- `examples/output/styled_parquet_output.part001.xlsx`
- `examples/output/styled_parquet_output.pdf`

## Current Scope

- Template input: `.xlsx`
- Canonical intermediate: `ReportBundle` directory
- Production output: `.xlsx`, `.pdf`
- Reserved output interfaces: `image`
