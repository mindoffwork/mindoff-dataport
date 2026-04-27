# Mindoff Dataport

Extract Excel templates, compile runtime data into a canonical `ReportBundle`, and export production `.xlsx` files with layout and styling preserved.

Primary entrypoint: `from mindoff_dataport import mode`

## What This Library Does

- Extract an `.xlsx` workbook into a JSON-like template schema
- Discover typed template inputs from placeholders such as `{{key:type}}`
- Compile template + runtime data into a portable `ReportBundle` directory
- Store dataframe sources under `data/*` without expanding rows into `report.json`
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

Optional parquet bundle support:

```bash
pip install -e ".[parquet]"
```

Optional demo/dataframe dependency:

```bash
pip install polars
```

## Quick Start

```python
from mindoff_dataport import mode

template = mode.extract("template.xlsx")
required_inputs = mode.inputs(template)

bundle = mode.compile(
    template,
    data={
        "Sheet1": {
            "customer_name": "Acme Industries",
            "invoice_number": 1024,
        },
    },
    bundle_path="report_bundle",
)

mode.export(
    bundle,
    "filled.xlsx",
    format="xlsx",
    column_width_mode="fixed",
    row_height_mode="fixed",
)
```

## API Surface

- `mode.extract(path)`
- `mode.inputs(template)`
- `mode.compile(template, data, bundle_path=None)`
- `mode.export(bundle_or_path, output_path, format="xlsx", **options)`

Top-level exports mirror the namespace:

- `extract_template`
- `get_template_inputs`
- `compile_report_bundle`
- `export_report_bundle`
- `mode`

## Workflow

```python
from mindoff_dataport import mode

template = mode.extract("invoice_template.xlsx")
inputs = mode.inputs(template)
bundle = mode.compile(template, {"Sheet1": {"customer_name": "Acme"}})
mode.export(bundle, "filled.xlsx")
```

`mode.compile(...)` returns an in-memory `ReportBundle`. When `bundle_path` is provided, the same artifact is also written as a directory containing:

- `manifest.json`: bundle version, inputs, sheet/page metadata, dataframe sources, assets, and output capabilities
- `report.json`: resolved scalar/static cells plus dataframe anchors
- `data/*.parquet`: dataframe sources when parquet support is installed
- `data/*.csv`: fallback dataframe sources without optional parquet dependencies
- `assets/*`: reserved for future image/logo payloads

## Export Options

`mode.export(..., format="xlsx")` supports the current sizing controls:

- `column_width_mode`: `"fixed"`, `"even"`, or `"hug"`
- `row_height_mode`: `"fixed"`, `"even"`, or `"hug"`
- `default_column_width`
- `default_row_height`
- `export_mode`: `"fidelity"` or `"streaming"`
- `max_rows_per_workbook`

`mode.export(..., format="pdf")` renders each workbook sheet as a styled report page
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
        "line_items": dataframe,
    },
    "region_sheet": {
        "North": {"region_name": "North"},
        "South": {"region_name": "South"},
    },
}
```

Supported placeholder types:

- Scalars: `string`, `number`, `int`, `float`, `date`, `boolean`
- Dataframes: `dataframe-headers`, `dataframe-content`

## Demo

```bash
python examples/xlsx_output.py
python examples/pdf_output.py
```

## Current Scope

- Template input: `.xlsx`
- Canonical intermediate: `ReportBundle` directory
- Production output: `.xlsx`, `.pdf`
- Reserved output interfaces: `image`
