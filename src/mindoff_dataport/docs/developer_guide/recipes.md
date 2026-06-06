# Recipes

Short, copy-ready solutions for the patterns that come up most. Each one is a complete, working snippet; adapt the sheet names and columns to your own template. Every recipe here also has a runnable version under the repository's `examples/` folder.

## Prerequisites

- `mindoff-dataport` installed with the Polars extra: `pip install "mindoff-dataport[polars]"`
- Template extracted with `mo_dataport.extract()` and placeholder keys confirmed with `mo_dataport.inputs(schema)`

## Implementation

### 1. Scalar Values and a Dataframe Table

The everyday report: a title, a date, and a table that grows to fit.

```python
import datetime as dt
import polars as pl
from mindoff_dataport import mo_dataport

schema = mo_dataport.extract("template.xlsx")
rows   = pl.scan_parquet("sales.parquet").select(["product", "units", "revenue"])

bundle = mo_dataport.compile(
    schema,
    {
        "Sales Summary": {
            "report_title": "Q1 2026 Sales",
            "generated_on": dt.date(2026, 4, 28),
            "sales_rows":   rows,
        }
    },
)
mo_dataport.export(bundle, "report.xlsx", export_mode="streaming")
```

### 2. Repeat Sections (per-customer blocks)

Stamp the same block once per record, stacked down the sheet.

Template cells:

```text
{{reports:repeat-start}}
Customer: {{customer_name:string}}
{{line_items:dataframe-header}}
{{line_items:dataframe-content}}
{{reports:repeat-end}}
```

Code:

```python
bundle = mo_dataport.compile(
    schema,
    {
        "Sheet1": {
            "reports": [
                {"customer_name": "Acme",   "line_items": acme_df},
                {"customer_name": "Globex", "line_items": globex_df},
            ]
        }
    },
)
mo_dataport.export(bundle, "combined.xlsx", export_mode="streaming")
mo_dataport.export(bundle, "combined.pdf",  format="pdf")
```

Repeat section rules:

- One or more **non-overlapping sibling** vertical sections per sheet.
- Static rows are allowed before, between, and after sections.
- Repeat keys must be unique per sheet.
- Merged cells are supported in fixed/static rows, but **not** over `dataframe-content` rows.
- No nested repeats.

### 3. Dynamic Sheets (one sheet per region)

One output tab per group, generated from a single template sheet named `{{region_sheet}}`.

```python
bundle = mo_dataport.compile(
    schema,
    {
        "region_sheet": {
            "North Sheet": {"region_name": "North", "owner": "Alice", "sales_rows": north_df},
            "South Sheet": {"region_name": "South", "owner": "Bob",   "sales_rows": south_df},
        }
    },
)
mo_dataport.export(bundle, "regions.xlsx", export_mode="streaming")
```

### 4. Dataframe Column Occupation and Alignment

Make columns span template columns and align headers and numbers independently.

```python
rows = pl.scan_parquet("data.parquet").select(["Employee Name", "Department", "Amount"])

bundle = mo_dataport.compile(
    schema,
    {
        "Column Layout": {
            "report_title": "Dataframe Column Occupation",
            "headers": rows,
            "rows": rows,
        }
    },
    dataframe_options={
        "Column Layout": {
            "headers": {
                "columns": {
                    "Employee Name": {"occupation": 2, "alignment": "center"},
                    "Department":    {"occupation": 2, "alignment": "center"},
                    "Amount":        {"occupation": 1, "alignment": "center"},
                }
            },
            "rows": {
                "columns": {
                    "Employee Name": {"occupation": 2, "alignment": "left"},
                    "Department":    {"occupation": 2, "alignment": "center"},
                    "Amount":        {"occupation": 1, "alignment": "right"},
                }
            },
        }
    },
)
mo_dataport.export(bundle, "column_layout.xlsx", export_mode="streaming")
mo_dataport.export(bundle, "column_layout.pdf", format="pdf", orientation="portrait", row_height_mode="fixed")
```

### 5. Discover Inputs Before Compiling

Don't guess the payload shape; ask the template.

```python
import pprint

schema = mo_dataport.extract("template.xlsx")
pprint.pp(mo_dataport.inputs(schema))
# {'Sales Summary': {'report_title': 'string', 'generated_on': 'date', 'sales_rows': 'dataframe'}}
```

### 6. Persist a Bundle for Later Re-Export

Compile once; export many times, from anywhere.

```python
bundle = mo_dataport.compile(schema, data, bundle_path="saved_bundle")

# Later, in a separate process or script:
mo_dataport.export("saved_bundle", "report.xlsx")
mo_dataport.export("saved_bundle", "report.pdf", format="pdf")
```

### 7. Split Large Exports Across Workbooks

Keep individual files manageable when row counts are huge.

```python
outputs = mo_dataport.export(
    bundle,
    "output.xlsx",
    export_mode="streaming",
    max_rows_per_workbook=500_000,   # split when a sheet exceeds this many rows
)
# outputs -> list[str] with a single `.zip` path when the export is split
```

### 8. PDF with Custom Fonts and Landscape Layout

```python
mo_dataport.export(
    bundle,
    "report.pdf",
    format="pdf",
    page_size="A4",
    orientation="landscape",
    margin=28,
    fonts={
        "Inter": {
            "regular":     "fonts/Inter-Regular.ttf",
            "bold":        "fonts/Inter-Bold.ttf",
            "italic":      "fonts/Inter-Italic.ttf",
            "bold_italic": "fonts/Inter-BoldItalic.ttf",
        }
    },
)
```

## Runnable Examples

Clone the repo and run any example directly. Each folder contains a `template.xlsx`, a `run.py`, and (where relevant) a `data.parquet`. Output files are written to `examples/<name>/output/` and are not tracked by git.

```bash
git clone https://github.com/mindoffwork/mindoff-dataport
cd mindoff-dataport
pip install -e ".[dev]"
python examples/<name>/run.py
```

| Example | What it shows |
|---|---|
| `basic/` | Minimal XLSX and PDF export from a parquet-backed template |
| `bundle_path/` | Compile to a persistent bundle directory, export later |
| `dataframe_options/` | Split `dataframe-header` / `dataframe-content` anchors with per-column occupation and alignment |
| `dataframe_shift/` | `dataframe_shift="both"`: a table expanding right and down inside repeat blocks |
| `dynamic_sheets/` | One output sheet per data group via `{{key:sheet-name}}` expansion |
| `input_discovery/` | Introspect required template inputs before building a payload |
| `repeat_block/` | One repeat block per customer: per-block scalars and dataframes |
| `repeat_dataframe_headers/` | `repeat_dataframe_headers=True`: repeat headers across paginated PDF blocks |
| `split_workbooks_streaming/` | `max_rows_per_workbook`: split large exports across multiple workbooks |
| `style_showcase/` | Full style coverage exported via openpyxl, xlsxwriter, and PDF |
| `validation_errors/` | How validation errors surface before any file is written |
| `benchmark/` | Runtime and memory benchmarks vs. raw openpyxl / xlsxwriter / ReportLab |
