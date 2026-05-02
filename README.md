<h1>Mindoff Dataport</h1>

_Build high-fidelity Excel and PDF reports from reusable `.xlsx` templates._

Mindoff Dataport turns styled Excel workbooks into reusable report templates, compiles runtime data into a portable `ReportBundle`, and exports production-ready `.xlsx` and `.pdf` outputs while preserving layout, structure, and visual fidelity.

[![Coverage Status](https://codecov.io/gh/mindoffwork/mindoff-dataport/branch/root/graph/badge.svg)](https://codecov.io/gh/mindoffwork/mindoff-dataport)
[![PyPI version](https://img.shields.io/pypi/v/mindoff-dataport.svg?logo=pypi&logoColor=white)](https://pypi.org/project/mindoff-dataport/)
[![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-3776AB?logo=python&logoColor=white)](https://github.com/mindoffwork/mindoff-dataport/actions/workflows/root_ci.yml)

**Source**: [https://github.com/mindoffwork/mindoff-dataport](https://github.com/mindoffwork/mindoff-dataport)

## Key Features

1. **Template-First Report Generation**  
   Turn real Excel workbooks into reusable report templates without rebuilding layouts in code.

2. **Compile Once. Export Natively to XLSX and PDF.**  
   Build reports once and export polished `.xlsx` and `.pdf` outputs from the same source with consistent fidelity.

3. **Dataframes Plug Directly Into Templates**  
   Connect dataframe inputs directly to templates so report generation fits naturally into modern data workflows.

4. **Built for Large Exports Without Memory Bloat**  
   Export large datasets with confidence, without turning memory usage into a bottleneck.

5. **Flexible Repeating and Dynamic Sheets**  
   Generate repeated sections and dynamic sheets for customer-wise, region-wise, or report-wise output from a single template system.

6. **Runtime Layout Control Without Template Rework**  
   Fine-tune output layout programmatically without redesigning the original workbook.

## Documentation
### Table of Contents

1. [Purpose](#1-purpose)
2. [Install](#2-install)
3. [Quick Start](#3-quick-start)
4. [Core Concepts](#4-core-concepts)
5. [API Reference](#5-api-reference)
6. [Template Placeholders](#6-template-placeholders)
7. [Data Contract](#7-data-contract)
8. [Export Options](#8-export-options)
9. [Dataframe Column Layout](#9-dataframe-column-layout)
10. [Sizing Options](#10-sizing-options)
11. [Supported Styling](#11-supported-styling)
12. [Custom Fonts for PDF](#12-custom-fonts-for-pdf)
13. [ReportBundle Directory](#13-reportbundle-directory)
14. [Recipes](#14-recipes)
15. [Current Scope](#15-current-scope)
16. [License](#16-license)

## 1. Purpose

Mindoff Dataport is built to turn Excel-based report designs into reusable, data-driven outputs with a template format that is convenient to create, review, and maintain.

- Reuse existing Excel report layouts instead of rebuilding them from scratch in code.
- Fill those layouts with live business data and keep the final output polished and presentation-ready.
- Generate both Excel and PDF from the same report source, so teams do not maintain separate reporting flows.
- Scale one template into many outputs, whether that means repeated sections, multiple sheets, or report variants for different audiences.
- Support larger exports more reliably as report volume grows.

## 2. Install

```bash
pip install mindoff-dataport
```

For dataframe support (required when passing Polars DataFrames or LazyFrames):

```bash
pip install "mindoff-dataport[polars]"
```

## 3. Quick Start

```python
import polars as pl
from mindoff_dataport import mo_dataport

# 1. Extract the template
template = mo_dataport.extract("invoice_template.xlsx")

# 2. Inspect what the template requires
required_inputs = mo_dataport.inputs(template)
# {'Invoice': {'customer_name': 'string', 'invoice_number': 'number', 'line_items': 'dataframe'}}

# 3. Compile: bind data to the template
polars_dataframe = pl.DataFrame(
    {
        "item": ["Widget A", "Widget B"],
        "amount": [125, 275],
    }
)

bundle = mo_dataport.compile(
    template,
    data={
        "Invoice": {
            "customer_name": "Acme Industries",
            "invoice_number": 1024,
            "line_items": polars_dataframe,
        }
    },
)

# 4. Export to XLSX
mo_dataport.export(bundle, "invoice_filled.xlsx")

# 4b. Export to PDF
mo_dataport.export(bundle, "invoice_filled.pdf", format="pdf")
```

## 4. Core Concepts

### Workflow

```
.xlsx template  ──extract()──►  WorkbookSchema
                                     │
                              compile(schema, data)
                                     │
                                     ▼
                              ReportBundle (directory)
                              ├── manifest.json
                              ├── report.json
                              └── data/*.parquet
                                     │
                            export(bundle, path, format=…)
                                     │
                             ┌───────┴───────┐
                          .xlsx           .pdf
```

### Import Alias

The recommended entrypoint is:

```python
from mindoff_dataport import mo_dataport
```

All four public functions are also importable at the top level:

```python
from mindoff_dataport import (
    extract_template,
    get_template_inputs,
    compile_report_bundle,
    export_report_bundle,
)
```

`mo_dataport.extract` / `mo_dataport.inputs` / `mo_dataport.compile` / `mo_dataport.export` are short aliases for the same functions.

## 5. API Reference

### `extract(path)` — `extract_template(path)`

Reads an `.xlsx` file and returns a `WorkbookSchema` containing cell styles, dimensions, merged regions, manual print breaks, and discovered placeholder types.

| Parameter | Type  | Required | Description                          |
|-----------|-------|----------|--------------------------------------|
| `path`    | `str` | Yes      | Path to the `.xlsx` template file    |

**Returns:** `WorkbookSchema`

### `inputs(schema)` — `get_template_inputs(schema)`

Inspects the schema and returns a sheet-scoped dictionary of all inputs the template requires, keyed by sheet name and then by placeholder key.

| Parameter | Type            | Required | Description                              |
|-----------|-----------------|----------|------------------------------------------|
| `schema`  | `WorkbookSchema`| Yes      | Schema produced by `extract()`           |

**Returns:** `dict[str, dict[str, str | list]]`

Example output:

```python
{
    "Sales Summary": {
        "report_title": "string",
        "generated_on": "date",
        "sales_rows": "dataframe",
    }
}
```

### `compile(template, data, bundle_path=None, dataframe_options=None, dataframe_shift="both")` - `compile_report_bundle(...)`

Binds runtime data to the template, validates all inputs against the sheet contract, materialises Polars DataFrames / LazyFrames to Parquet, and produces a `ReportBundle`.

| Parameter           | Type                     | Required | Description                                                                                 |
|---------------------|--------------------------|----------|---------------------------------------------------------------------------------------------|
| `template`          | `WorkbookSchema`         | Yes      | Schema from `extract()`                                                                     |
| `data`              | `dict[str, Any]`         | Yes      | Sheet-scoped payload. See [Data Contract](#data-contract)                                   |
| `bundle_path`       | `str \| None`            | No       | If provided, writes the bundle as a directory at this path. Omit for in-memory only         |
| `dataframe_options` | `dict[str, Any] \| None` | No       | Per-sheet, per-placeholder dataframe layout overrides. See [Dataframe Column Layout](#dataframe-column-layout) |
| `dataframe_shift`   | `str`                    | No       | How normal-sheet template cells/merges move around dataframe output: `"both"`, `"horizontal"`, `"vertical"`, or `"none"` |

**Returns:** `ReportBundle`

**Raises:** `KeyError` if a required placeholder key is missing from the payload.

### `export(bundle_or_path, output_path, format="xlsx", **options)` — `export_report_bundle(...)`

Renders the bundle to a file. Accepts an in-memory `ReportBundle` or a path to a persisted bundle directory.

| Parameter        | Type                  | Required | Default   | Description                                                                     |
|------------------|-----------------------|----------|-----------|---------------------------------------------------------------------------------|
| `bundle_or_path` | `ReportBundle \| str` | Yes      | —         | In-memory bundle or path to a bundle directory                                  |
| `output_path`    | `str`                 | Yes      | —         | Destination file path (`.xlsx` or `.pdf`)                                       |
| `format`         | `str`                 | No       | `"xlsx"`  | Output format: `"xlsx"`, `"pdf"`. (`"image"` is reserved; raises `NotImplementedError`) |
| `**options`      | —                     | No       | —         | Sizing and format-specific options. See [Export Options](#export-options)       |

**Returns:** `None` for `"fidelity"` XLSX and all PDF exports. `list[str]` for `"streaming"` XLSX: one workbook path when no split is needed, or one `.zip` path when the export is split across workbooks.

## 6. Template Placeholders

Mark cells in your `.xlsx` template using the `{{key:type}}` syntax. The extractor reads these markers and builds the input contract.

```
{{report_title:string}}
{{invoice_number:number}}
{{generated_on:date}}
{{line_items:dataframe}}
{{line_items:dataframe-header}}
{{line_items:dataframe-content}}
{{reports:repeat-start}}
  ...
{{reports:repeat-end}}
```

### Placeholder Types

#### Scalar Types

| Type      | Accepted Python values                                          |
|-----------|-----------------------------------------------------------------|
| `string`  | `str`                                                           |
| `number`  | `int`, `float`                                                  |
| `int`     | `int`                                                           |
| `float`   | `float`                                                         |
| `date`    | `datetime.date`, `datetime.datetime`                            |
| `boolean` | `bool`                                                          |

The placeholder cell is replaced in-place with the supplied value, inheriting all cell styles from the template.

#### Dataframe Types

| Type                 | What it writes                                             | Typical use                                              |
|----------------------|------------------------------------------------------------|----------------------------------------------------------|
| `dataframe`          | Headers on the anchor row, content starting the next row   | All-in-one table drop-in                                 |
| `dataframe-header`   | Column headers only, on the anchor row                     | Styled header row defined separately from content        |
| `dataframe-content`  | Data rows only, starting at the anchor row                 | Content area below a separately-styled header            |

The anchor cell inherits its style (font, fill, border, alignment) and applies it to all generated cells. Column names become header text.

**Streaming note:** `dataframe-content` placeholders support streaming from Parquet. Only one `dataframe-content` placeholder is allowed per non-repeat sheet in streaming mode.

### Manual Page Breaks

Templates may also contain manual Excel print breaks.

- `row_page_breaks`: 1-based template row indexes after which a new printed page begins
- `column_page_breaks`: 1-based template column indexes after which Excel starts a new printed page

These are extracted from Excel's manual print-break metadata, not placeholder syntax.

- During `compile()`, breaks are resolved against the rendered layout after dataframe expansion and `dataframe_shift`
- PDF uses resolved row breaks only, inserting a new PDF page before later rows
- XLSX preserves both resolved row and column breaks in fidelity and streaming exports

#### Repeat Types

Used in pairs to define a block that is rendered once per record in an ordered list payload.

| Type            | Description                                              |
|-----------------|----------------------------------------------------------|
| `repeat-start`  | Marks the first row of the repeating block (control row, not rendered) |
| `repeat-end`    | Marks the last row of the repeating block (control row, not rendered)  |

See [Repeat Sections](#repeat-sections) for usage.

## 7. Data Contract

Payloads are **sheet-scoped**. The top-level key must match the sheet name in the template.

### Static Sheet

```python
{
    "Invoice": {
        "customer_name": "Acme Industries",   # string
        "invoice_number": 1024,               # number
        "due_date": datetime.date(2026, 5, 1),# date
        "line_items": polars_dataframe,       # dataframe / LazyFrame
    }
}
```

### Dynamic Sheet Group

When a template sheet name is exactly `{{key}}`, it becomes a template for multiple output sheets. Pass a dict of `output_sheet_name -> payload` keyed under that placeholder key.

```python
{
    "region_sheet": {                          # sheet-name placeholder key
        "North Sheet": {                       # → output sheet name
            "region_name": "North",
            "owner": "Alice",
            "sales_rows": north_df,
        },
        "South Sheet": {
            "region_name": "South",
            "owner": "Bob",
            "sales_rows": south_df,
        },
    }
}
```

Output sheet order follows the payload dict insertion order.

`inputs(schema)` reports dynamic sheet groups under the same placeholder key:

```python
{
    "region_sheet": {
        "*": {
            "region_name": "string",
            "owner": "string",
            "sales_rows": "dataframe",
        }
    }
}
```

### Repeat Section

```python
{
    "Sheet1": {
        "reports": [                           # key must match repeat-start/end key
            {"customer_name": "Acme", "line_items": acme_df},
            {"customer_name": "Globex", "line_items": globex_df},
        ]
    }
}
```

### Using Polars LazyFrames (Recommended for Large Data)

```python
import polars as pl

rows = pl.scan_parquet("sales.parquet").select(["product", "units", "revenue"])

bundle = mo_dataport.compile(schema, {"Sheet1": {"sales_rows": rows}})
```

Polars `LazyFrame` inputs remain disk-backed until export time; rows are never fully materialised in memory.

## 8. Export Options

All options are passed as keyword arguments to `export()`.

## 9. Dataframe Column Layout

Use `dataframe_options` during `compile()` to control how dataframe columns occupy template columns and to override horizontal alignment per generated column.

The structure is:

```python
dataframe_options = {
    "Sheet Name": {
        "placeholder_key": {
            "columns": {
                "Column Name": {"occupation": 2, "alignment": "left"},
            }
        }
    }
}
```

For templates that split headers and rows across separate placeholders, configure each placeholder independently:

```python
dataframe_options = {
    "Column Layout": {
        "headers": {
            "columns": {
                "Employee Name": {"occupation": 2, "alignment": "center"},
                "Department": {"occupation": 2, "alignment": "center"},
                "Amount": {"occupation": 1, "alignment": "center"},
            }
        },
        "rows": {
            "columns": {
                "Employee Name": {"occupation": 2, "alignment": "left"},
                "Department": {"occupation": 2, "alignment": "center"},
                "Amount": {"occupation": 1, "alignment": "right"},
            }
        },
    }
}
```

Rules:

- `occupation` must be a positive integer
- `alignment` must be one of `"left"`, `"center"`, or `"right"`
- Options are keyed by resolved output sheet name, then placeholder key
- Unconfigured dataframe columns default to `occupation=1` and keep the template cell alignment

### Dataframe Collision Shifting

When dataframe output expands into adjacent template space, `compile()` can move normal-sheet template cells and merged regions out of the dataframe range before XLSX or PDF export.

```python
bundle = mo_dataport.compile(
    schema,
    data,
    dataframe_shift="both",  # "both", "horizontal", "vertical", or "none"
)
```

| Mode           | Behavior                                                                       |
|----------------|--------------------------------------------------------------------------------|
| `"both"`       | Shift right-side cells/merges horizontally and lower cells/merges vertically   |
| `"horizontal"` | Shift only cells/merges to the right of dataframe output                       |
| `"vertical"`   | Shift only cells/merges below dataframe output                                 |
| `"none"`       | Do not shift; template merges that overlap dataframe output raise `ValueError` |

The shift is metadata-only: dataframe rows remain in Parquet, `report.json` stores compact anchors, and streaming export still reads rows in batches. The same shifted bundle layout is used by XLSX and PDF. Repeat sections keep their stricter merge rules.

See `examples/dataframe_shift/xlsx.py` and `examples/dataframe_shift/pdf.py`.

### Manual Page Breaks

Excel manual print breaks from the template are extracted into schema metadata and resolved again after compile-time dataframe expansion.

- `row_page_breaks` start a new printed page after the given 1-based template row
- `column_page_breaks` start a new printed page after the given 1-based template column in XLSX output
- PDF uses resolved row breaks as manual page boundaries and ignores column breaks

See `examples/page_break/xlsx.py` and `examples/page_break/pdf.py`.

### XLSX Options

| Option                  | Type    | Default       | Description                                                                              |
|-------------------------|---------|---------------|------------------------------------------------------------------------------------------|
| `export_mode`           | `str`   | `"fidelity"`  | `"fidelity"`: full in-memory render (supports all features). `"streaming"`: row-by-row write (lower memory, limited features — see constraints below) |
| `column_width_mode`     | `str`   | schema value  | `"fixed"`, `"even"`, or `"hug"`. Overrides the value stored in the template schema      |
| `row_height_mode`       | `str`   | schema value  | `"fixed"`, `"even"`, or `"hug"`. Overrides the value stored in the template schema      |
| `default_column_width`  | `float` | schema value  | Fallback column width in Excel character units when mode is `"even"` or no width stored  |
| `default_row_height`    | `float` | schema value  | Fallback row height in points when mode is `"even"` or no height stored                  |
| `streaming_chunk_rows`  | `int`   | `50000`       | Number of Parquet rows read per batch during streaming                                   |
| `max_rows_per_workbook` | `int`   | `1048576`     | Split output into multiple `.xlsx` parts when this row limit is reached                  |
| `auto_delete_bundle`    | `bool`  | `False`       | Delete the bundle directory after a successful export                                    |

**Streaming mode constraints:**

- No `hug` sizing
- No merged cells may remain intersecting `dataframe-content` output rows after compile-time `dataframe_shift`
- Only one `dataframe-content` placeholder per non-repeat sheet

**Split output:** When `max_rows_per_workbook` is exceeded in streaming mode, `export()` writes workbook parts, bundles them into `output.zip`, deletes the individual part files, and returns a one-item `list[str]` containing the zip path.

### PDF Options

PDF-specific options are passed as keyword arguments alongside sizing options.

| Option                  | Type              | Default       | Description                                                          |
|-------------------------|-------------------|---------------|----------------------------------------------------------------------|
| `page_size`             | `str`             | `"A4"`        | Paper size: `"A4"`, `"LETTER"`, or `"LEGAL"`                        |
| `orientation`           | `str`             | `"portrait"`  | Page orientation: `"portrait"` or `"landscape"`                      |
| `margin`                | `float`           | `36`          | Page margin in points (≥ 0). Applied equally on all four sides       |
| `streaming_chunk_rows`  | `int`             | `50000`       | Rows read per batch for `dataframe-content` and repeat sections      |
| `fonts`                 | `dict \| None`    | `None`        | Custom TrueType / OpenType font families. See [Custom Fonts for PDF](#custom-fonts-for-pdf) |
| `column_width_mode`     | `str`             | schema value  | Same as XLSX. For sheets with `dataframe-content`, PDF supports `"fixed"` and `"even"` only |
| `row_height_mode`       | `str`             | schema value  | Same as XLSX. PDF also supports `"hug"` for `dataframe-content` row height |
| `default_column_width`  | `float`           | schema value  | Same as XLSX                                                         |
| `default_row_height`    | `float`           | schema value  | Same as XLSX                                                         |

> `export_mode` is ignored for PDF; PDF always paginates automatically.

## 10. Sizing Options

Sizing modes control how column widths and row heights are computed at render time.

### Column Width Modes

| Mode      | Source                                                     | Limitation                              |
|-----------|------------------------------------------------------------|-----------------------------------------|
| `"fixed"` | Reads widths stored in the template schema per column      | Requires widths to be set in the template |
| `"even"`  | Applies `default_column_width` uniformly to all columns    | Ignores per-column template widths       |
| `"hug"`   | Computes width from cell content at render time            | Not available in streaming mode          |

For PDF sheets that render `dataframe-content`, `column_width_mode="hug"` is not supported because it would require buffering all rows before sizing.

### Row Height Modes

| Mode      | Source                                                     | Limitation                              |
|-----------|------------------------------------------------------------|-----------------------------------------|
| `"fixed"` | Reads heights stored in the template schema per row        | Requires heights to be set in the template |
| `"even"`  | Applies `default_row_height` uniformly to all rows         | Ignores per-row template heights         |
| `"hug"`   | Auto-fits row height to content                            | Not available in streaming mode          |

For PDF sheets that render `dataframe-content`, `row_height_mode="hug"` is supported and auto-sizes each streamed row chunk.

### Width and Height Units

| Parameter              | Unit                        | Default in schema |
|------------------------|-----------------------------|-------------------|
| `default_column_width` | Excel character units        | `15.0`            |
| `default_row_height`   | Points                       | `15.0`            |
| `margin` (PDF)         | Points (1pt = 1/72 inch)     | `36`              |

Kwargs passed to `export()` override values stored in the template schema.

## 11. Supported Styling

Styles are defined in the `.xlsx` template itself. The library extracts them during `extract()` and reapplies them faithfully at export time. No runtime style configuration is needed.

### Font Properties

| Property    | Values / Range                            | Notes                                          |
|-------------|-------------------------------------------|------------------------------------------------|
| `name`      | Any font family name                      | Falls back to Helvetica in PDF if not registered as a custom font |
| `size`      | `float` (points)                          | Default `11.0`                                 |
| `bold`      | `True` / `False`                          |                                                |
| `italic`    | `True` / `False`                          |                                                |
| `underline` | `"single"`, `"double"`, `None`            | Rendered in PDF via `<u>` markup               |
| `color`     | Hex ARGB string or `theme:<index>:<tint>` | PDF falls back to the default Office theme palette for theme colors |

### Fill Properties

| Property    | Values                  | Notes                                       |
|-------------|-------------------------|---------------------------------------------|
| `bg_color`  | Hex ARGB string, `theme:<index>:<tint>`, or None | Solid fills only (`fgColor` in openpyxl)    |

Patterned fills are not extracted or rendered.

### Alignment Properties

| Property      | Values                                              |
|---------------|-----------------------------------------------------|
| `horizontal`  | `"left"`, `"center"`, `"right"`, `"centerContinuous"` |
| `vertical`    | `"top"`, `"center"`, `"bottom"`                     |
| `wrap_text`   | `True` / `False`                                    |

In PDF output, newline characters render as line breaks only when `wrap_text=True`; otherwise they are flattened to spaces.

### Border Properties

Each cell has four border sides: `top`, `bottom`, `left`, `right`. Each side has a `style` and optional `color`.

| Border Style | Rendered Width (PDF points) |
|--------------|-----------------------------|
| `hair`       | 0.25                        |
| `thin`       | 0.5                         |
| `medium`     | 1.0                         |
| `thick`      | 1.5                         |
| `dashed`     | 0.75                        |
| `dotted`     | 0.5                         |
| `double`     | 1.25                        |

Borders on merged cells are drawn around the **full merged region**, not only the anchor cell.

### Merged Cells

Merged regions are extracted from the template and preserved in both XLSX and PDF output. During XLSX fidelity export, the full merged region is re-applied. During PDF export, merged cells are rendered as `SPAN` table commands.

### Sheet Gridlines

The template's `show_gridlines` property is preserved in XLSX output.

## 12. Custom Fonts for PDF

By default the PDF renderer maps all cell fonts to ReportLab's built-in **Helvetica** family. To use your own TrueType or OpenType fonts, pass a `fonts` dict to `export()`.

### Shorthand — Regular Only

Provide a single file path when you only have a regular weight:

```python
mo_dataport.export(
    bundle,
    "report.pdf",
    format="pdf",
    fonts={
        "Inter": "/path/to/fonts/Inter-Regular.ttf",
    },
)
```

Any cell whose template font name is `"Inter"` will use this file. Bold and italic variants fall back to the regular file.

### Full Variant Map

Provide a dict with `regular`, `bold`, `italic`, and `bold_italic` keys to enable distinct variants:

```python
mo_dataport.export(
    bundle,
    "report.pdf",
    format="pdf",
    fonts={
        "Inter": {
            "regular":     "/path/to/fonts/Inter-Regular.ttf",
            "bold":        "/path/to/fonts/Inter-Bold.ttf",
            "italic":      "/path/to/fonts/Inter-Italic.ttf",
            "bold_italic": "/path/to/fonts/Inter-BoldItalic.ttf",
        }
    },
)
```

### Font Config Reference

| Key           | Required | Description                                              |
|---------------|----------|----------------------------------------------------------|
| `regular`     | Yes      | Path to the regular (normal weight, upright) font file   |
| `bold`        | No       | Path to the bold variant; falls back to `regular` if absent |
| `italic`      | No       | Path to the italic variant; falls back to `regular` if absent |
| `bold_italic` | No       | Path to bold-italic; falls back to `bold` then `regular` |

### Matching Behaviour

The renderer matches the `font.name` stored in the template cell against the keys in the `fonts` dict (case-sensitive). If no match is found, Helvetica is used. Multiple font families can be registered in one call:

```python
fonts={
    "Inter": {...},
    "Roboto Mono": "/path/to/RobotoMono-Regular.ttf",
}
```

### Requirements and Errors

- Font files must exist on disk at the time `export()` is called; a missing file raises `ValueError`
- Each family **must** supply a `regular` file; omitting it raises `ValueError`
- Font files are registered with ReportLab once per process; re-registering the same path is a no-op

## 13. ReportBundle Directory

When `bundle_path` is passed to `compile()`, the bundle is persisted as a directory. The same directory can be re-loaded and re-exported without rerunning `compile()`.

```
report_bundle/
├── manifest.json      # bundle version, inputs, sheet metadata, dataframe sources, capabilities
├── report.json        # resolved scalar cells and dataframe anchor/repeat plans
└── data/
    └── *.parquet      # dataframe sources materialised from Polars inputs
```

> `report.json` stores dataframe **anchors** (column names, start row/column, style), not the expanded row data. Rows stay in Parquet and are read at export time.

Loading a persisted bundle:

```python
mo_dataport.export("report_bundle/", "output.xlsx")
# or load manually:
from mindoff_dataport import ReportBundle
bundle = ReportBundle.load("report_bundle/")
```

Setting `auto_delete_bundle=True` in `export()` deletes the bundle directory after a successful export.

## 14. Recipes

### Scalar Values + Dataframe Table

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

### Repeat Sections (per-customer invoice blocks)

Template cells:

```
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

Repeat section constraints:

- One or more **non-overlapping sibling** vertical sections per sheet
- Static rows are allowed before, between, and after sections
- Repeat keys must be unique per sheet
- Merged cells are supported in fixed/static rows, but **not** over `dataframe-content` rows
- No nested repeats

### Dynamic Sheets (one sheet per region)

```python
bundle = mo_dataport.compile(
    schema,
    {
        "region_sheet": {           # sheet-name placeholder key
            "North Sheet": {"region_name": "North", "owner": "Alice", "sales_rows": north_df},
            "South Sheet": {"region_name": "South", "owner": "Bob",   "sales_rows": south_df},
        }
    },
)
mo_dataport.export(bundle, "regions.xlsx", export_mode="streaming")
```

### Dataframe Column Occupation and Alignment

```python
rows = pl.scan_parquet("data.parquet").select(
    ["Employee Name", "Department", "Amount"]
)

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
                    "Department": {"occupation": 2, "alignment": "center"},
                    "Amount": {"occupation": 1, "alignment": "center"},
                }
            },
            "rows": {
                "columns": {
                    "Employee Name": {"occupation": 2, "alignment": "left"},
                    "Department": {"occupation": 2, "alignment": "center"},
                    "Amount": {"occupation": 1, "alignment": "right"},
                }
            },
        }
    },
)
mo_dataport.export(bundle, "column_layout.xlsx", export_mode="streaming")
mo_dataport.export(
    bundle,
    "column_layout.pdf",
    format="pdf",
    orientation="portrait",
    row_height_mode="fixed",
)
```

See `examples/dataframe_column_layout/xlsx.py` and `examples/dataframe_column_layout/pdf.py`.

### Discover Inputs Before Compiling

```python
schema = mo_dataport.extract("template.xlsx")
import pprint
pprint.pp(mo_dataport.inputs(schema))
# {'Sales Summary': {'report_title': 'string', 'generated_on': 'date', 'sales_rows': 'dataframe'}}
```

### Persist Bundle for Later Re-Export

```python
bundle = mo_dataport.compile(schema, data, bundle_path="saved_bundle")

# Later in a separate process or script:
mo_dataport.export("saved_bundle", "report.xlsx")
mo_dataport.export("saved_bundle", "report.pdf", format="pdf")
```

### Split Large Exports Across Workbooks

```python
outputs = mo_dataport.export(
    bundle,
    "output.xlsx",
    export_mode="streaming",
    max_rows_per_workbook=500_000,  # split when a sheet exceeds this row count
)
# outputs -> list[str] with a single `.zip` path when the export is split
```

### PDF with Custom Fonts and Landscape Layout

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

## 15. Current Scope

| Feature                  | Status                             |
|--------------------------|------------------------------------|
| Template input           | `.xlsx`                            |
| Canonical intermediate   | `ReportBundle` directory           |
| XLSX export (fidelity)   | Supported                          |
| XLSX export (streaming)  | Supported                          |
| PDF export               | Supported (ReportLab)              |
| Image export             | Reserved — raises `NotImplementedError` in v1 |
| Nested repeat sections   | Not supported in v1                |
| Patterned fills          | Not extracted or rendered          |

## 16. License

Released under the [MIT License](https://github.com/mindoffwork/mindoff-dataport/blob/root/LICENSE).
