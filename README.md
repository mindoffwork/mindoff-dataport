# Mindoff Data Export

Extract Excel templates, fill typed placeholders, and rebuild `.xlsx` files with layout and styling preserved.

## What This Library Does

- Extract an Excel workbook into a JSON-like schema
- Detect typed template inputs from placeholders (`{{key:type}}`)
- Render runtime data into the schema safely
- Build final `.xlsx` outputs with styling, merges, formulas, and sizing behavior

## Install

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -e .
```

Optional (for demo and DataFrame examples):

```bash
pip install polars
```

## Quick Start

```python
from mindoff_data_export import (
    extract_template,
    get_template_inputs,
    build_template_with_data,
)

schema = extract_template("template.xlsx")
required_inputs = get_template_inputs(schema)

build_template_with_data(
    schema,
    data={
        "customer_name": "Acme Industries",
        "invoice_number": 1024,
    },
    output_path="filled.xlsx",
    column_width_mode="fixed",
    row_height_mode="fixed",
)
```

## Documentation

### Available Endpoints

Public functions exposed by this package:

- `extract_template`
- `build_template_with_data`
- `get_template_inputs`
- `render_schema`

### 1) `extract_template`

What it is: Reads an `.xlsx` file and converts it into a `WorkbookSchema` dictionary.

Why it exists: This is the entry point for turning a designer-authored Excel file into a portable, editable template schema.

Import and call:

```python
from mindoff_data_export import extract_template

schema = extract_template("invoice_template.xlsx")
```

| Parameter | Accepted Input (Type) | Required | Description |
|---|---|---|---|
| `path` | `str` (path to `.xlsx`) | Required | File path of the Excel template to extract. |

Real-world usage example:

```python
from mindoff_data_export import extract_template

# Extract and persist a schema version for later rendering/building.
schema = extract_template("templates/monthly_sales_report.xlsx")
```

### 2) `build_template_with_data`

What it is: Validates input data, renders placeholders into schema, and builds output workbook(s) in one call.

Why it exists: This is the main production API for template-to-export workflows.

Import and call:

```python
from mindoff_data_export import build_template_with_data

build_template_with_data(schema, data, "filled.xlsx")
```

| Parameter | Accepted Input (Type) | Required | Description |
|---|---|---|---|
| `schema` | `WorkbookSchema` (`dict`) | Required | Template schema containing placeholders. |
| `data` | `dict[str, Any]` | Required | Runtime values mapped by placeholder key. |
| `output_path` | `str` | Required | Output file path; in streaming mode, used as base name for chunked files. |
| `column_width_mode` | `str \| None` (`"fixed"`, `"even"`, `"hug"`) | Optional | Global override for column sizing. |
| `row_height_mode` | `str \| None` (`"fixed"`, `"even"`, `"hug"`) | Optional | Global override for row sizing. |
| `default_column_width` | `float \| None` | Optional | Default even width override. |
| `default_row_height` | `float \| None` | Optional | Default even height override. |
| `export_mode` | `Literal["fidelity", "streaming"]` | Optional (default: `"fidelity"`) | `"fidelity"` writes a single workbook preserving merges/styles. `"streaming"` is optimized for large dataframe-content exports. |
| `streaming_chunk_rows` | `int` | Optional (default: `50000`) | Rows per batch in streaming mode. |
| `max_rows_per_workbook` | `int` | Optional (default: `1048576`) | Workbook row cap in streaming mode before splitting into `*.partNNN.xlsx`. |

Real-world usage example:

```python
from mindoff_data_export import extract_template, build_template_with_data

schema = extract_template("templates/customer_statement.xlsx")

result = build_template_with_data(
    schema=schema,
    data={
        "customer_name": "Acme Industries",
        "statement_date": "2026-04-22",
        "balance_due": 12450.75,
    },
    output_path="exports/acme_statement_apr_2026.xlsx",
    export_mode="fidelity",
    column_width_mode="fixed",
    row_height_mode="fixed",
)

# Fidelity mode returns None.
assert result is None
```

### 3) `get_template_inputs`

What it is: Scans schema placeholders and returns required inputs as `{key: type}`.

Why it exists: Lets you validate and build data contracts before rendering/export.

Import and call:

```python
from mindoff_data_export import get_template_inputs

inputs = get_template_inputs(schema)
```

| Parameter | Accepted Input (Type) | Required | Description |
|---|---|---|---|
| `schema` | `WorkbookSchema` (`dict`) | Required | Template schema to inspect for `{{key:type}}` placeholders. |

Real-world usage example:

```python
from mindoff_data_export import extract_template, get_template_inputs

schema = extract_template("templates/invoice_template.xlsx")
required_inputs = get_template_inputs(schema)

print(required_inputs)
# Example:
# {
#   "invoice_number": "number",
#   "invoice_date": "date",
#   "customer_name": "string"
# }
```

### 4) `render_schema`

What it is: Validates runtime data and returns a new schema with placeholders resolved.

Why it exists: Use this when you want to inspect or transform rendered schema before writing the final workbook.

Import and call:

```python
from mindoff_data_export import render_schema

resolved_schema = render_schema(schema, data)
```

| Parameter | Accepted Input (Type) | Required | Description |
|---|---|---|---|
| `schema` | `WorkbookSchema` (`dict`) | Required | Template schema containing placeholders. |
| `data` | `dict[str, Any]` | Required | Runtime values used to replace placeholders. |

Real-world usage example:

```python
from mindoff_data_export import extract_template, render_schema, build_template_with_data

schema = extract_template("templates/quote_template.xlsx")
resolved_schema = render_schema(
    schema,
    {
        "quote_id": 22019,
        "client_name": "Contoso Retail",
        "valid_until": "2026-05-15",
    },
)

# Optional inspection/transformation point before file output.
build_template_with_data(resolved_schema, {}, "exports/quote_22019.xlsx")
```

## Demo

```bash
python examples/demo.py
```

## Current Scope

- Input template: `.xlsx`
- Output export: `.xlsx`
- Public API: `extract_template`, `build_template_with_data`, `get_template_inputs`, `render_schema`
