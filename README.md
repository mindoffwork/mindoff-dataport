# Mindoff Data Export

Extract Excel templates, fill typed placeholders, and rebuild `.xlsx` files with layout and styling preserved.

Primary entrypoint: `from mindoff_data_export import mode`

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

## Quick Start (Recommended Entrypoint)

```python
from mindoff_data_export import mode

schema = mode.extract("template.xlsx")
required_inputs = mode.get_inputs(schema)

mode.build(
    schema,
    data={
        "Sheet1": {
            "customer_name": "Acme Industries",
            "invoice_number": 1024,
        },
    },
    output_path="filled.xlsx",
    column_width_mode="fixed",
    row_height_mode="fixed",
)
```

## Documentation

### API Surface

The package exposes a single entrypoint namespace:

- `mode.extract(path)`
- `mode.get_inputs(schema)`
- `mode.alter_schema(schema, data)`
- `mode.build(schema, data, output_path, **options)`

### Which Function To Use

| Function | What it provides | When to use |
|---|---|---|
| `mode.extract(path)` | `WorkbookSchema` parsed from an Excel template | Start of the pipeline when your source is an `.xlsx` template file. |
| `mode.get_inputs(schema)` | Required dynamic input contract in sheet-scoped form | You do not know all placeholder keys yet and need the template to tell you which inputs are required. |
| `mode.alter_schema(schema, data)` | A new rendered schema with placeholders resolved | You need an intermediate artifact for review, QA, transformation, or approval before file output. |
| `mode.build(schema, data, output_path, **options)` | Final workbook output files on disk | You are ready to produce exports from schema + runtime data in one call. |

If you already know the required keys and types, you can construct `data` directly and call `mode.build(...)` without `mode.get_inputs(...)`.

### Recommended Workflow

```python
from mindoff_data_export import mode

schema = mode.extract("invoice_template.xlsx")
inputs = mode.get_inputs(schema)
resolved = mode.alter_schema(schema, {"Sheet1": {"customer_name": "Acme"}})
mode.build(resolved, {"Sheet1": {}}, "filled.xlsx")
```

### Method Reference

### 1) `mode.extract`

Converts a designer-authored `.xlsx` template into a portable `WorkbookSchema` that serves as the source-of-truth contract for downstream validation, rendering, and export operations.

Provides: `WorkbookSchema`

Use when: You need to convert an Excel template into the schema object used by all other `mode` functions.

```python
from mindoff_data_export import mode

schema = mode.extract("invoice_template.xlsx")
```

| Parameter | Type                    | Required | Description                                 |
| --------- | ----------------------- | -------- | ------------------------------------------- |
| `path`    | `str` (path to `.xlsx`) | Required | File path of the Excel template to extract. |

Example:

```python
from mindoff_data_export import mode

# Extract and persist a schema version for later rendering/building.
schema = mode.extract("templates/monthly_sales_report.xlsx")
```

### 2) `mode.build`

Production export operation that validates runtime data, resolves placeholders, and writes final workbook output. Supports high-fidelity export and streaming export modes.

Provides: Exported workbook output on disk (`None` in fidelity mode, `list[str]` output paths in streaming mode).

Use when: You already have `schema` and `data`, and want to generate final `.xlsx` output files.

```python
from mindoff_data_export import mode

mode.build(schema, data, "filled.xlsx")
```

| Parameter               | Type                                         | Required                         | Description                                                                                                                     |
| ----------------------- | -------------------------------------------- | -------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `schema`                | `WorkbookSchema` (`dict`)                    | Required                         | Template schema containing placeholders.                                                                                        |
| `data`                  | `dict[str, Any]`                             | Required                         | Sheet-scoped runtime values (`sheet_name -> values`) or dynamic sheet group values (`sheet_key -> {output_sheet -> values}`). |
| `output_path`           | `str`                                        | Required                         | Output file path; in streaming mode, used as base name for chunked files/zip bundle naming.                                    |
| `column_width_mode`     | `str \| None` (`"fixed"`, `"even"`, `"hug"`) | Optional                         | Global override for column sizing.                                                                                              |
| `row_height_mode`       | `str \| None` (`"fixed"`, `"even"`, `"hug"`) | Optional                         | Global override for row sizing.                                                                                                 |
| `default_column_width`  | `float \| None`                              | Optional                         | Default even width override.                                                                                                    |
| `default_row_height`    | `float \| None`                              | Optional                         | Default even height override.                                                                                                   |
| `export_mode`           | `Literal["fidelity", "streaming"]`           | Optional (default: `"fidelity"`) | `"fidelity"` writes a single workbook preserving merges/styles. `"streaming"` is optimized for large dataframe-content exports. |
| `streaming_chunk_rows`  | `int`                                        | Optional (default: `50000`)      | Rows per batch in streaming mode.                                                                                               |
| `max_rows_per_workbook` | `int`                                        | Optional (default: `1048576`)    | Workbook row cap in streaming mode before splitting into `*.partNNN.xlsx` parts.                                                |

Return behavior:

- `export_mode="fidelity"` returns `None`
- `export_mode="streaming"` returns `list[str]`:
  - single-part export: `[...part001.xlsx]`
  - multi-part export: `[...zip]` (contains `*.partNNN.xlsx` files)

Input strategy:

- If input keys are already known, pass `data` directly to `mode.build(...)`.
- If input keys are template-driven or unknown, call `mode.get_inputs(...)` first.

Example:

```python
from mindoff_data_export import mode

schema = mode.extract("templates/customer_statement.xlsx")

result = mode.build(
    schema=schema,
    data={
        "Sheet1": {
            "customer_name": "Acme Industries",
            "statement_date": "2026-04-22",
            "balance_due": 12450.75,
        },
    },
    output_path="exports/acme_statement_apr_2026.xlsx",
    export_mode="fidelity",
    column_width_mode="fixed",
    row_height_mode="fixed",
)

# Fidelity mode returns None.
assert result is None
```

### 3) `mode.get_inputs`

Derives the required input contract from schema placeholders and returns a sheet-scoped contract for pre-validation, request schema generation, and runtime guardrails.

Provides: A sheet-scoped dictionary of required placeholder keys and expected types.

Use when: You need to discover dynamic input keys from the template before building your `data` payload.

```python
from mindoff_data_export import mode

inputs = mode.get_inputs(schema)
```

| Parameter | Type                      | Required | Description                                                 |
| --------- | ------------------------- | -------- | ----------------------------------------------------------- |
| `schema`  | `WorkbookSchema` (`dict`) | Required | Template schema to inspect for `{{key:type}}` placeholders. |

Example:

```python
from mindoff_data_export import mode

schema = mode.extract("templates/invoice_template.xlsx")
required_inputs = mode.get_inputs(schema)

print(required_inputs)
# Example:
# {
#   "Sheet1": {
#     "invoice_number": "number",
#     "invoice_date": "date",
#     "customer_name": "string"
#   },
#   "sheet_2": {
#     "*": {
#       "customer_name": "string"
#     }
#   }
# }
```

### 4) `mode.alter_schema`

Validates runtime data and returns a new schema with placeholders resolved. Use this for audit, transformation, or approval steps before writing output files.

Provides: A rendered schema object with placeholders replaced by runtime values.

Use when: You want to inspect or modify rendered results before calling `mode.build(...)`.

```python
from mindoff_data_export import mode

resolved_schema = mode.alter_schema(schema, data)
```

| Parameter | Type                      | Required | Description                                  |
| --------- | ------------------------- | -------- | -------------------------------------------- |
| `schema`  | `WorkbookSchema` (`dict`) | Required | Template schema containing placeholders.     |
| `data`    | `dict[str, Any]`          | Required | Sheet-scoped runtime values used to replace placeholders. |

Example:

```python
from mindoff_data_export import mode

schema = mode.extract("templates/quote_template.xlsx")
resolved_schema = mode.alter_schema(
    schema,
    {
        "Sheet1": {
            "quote_id": 22019,
            "client_name": "Contoso Retail",
            "valid_until": "2026-05-15",
        },
    },
)

# Optional inspection/transformation point before file output.
mode.build(resolved_schema, {"Sheet1": {}}, "exports/quote_22019.xlsx")
```

### Operational Notes

- Use `mode.get_inputs` before export to validate payload completeness and types early.
- Choose `export_mode="fidelity"` for styling/merge-preserving outputs.
- Choose `export_mode="streaming"` for large `dataframe-content` workloads and chunked output generation.
- In streaming mode, output may split into multiple workbook parts.
- When multiple parts are produced, they are bundled into `<output_stem>.zip` and the return value contains that zip path.

## Demo

```bash
python examples/demo.py
```

## Current Scope

- Input template: `.xlsx`
- Output export: `.xlsx`
- Public API: `mode.extract`, `mode.build`, `mode.get_inputs`, `mode.alter_schema`
