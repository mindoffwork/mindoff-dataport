# mindoff_data_export

Extract an `.xlsx` template into JSON, detect typed placeholders, and regenerate a filled `.xlsx` with the same layout and formatting.

## Install

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -e .
```

## Actual Demo (Direct Extract -> Typed Data -> Filled Excel)

The default demo now runs fully in memory (no intermediate JSON files):

1. Start from `.xlsx` template  
   `examples/input/customer_statement_template.xlsx`
2. Extract schema directly from template
3. Discover required variables + types (`get_template_inputs`)
4. Build typed dictionary values (scalar + Polars DataFrame / LazyFrame)
5. Regenerate one or more filled `.xlsx` files

Run:

```bash
python examples/demo.py
# or
python examples/real_usage_case.py
```

Generated outputs:

- `examples/output/real_usage_case/customer_statement_fixed.xlsx`
- `examples/output/real_usage_case/customer_statement_hug_columns.xlsx`
- `examples/output/real_usage_case/customer_statement_lazyframe.xlsx`

## Placeholder Format

Use placeholders in template cells as:

```text
{{variable_name:type}}
```

Supported scalar types:

- `string`
- `number`
- `int`
- `float`
- `boolean`
- `date`

Supported table types:

- `dataframe-headers`
- `dataframe-data`

## Minimal Library Example

```python
import datetime as dt
import polars as pl
from mindoff_data_export import extract_template, get_template_inputs, build_template_with_data

schema = extract_template("template.xlsx")
variables = get_template_inputs(schema)   # {"customer_name": "string", ...}

data = {
    "customer_name": "Acme",
    "invoice_number": 1024,
    "invoice_amount": 24999.75,
    "is_paid": True,
    "invoice_date": dt.date(2026, 4, 22),
    "line_items": pl.DataFrame({
        "Item": ["A", "B"],
        "Qty": [2, 3],
    }),
}

build_template_with_data(
    schema,
    data,
    "filled.xlsx",
    column_width_mode="hug",
    row_height_mode="even",
    default_row_height=20.0,
)
```

## Demo Customizations

The demo covers these customization points:

- Placeholder typing and validation:
  - scalar: `string`, `number`, `int`, `float`, `boolean`, `date`
  - table: `dataframe-headers`, `dataframe-data`
- Polars integration:
  - `pl.DataFrame` as placeholder value
  - `pl.LazyFrame` support (auto-collected during rendering)
- Output sizing:
  - `column_width_mode`: `fixed`, `hug`, `even`
  - `row_height_mode`: `fixed`, `hug`, `even`
  - `default_column_width`, `default_row_height` overrides
- Multi-output generation:
  - generate different workbook variants from same extracted schema + data

## What Is Preserved

When rebuilding, the library preserves:

- Sheet names and dimensions
- Cell values and formulas
- Fonts, fills, alignment, borders
- Merged ranges
- Column widths and row heights

## Project Files

- `src/mindoff_data_export/extractor.py`: `.xlsx` -> schema dict
- `src/mindoff_data_export/renderer.py`: placeholder discovery + rendering
- `src/mindoff_data_export/builder.py`: schema dict -> `.xlsx`
- `examples/demo.py`: main demo entrypoint
- `examples/real_usage_case.py`: full end-to-end implementation used by demo
- `examples/input/customer_statement_template.xlsx`: sample template
