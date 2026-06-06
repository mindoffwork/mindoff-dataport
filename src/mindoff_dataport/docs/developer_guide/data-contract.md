# The Data Contract

Payloads in `mindoff-dataport` are **sheet-scoped**: the top-level keys of your data dict are sheet names, and each maps to that sheet's values. This one rule keeps multi-sheet reports unambiguous: every value knows exactly which sheet it belongs to. If you ever forget the exact shape a template wants, `mo_dataport.inputs(schema)` will tell you (see [Input Discovery](api-reference.md)).

## Prerequisites

- Template extracted with `mo_dataport.extract()` (see [Installation](installation.md))
- Placeholder keys confirmed with `mo_dataport.inputs(schema)` before building a payload

## Implementation

### 1. Static Sheet

The common case: a sheet with some scalar values and a table. The outer key matches the sheet name in the template; the inner keys match your placeholder keys.

```python
{
    "Invoice": {
        "customer_name": "Acme Industries",    # string
        "invoice_number": 1024,                 # number
        "due_date": datetime.date(2026, 5, 1),  # date
        "line_items": polars_dataframe,         # dataframe / LazyFrame
    }
}
```

### 2. Repeat Section

When a sheet contains a repeat block (`{{key:repeat-start}}` ... `{{key:repeat-end}}`), the payload key matching that block holds an **ordered list** of record payloads, one per rendered block.

```python
{
    "Sheet1": {
        "reports": [                            # key matches the repeat-start/end key
            {"customer_name": "Acme",   "line_items": acme_df},
            {"customer_name": "Globex", "line_items": globex_df},
        ]
    }
}
```

The blocks render top to bottom in list order. For the layout rules (sibling sections, static rows, merged-cell limits), see the [Repeat Sections recipe](recipes.md#2-repeat-sections-per-customer-blocks).

### 3. Dynamic Sheet Group

When a template sheet is named exactly `{{key}}`, it's a stencil for many output sheets. Pass a dict of `output_sheet_name -> payload` under that placeholder key:

```python
{
    "region_sheet": {                           # the sheet-name placeholder key
        "North Sheet": {                        # becomes an output sheet
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

Output sheet order follows the payload dict's insertion order. Output sheet names must be unique.

When you introspect a template with a dynamic group, `inputs(schema)` reports it under the same placeholder key, using `"*"` to mean "every generated sheet shares this contract":

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

### 4. Using Polars LazyFrames for Large Data

For anything big, prefer a `LazyFrame` over a `DataFrame`. A `LazyFrame` stays disk-backed and is never fully loaded into memory; the library reads it in batches at export time.

```python
import polars as pl

rows = pl.scan_parquet("sales.parquet").select(["product", "units", "revenue"])

bundle = mo_dataport.compile(schema, {"Sheet1": {"sales_rows": rows}})
```

<div class="admonition tip">
<p class="admonition-title">Rule of thumb</p>
<p>Use <code>pl.scan_parquet(...)</code> (lazy) for large or unbounded data and <code>pl.DataFrame(...)</code> (eager) for small, already-in-memory tables. Both are accepted anywhere a dataframe is expected.</p>
</div>

## Troubleshooting

- **`KeyError` during `compile()`.**
  A placeholder the template requires is missing from your payload. Run `inputs(schema)` and compare keys side by side.
- **A sheet renders empty.**
  Confirm the outer payload key exactly matches the sheet name (including spaces and case).
- **Type validation failed.**
  The value you passed doesn't match the placeholder type, for example a `str` where the template marked `:number`. The error names the offending key.
