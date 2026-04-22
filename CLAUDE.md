# mindoff_data_export

Python package: extract `.xlsx` templates → JSON schema; rebuild `.xlsx` from JSON with exact visual fidelity. Supports template variable placeholders and automatic column/row sizing.

## Layout

```
src/mindoff_data_export/
  schema.py      TypedDicts (WorkbookSchema, SheetSchema, CellSchema, …)
  utils.py       normalize_color, argb_to_color, border_side_to_dict, dict_to_border_side
  extractor.py   extract_template(path) → WorkbookSchema
  builder.py     build_template(schema, output_path, **sizing_kwargs)
  renderer.py    get_template_inputs(schema), render_schema(schema, data)
  __init__.py    re-exports all public API
tests/
  conftest.py         session fixtures: fixture_path, workbook_schema; auto-creates sample_template.xlsx
  test_extractor.py
  test_builder.py
  test_roundtrip.py
  test_sizing.py      column_width_mode / row_height_mode tests
  test_renderer.py    {{key:type}} placeholder and dataframe expansion tests
examples/
  demo.py             end-to-end walkthrough of all four public workflows
  input/              place .xlsx files here
  config/             extracted JSON schemas and template JSONs
  output/             built .xlsx files land here
```

## Public API

```python
from mindoff_data_export import (
    extract_template,
    build_template,
    build_template_with_data,
    get_template_inputs,
    render_schema,
)

# 1. Extract
schema = extract_template("file.xlsx")   # data_only=False — preserves formulas

# 2. Rebuild (exact visual fidelity)
build_template(schema, "out.xlsx")

# 3. Rebuild with automatic sizing (kwargs override per-sheet values)
build_template(schema, "out.xlsx",
    column_width_mode="hug",    # "fixed" | "even" | "hug"
    row_height_mode="even",     # "fixed" | "even" | "hug"
    default_row_height=20.0,    # used by "even" mode
    default_column_width=15.0,  # used by "even" mode
)

# 4. Template variables — fill {{key:type}} placeholders at runtime
inputs = get_template_inputs(schema)   # → {"name": "string", "rows": "dataframe-headers"}
build_template_with_data(schema, data, "out.xlsx",
    column_width_mode="hug",
    row_height_mode="even",
    default_row_height=20.0,
)
```

## JSON Schema shape

```
WorkbookSchema { sheets: SheetSchema[] }
SheetSchema {
  name, dimensions, merged_regions: str[],
  column_widths: {col: float}, row_heights: {str(int): float},
  cells: { coord: CellSchema }

  # Optional sizing — also settable via build_template kwargs (kwargs take precedence)
  column_width_mode?: "fixed" | "even" | "hug"   # default "fixed"
  default_column_width?: float                    # used by "even"
  row_height_mode?: "fixed" | "even" | "hug"     # default "fixed"
  default_row_height?: float                      # used by "even"
}
CellSchema {
  coordinate, value, cell_type: string|number|date|formula|empty,
  number_format, font, fill: {bg_color: ARGB|null},
  alignment: {horizontal, vertical, wrap_text},
  borders: {top,bottom,left,right: {style, color: ARGB|null}},
  merged: bool, merge_anchor: coord|null
}
```

Colors are ARGB hex (`"FF003366"`) or `"theme:<idx>:<tint>"` or `null`.

## Template variable placeholders

Cell values in a schema JSON may contain `{{key:type}}` tokens:

| Placeholder | Type | Behavior |
|---|---|---|
| `{{name:string}}` | scalar | inline or whole-cell string substitution |
| `{{count:number}}` | scalar | whole-cell numeric value |
| `{{date:date}}` | scalar | whole-cell date value |
| `{{rows:dataframe-headers}}` | DataFrame | writes column names (bold) then data rows |
| `{{rows:dataframe-data}}` | DataFrame | writes data rows only (developer controls headers) |

- `get_template_inputs(schema)` scans all cells and returns `{key: type}` for every placeholder.
- `render_schema(schema, data)` validates types (raises `KeyError`/`TypeError`) and returns a resolved schema.
- `build_template_with_data(schema, data, path, **sizing_kwargs)` renders then builds in one call.
- polars `DataFrame` and `LazyFrame` and pandas `DataFrame` are all accepted for dataframe types.

## Sizing modes

Controlled via kwargs on `build_template` / `build_template_with_data` (preferred) or as optional fields in `SheetSchema`. kwargs always win.

| Mode | Column behavior | Row behavior |
|---|---|---|
| `"fixed"` | uses `column_widths` dict | uses `row_heights` dict |
| `"even"` | all columns → `default_column_width` (default 15.0) | all rows → `default_row_height` (default 15.0) |
| `"hug"` | each column fits widest cell (char count × bold factor + 2) | each row height = max font size × 1.5 |

`"hug"` runs after cells are written; `"fixed"` and `"even"` run before.

## Key invariants

- `merged_regions` on SheetSchema is authoritative; builder calls `ws.merge_cells()` **after** writing all cells.
- Non-anchor merged cells (MergedCell stubs) have no style data; extractor returns empty font/fill/borders for them.
- Fill: always use `fgColor` for solid fills (`bgColor` is two-color pattern only).
- Row dimension keys: stored as `str` in JSON, converted to `int` when writing to openpyxl.
- openpyxl style objects are immutable — always construct new Font/PatternFill/Alignment/Border; never mutate.
- Load with `data_only=False` to preserve formula strings.
- `render_schema` never mutates the input schema — always returns new dicts.

## Run tests

```bash
PYTHONPATH=src python -m pytest tests/ -v
```

Gold-standard: `test_roundtrip.py::test_roundtrip_schema_identical` — extract → build → re-extract → deep JSON compare.

## Run examples

```bash
python examples/demo.py
```

Outputs land in `examples/output/`. Template placeholders are defined in `examples/config/report_template.json`.

## Adding features

When extending the schema (new cell properties, sheet-level properties, etc.):
1. Add/update TypedDicts in `schema.py`
2. Add extraction logic in `extractor.py` (mind MergedCell stub edge case)
3. Add reconstruction logic in `builder.py`
4. Add rendering logic in `renderer.py` if new placeholder types are involved
5. Update `tests/conftest.py` fixture if new properties need coverage
6. Add test assertions; ensure roundtrip test still passes
7. Update `examples/demo.py` to reflect the change
8. Update this file

## Dependencies

- `openpyxl>=3.1.0` — only xlsx library used
- `polars` or `pandas` — optional; required only when using dataframe placeholder types
- `pytest>=8.0`, `pytest-cov>=5.0` (dev)
- Build backend: `hatchling`
- Python `>=3.10`
