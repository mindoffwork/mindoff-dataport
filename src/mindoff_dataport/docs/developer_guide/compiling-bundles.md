# Compiling Report Bundles

Compilation is the middle step that turns a template plus your data into a `ReportBundle`, a self-contained portable artifact that the exporters know how to render. It's where validation happens, where scalar cells get resolved, and where large tables are parked in Parquet instead of being held in memory. Get a clean bundle and the export is essentially mechanical.

## Prerequisites

- Template extracted with `mo_dataport.extract()` (see [Installation](installation.md))
- Payload validated against `mo_dataport.inputs(schema)` (see [The Data Contract](data-contract.md))

## Implementation

### 1. Call `compile()`

In one call, `compile()`:

1. **Validates** your payload against the sheet contract; wrong types and missing keys fail here, before any file is written.
2. **Resolves scalars** directly into the cells they belong to, inheriting the template's styling.
3. **Materialises dataframes** to Parquet under the bundle, storing only compact *anchors* (column names, start position, style) in the plan, not the expanded rows.
4. **Plans repeats and dynamic sheets** so the exporter can stamp them out later.
5. **Shifts surrounding content** out of the path of expanding tables, if you asked it to (see [Dataframe Layout & Shifting](dataframe-layout.md)).

The signature:

```python
bundle = mo_dataport.compile(
    template=schema,
    data=payload,
    bundle_path=None,            # write to disk if set; in-memory if None
    dataframe_options=None,      # per-column occupation/alignment
    dataframe_shift="both",      # how content moves around tables
)
```

Full parameter details live in the [API Reference](api-reference.md#3-bundle-compilation).

### 2. In-Memory vs. Persisted Bundles

By default the bundle lives in memory and you export it right away. Pass a `bundle_path` and it's written to disk as a directory instead; you can re-export it any number of times, from any process, without recompiling.

```python
# Compile now, persist to disk
bundle = mo_dataport.compile(schema, data, bundle_path="saved_bundle")

# Later, even in a different script or service:
mo_dataport.export("saved_bundle", "report.xlsx")
mo_dataport.export("saved_bundle", "report.pdf", format="pdf")
```

This split is genuinely useful: a nightly job can compile once, and a web request can export on demand without paying the compile cost again. The on-disk layout is documented in [Architecture: The Pipeline](../architecture/pipeline.md#the-reportbundle-directory).

### 3. Load a Persisted Bundle

`export()` accepts a path directly, but you can also load a bundle explicitly when you need the object:

```python
from mindoff_dataport import ReportBundle

bundle = ReportBundle.load("saved_bundle")
```

### 4. Clean Up Automatically

If a persisted bundle is a throwaway (compiled only to export once), set `auto_delete_bundle=True` on the export call. The directory is removed **after** a successful export, and left in place if the export fails so you can retry.

```python
mo_dataport.export(bundle, "report.xlsx", auto_delete_bundle=True)
```

<div class="admonition warning">
<p class="admonition-title">Templates are never mutated</p>
<p>Compilation builds a new bundle and leaves your <code>WorkbookSchema</code> untouched. You can compile the same template against many different payloads without re-extracting.</p>
</div>

### 5. Cache the Extracted Schema in Production

`extract()` reads the `.xlsx` file with openpyxl every time it's called. In production, where the same template drives hundreds or thousands of reports, paying that I/O cost on every run is wasteful. Extract once, serialize the schema to disk, and load from JSON on every subsequent compile.

`WorkbookSchema` is a plain Python dict, so standard `json` works:

```python
import json
from mindoff_dataport import mo_dataport

# Once: at deploy time or when the template file changes
schema = mo_dataport.extract("invoice_template.xlsx")
with open("invoice_schema.json", "w", encoding="utf-8") as f:
    json.dump(schema, f)
```

```python
# Every report run: no .xlsx read
import json
from mindoff_dataport import mo_dataport

with open("invoice_schema.json", encoding="utf-8") as f:
    schema = json.load(f)

bundle = mo_dataport.compile(schema, data)
```

Re-extract and re-save whenever the template file itself changes. The schema is a snapshot of the template at extraction time; a stale schema will miss any placeholder or style edits made to the `.xlsx` after the last extraction.

## Core Concepts

### 1. The Bundle Stores a Plan, Not a Rendering

`report.json` holds resolved scalar cells and compact dataframe anchors (column names, start position, style), but never the expanded row data. Rows stay in Parquet under `data/`. A bundle containing a million-row table is almost as small as one containing a thousand-row table because the bundle never holds the rows, only the instructions for where and how to write them. This is what keeps persisted bundles small and what lets streaming export read in batches and hold near-constant memory.

### 2. Why Compile and Export Are Separate Steps

Compilation is the expensive part: reading the template contract, validating types, materialising dataframes to Parquet, and resolving the layout plan. Once a bundle exists, export is largely mechanical. Keeping them separate enables the pattern where a scheduled job compiles once when data is ready and a web request or script exports on demand without recompiling. The bundle is the handoff artifact between those two concerns.

## Troubleshooting

- **`KeyError` on a required placeholder.**
  The payload is missing a key the template needs. Compare against `inputs(schema)`.
- **`ValueError` about merges overlapping dataframe output.**
  A merged region sits where a table wants to expand. Either move it in the template or let the library move it with `dataframe_shift`; see [Dataframe Layout & Shifting](dataframe-layout.md#2-dataframe-collision-shifting).
- **The bundle directory wasn't deleted.**
  `auto_delete_bundle` only deletes after a *successful* export. A failed export keeps the bundle so you don't lose work.
