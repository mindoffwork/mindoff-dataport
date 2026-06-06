# The Pipeline

`mindoff-dataport` is built around a single, linear pipeline. Data flows one way: from a template file, through an in-memory schema, into a portable bundle, and out to a rendered file. Each stage has one job and one output.

For a user-facing walkthrough of these steps, see the [Quick Start](../index.md#quick-start).

## Data Flow

```text
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

## The Three Stages

| Stage         | Entry point                          | Input → Output                         | Owns |
|---------------|--------------------------------------|----------------------------------------|------|
| **Extract**   | `extractor.py` / `extract_template`  | `.xlsx` → `WorkbookSchema`             | Reading styles, dimensions, merges, breaks, placeholders |
| **Compile**   | `bundle.py` / `compile_report_bundle`| schema + data → `ReportBundle`         | Validation, scalar resolution, anchors, repeat plans, shifting |
| **Export**    | `xlsx_renderer.py`, `pdf_renderer.py`| bundle → `.xlsx` / `.pdf`              | Materialising the plan into a styled file |

Each stage is covered in depth on its own page: [Extraction](extraction.md), [Compilation & Bundle](compilation.md), [XLSX Rendering](xlsx-rendering.md), and [PDF Rendering](pdf-rendering.md).

## The Core Design Idea

The pipeline's defining decision is that **the bundle stores a plan, not a rendering**. Scalars are resolved into cells, but tables are *not* expanded; `report.json` holds compact anchors (column names, start position, style), and the actual rows stay in Parquet under `data/`. Expansion happens only at export time, row by row.

That single choice is what makes large exports possible. A million-row table costs the same in the bundle as a thousand-row one, because the bundle never contains the rows, only the instructions for where and how to write them. The streaming renderers then read Parquet in batches and emit output incrementally, so peak memory stays flat regardless of dataset size. The proof is in [Benchmarking](benchmarking.md).

## The ReportBundle Directory

When you pass `bundle_path` to `compile()`, the bundle is written as a directory. The same directory can be reloaded and re-exported without recompiling.

```text
report_bundle/
├── manifest.json      # bundle version, inputs, sheet metadata, dataframe sources, capabilities
├── report.json        # resolved scalar cells and dataframe anchor/repeat plans
└── data/
    └── *.parquet      # dataframe sources materialised from Polars inputs
```

| File / dir      | Holds                                                                              |
|-----------------|------------------------------------------------------------------------------------|
| `manifest.json` | Bundle version, the input contract, per-sheet metadata, dataframe sources, capabilities |
| `report.json`   | Resolved scalar/static cells, and dataframe **anchors** plus repeat plans (not expanded rows) |
| `data/*.parquet`| Dataframe sources materialised from the Polars inputs; read in batches at export time |

<div class="admonition note">
<p class="admonition-title">Anchors, not rows</p>
<p><code>report.json</code> stores dataframe anchors (column names, start row/column, and style) and never the expanded row data. Rows live in Parquet and are read at export time. This is why a persisted bundle stays small even for enormous datasets.</p>
</div>

## Invariants the Pipeline Guarantees

A few properties hold across the whole pipeline, by design:

- **Templates are never mutated.** Compilation builds a new bundle; the input schema is untouched.
- **Formulas are preserved.** Extraction reads with `data_only=False`, so formula cells survive.
- **XLSX and PDF consume the same resolved layout.** The two renderers share one plan; differences are limited to supplied font availability and deterministic page scaling in PDF.
- **Output sheet names stay unique**, and sheet order follows template order.

These are the contracts the rest of the architecture relies on. The pages that follow show how each stage upholds them.
