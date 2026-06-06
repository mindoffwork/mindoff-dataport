# Exporting to Excel (XLSX)

Once you have a bundle, exporting to Excel is a single call. The interesting choice is **how**: a full-fidelity render that supports every feature, or a streaming render that holds memory flat no matter how many rows you throw at it.

```python
mo_dataport.export(bundle, "report.xlsx")                       # fidelity (default)
mo_dataport.export(bundle, "report.xlsx", export_mode="streaming")
```

## Prerequisites

- A compiled `ReportBundle` or a path to a persisted bundle directory (see [Compiling Report Bundles](compiling-bundles.md))
- Polars installed if the bundle contains dataframe sources: `pip install "mindoff-dataport[polars]"`

## Implementation

### 1. Choose an Export Mode

| Mode          | Best for                       | Trade-off                                                   |
|---------------|--------------------------------|-------------------------------------------------------------|
| `"fidelity"`  | Up to 50K rows, full styling   | Renders in memory; supports every feature, no constraints  |
| `"streaming"` | Large or unbounded datasets    | Near-constant memory; a few feature limits (listed below)  |

Fidelity is the default because most reports are small and want every feature. Reach for streaming when row counts climb into the hundreds of thousands and memory becomes the bottleneck. Under the hood, streaming reads Parquet in batches and writes rows incrementally, so peak memory barely moves as the dataset grows. The numbers are in [Architecture: Benchmarking](../architecture/benchmarking.md).

Streaming trades a little flexibility for that flat memory profile:

- No `hug` sizing (it would require buffering every row to measure content).
- No merged cells may remain intersecting `dataframe-content` rows after compile-time shifting.
- Only one `dataframe-content` placeholder per non-repeat sheet.

If your report fits within these, streaming is a free win on memory.

### 2. Configure Export Options

All options are keyword arguments to `export()`.

| Option                  | Type    | Default      | Description                                                                                   |
|-------------------------|---------|--------------|-----------------------------------------------------------------------------------------------|
| `export_mode`           | `str`   | `"fidelity"` | `"fidelity"` (full in-memory render) or `"streaming"` (row-by-row, lower memory)              |
| `streaming_engine`      | `str`   | `"openpyxl"` | Streaming writer backend: `"openpyxl"` or `"xlsxwriter"`. Rejected in fidelity mode           |
| `column_width_mode`     | `str`   | schema value | `"fixed"`, `"even"`, or `"hug"`. Overrides the template schema. See [Sizing & Styling](sizing-and-styling.md) |
| `row_height_mode`       | `str`   | schema value | `"fixed"`, `"even"`, or `"hug"`. Overrides the template schema                                |
| `default_column_width`  | `float` | schema value | Fallback width (Excel character units) for `"even"` mode or when no width is stored           |
| `default_row_height`    | `float` | schema value | Fallback height (points) for `"even"` mode or when no height is stored                        |
| `streaming_chunk_rows`  | `int`   | `50000`      | Parquet rows read per batch during streaming                                                  |
| `max_rows_per_workbook` | `int`   | `1048576`    | Split output into multiple `.xlsx` parts when a sheet reaches this row limit                  |
| `auto_delete_bundle`    | `bool`  | `False`      | Delete the bundle directory after a successful export                                         |

<div class="admonition tip">
<p class="admonition-title">Which engine?</p>
<p>OpenPyXL is the default streaming engine and the safe choice. <code>xlsxwriter</code> is opt-in and can be faster for some pure-write workloads; benchmark both against your own data if write speed is critical.</p>
</div>

### 3. Split Large Output

Excel caps a worksheet at 1,048,576 rows. When you're exporting beyond that (or just want manageable file sizes), set `max_rows_per_workbook`. In streaming mode, the library writes the workbook parts, bundles them into `output.zip`, deletes the loose parts, and returns a one-item `list[str]` containing the zip path.

```python
outputs = mo_dataport.export(
    bundle,
    "output.xlsx",
    export_mode="streaming",
    max_rows_per_workbook=500_000,
)
# outputs -> ["output.zip"] when a split occurred
```

### 4. Understand the Return Value

- Fidelity XLSX returns `None`.
- Streaming XLSX returns a `list[str]`: a single workbook path when no split is needed, or one `.zip` path when the output was split.

## Core Concepts

### 1. How Streaming Achieves Flat Memory

Streaming reads `dataframe-content` from Parquet in batches (controlled by `streaming_chunk_rows`, default 50K rows) and writes each batch to the workbook immediately. At any point, only one batch is in memory. That's why peak RSS barely moves as dataset size grows. Fidelity, by contrast, expands all rows before writing, so memory scales with row count. The numbers are in [Architecture: Benchmarking](../architecture/benchmarking.md).

### 2. When to Choose Streaming over Fidelity

Row count alone isn't the deciding factor. Reach for streaming when:

- Row counts exceed tens of thousands and memory is a constraint.
- The data source is a Polars `LazyFrame` or Parquet file you'd rather not fully load into memory.
- You need to split output across multiple workbooks with `max_rows_per_workbook`.

Stay with fidelity when the report needs `hug` sizing, retains merged cells across `dataframe-content` rows, or uses more than one `dataframe-content` placeholder per non-repeat sheet.

## Troubleshooting

- **Streaming rejected my sheet.**
  Check the three streaming constraints above; usually it's `hug` sizing or a merged cell sitting on `dataframe-content` rows.
- **`streaming_engine` raised an error.**
  It's only valid in streaming mode (and never for PDF). Remove it for fidelity exports.
- **I expected one file but got a zip.**
  A sheet crossed `max_rows_per_workbook`, so the export was split. Inspect the returned `list[str]`.
