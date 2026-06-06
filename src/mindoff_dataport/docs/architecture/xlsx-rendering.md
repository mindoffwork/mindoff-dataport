# XLSX Rendering

The XLSX renderer turns a bundle's plan into a styled Excel file. It offers two strategies behind one entry point, `export_report_bundle`: a full-fidelity in-memory render and a low-memory streaming render. The owning module is `xlsx_renderer.py`, with style and sizing helpers split into `xlsx_builder.py` and color/border conversion into `style_conversion.py`.

For the user-facing options, see [Developer Guide → Exporting to Excel](../developer_guide/exporting-xlsx.md).

## Module Layout

| Module                | Responsibility |
|-----------------------|----------------|
| **`xlsx_renderer.py`** | The export entry point, mode selection, sheet iteration, dataframe expansion, streaming and split logic |
| **`xlsx_builder.py`**  | Style and sizing helper functions used while building cells |
| **`style_conversion.py`** | OpenPyXL color and border conversion helpers (including theme resolution) |

## Two Modes, One Plan

Both modes read the *same* compiled plan; they differ only in how they write it.

### 1. Fidelity Mode

Fidelity renders the whole workbook in memory. It supports every feature: all merge behavior, all sizing modes (including `hug`), full style fidelity. It's the default because most reports are modest in size and want everything. The trade-off is memory: peak usage scales with the dataset, so it's intended for outputs up to roughly tens of thousands of rows.

### 2. Streaming Mode

Streaming reads `dataframe-content` from Parquet in batches (`streaming_chunk_rows`, default 50K) and writes rows incrementally, so peak memory stays near-constant as row count grows. That flat profile is the reason streaming exists; see [Benchmarking](benchmarking.md). The constraints follow directly from "never buffer all rows":

- No `hug` sizing (it would require measuring every row).
- No merged cells intersecting `dataframe-content` rows after compile-time shifting.
- One `dataframe-content` placeholder per non-repeat sheet.

Streaming repeat sheets consume dataframe-content-covered row offsets exactly once, and must not emit extra blank rows after repeated output. The renderer enforces this.

## Streaming Engines

Streaming can write through one of two backends:

| Engine        | Status                | Notes |
|---------------|-----------------------|-------|
| `openpyxl`    | Default               | The safe, fully-supported path; used for both fidelity and streaming |
| `xlsxwriter`  | Opt-in (streaming only) | Can be faster for pure-write workloads; rejected in fidelity mode |

XlsxWriter resolves theme colors through the bundle's `theme_colors` map before writing RGB styles, whereas OpenPyXL keeps theme colors symbolic and resolves them natively. For generated dataframe `occupation` merges, XLSX streaming uses a fast renderer-owned merge insertion path; template merges still go through normal validation.

## Materialising Merges Correctly

Merged cells are a frequent source of subtle Excel bugs, so the renderer is deliberate about them:

- A merged region's borders are drawn around the **full** region, not only the anchor.
- Materialisation preserves visible outline and diagonal borders **without** emitting interior `horizontal`/`vertical` merge fields that can make Excel suppress left/right edges.
- Renderer-generated dataframe `occupation` merges apply the anchor's border styling on **every** generated row, not just the first.

## Sizing & Splitting

Sizing modes (`fixed`, `even`, `hug`) resolve per the rules in [Sizing & Styling](../developer_guide/sizing-and-styling.md); for `dataframe-content` in fixed row mode, generated rows inherit the anchor row height when no explicit height exists. When a sheet crosses `max_rows_per_workbook`, streaming writes workbook parts, zips them into `output.zip`, removes the loose parts, and returns the zip path in a one-item list.

## Invariants

- Builder converts JSON row keys from `str` back to `int`.
- OpenPyXL styles are immutable, so the builder always creates new style objects rather than mutating shared ones.
- `FillSchema` canonical form requires `pattern_type`, `fg_color`, and `bg_color`; the older `bg_color`-only shape is handled defensively.
- Sheet gridline visibility (`show_gridlines`) is preserved.
