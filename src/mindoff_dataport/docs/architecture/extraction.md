# Template Extraction

Extraction is the front door of the pipeline. It reads an `.xlsx` file and produces a `WorkbookSchema`, a faithful in-memory blueprint of the template's content, style, geometry, and the placeholders that tell the rest of the system where data belongs. Everything downstream trusts this schema, so extraction's job is to capture the template *exactly* and lose nothing.

The owning module is `extractor.py`, supported by `schema.py` (the data structures) and `page_breaks.py` (manual print-break resolution).

## What Gets Captured

For every sheet, extraction reads:

| Captured                | Detail |
|-------------------------|--------|
| **Cell values & types** | Raw values plus an inferred type (`string`, `number`, `date`, `boolean`) |
| **Placeholders**        | Every `{{key:type}}` marker, including dataframe and repeat markers |
| **Font**                | name, size, bold, italic, underline, color (hex or symbolic theme) |
| **Fill**                | `pattern_type`, `fg_color`, `bg_color` (solid fills only) |
| **Alignment**           | horizontal, vertical, wrap_text |
| **Borders**             | four sides with per-side style and color, plus merged-region edge handling |
| **Dimensions**          | per-column widths and per-row heights |
| **Merged regions**      | the authoritative set of merged ranges for the sheet |
| **Page breaks**         | manual `row_page_breaks` / `column_page_breaks` from Excel metadata |
| **Theme colors**        | a compact `theme_colors` map so non-OpenPyXL renderers can resolve `theme:index:tint` |
| **Gridlines**           | the sheet's `show_gridlines` flag |

## Key Behaviors

### 1. Formulas Are Preserved

The workbook is opened with `data_only=False`, so formula cells keep their formulas rather than collapsing to cached values. A template that computes a total keeps computing it in the output.

### 2. Column Spans Are Expanded

Excel stores column dimensions as ranges (for example, `C:CV` sharing one width). Extraction expands these into per-column widths so that ranged widths survive faithfully into export, where each column is addressed individually.

### 3. Theme Colors Stay Symbolic

A cell colored from the workbook theme is stored as a symbolic `theme:<index>:<tint>` value, **not** flattened to RGB. This preserves perfect fidelity for the OpenPyXL renderer (which understands themes natively) while a compact `theme_colors` map travels alongside so renderers that cannot (XlsxWriter, ReportLab) can resolve the same symbol to concrete RGB.

### 4. Merged-Region Borders

A border drawn around a merged block in Excel belongs to the whole region, not just its anchor cell. Extraction records the merged region's outer edges so the renderers can draw borders around the **full** region later. It is easy to get wrong, and the visual difference is immediate.

### 5. Manual Page Breaks

`page_breaks.py` reads Excel's manual print-break metadata into `row_page_breaks` and `column_page_breaks` on the sheet schema. At this stage they're recorded against template row/column indexes; they're re-resolved against the final layout during compilation (see [Compilation & Bundle](compilation.md)).

## The WorkbookSchema

`schema.py` defines the structures that hold all of the above: the workbook, its sheets, and per-cell style/value records. Two properties matter for the rest of the pipeline:

- It is **complete**: a renderer can reproduce the template from the schema alone.
- It is **inert**: extraction never writes back to the source file, and compilation never mutates the schema.

Because `WorkbookSchema` is a plain Python dict, it serializes directly with `json.dump` / `json.load`. In production, extract once when the template changes, persist the result, and load from JSON on every subsequent compile. This skips the openpyxl `.xlsx` read entirely. See [Compiling Report Bundles: Cache the Extracted Schema in Production](../developer_guide/compiling-bundles.md#5-cache-the-extracted-schema-in-production) for the full pattern.

## Troubleshooting Extraction

1. **A placeholder wasn't discovered.** Verify the exact `{{key:type}}` syntax and a supported type; malformed markers are treated as plain text.
2. **A color looks wrong downstream.** Confirm whether it's a theme color; symbolic theme values are resolved per-renderer, and non-OpenPyXL renderers rely on the captured `theme_colors` map.
3. **A ranged column width didn't apply.** Ranged widths are expanded per column at extraction; if a width is missing, check that the template actually set it on that column.

