# Dataframe Layout & Shifting

Tables are the part of a report that doesn't know its own size until runtime. A template has a fixed grid; a dataframe might be 3 rows today and 30,000 tomorrow. This page covers the two levers that reconcile that tension: **column layout** (how a table's columns map onto template columns) and **collision shifting** (how surrounding content moves out of the way as a table grows).

## Prerequisites

- Template extracted with `mo_dataport.extract()` and placeholders understood (see [Templates & Placeholders](templates-and-placeholders.md))
- Polars installed: `pip install "mindoff-dataport[polars]"`

## Implementation

### 1. Column Layout: Occupation & Alignment

By default each dataframe column lands in one template column and keeps the anchor cell's alignment. When your design needs a column to span more space, or numbers to align right while labels align left, use `dataframe_options` during `compile()`.

The structure is keyed by **resolved sheet name → placeholder key → columns**:

```python
dataframe_options = {
    "Sheet Name": {
        "placeholder_key": {
            "columns": {
                "Column Name": {"occupation": 2, "alignment": "left"},
            }
        }
    }
}
```

`occupation` is how many template columns the dataframe column should span; `alignment` overrides the horizontal alignment for that column's generated cells.

When a template splits headers and content across separate placeholders, configure each independently. A common pattern is centered headers over right-aligned numbers:

```python
dataframe_options = {
    "Column Layout": {
        "headers": {
            "columns": {
                "Employee Name": {"occupation": 2, "alignment": "center"},
                "Department":    {"occupation": 2, "alignment": "center"},
                "Amount":        {"occupation": 1, "alignment": "center"},
            }
        },
        "rows": {
            "columns": {
                "Employee Name": {"occupation": 2, "alignment": "left"},
                "Department":    {"occupation": 2, "alignment": "center"},
                "Amount":        {"occupation": 1, "alignment": "right"},
            }
        },
    }
}
```

Rules:

- `occupation` must be a positive integer.
- `alignment` must be `"left"`, `"center"`, or `"right"`.
- Options are keyed by *resolved output sheet name*, then placeholder key.
- Any column you don't configure defaults to `occupation=1` and keeps the template cell's alignment.

### 2. Dataframe Collision Shifting

Here's the problem shifting solves: you put a label two columns to the right of a table anchor, the table grows to four columns wide, and now the table is sitting on top of your label. `dataframe_shift` tells `compile()` to move the colliding template cells and merged regions out of the way before export.

```python
bundle = mo_dataport.compile(
    schema,
    data,
    dataframe_shift="both",  # "both", "horizontal", "vertical", or "none"
)
```

| Mode           | Behavior                                                                        |
|----------------|---------------------------------------------------------------------------------|
| `"both"`       | Shift right-side cells/merges horizontally and lower cells/merges vertically    |
| `"horizontal"` | Shift only what sits to the right of dataframe output                           |
| `"vertical"`   | Shift only what sits below dataframe output                                     |
| `"none"`       | Don't shift; a template merge overlapping dataframe output raises `ValueError`  |

The shift is **metadata-only**. Dataframe rows stay in Parquet, `report.json` keeps compact anchors, and streaming export still reads in batches; nothing is materialised just to move it. The same shifted layout drives both XLSX and PDF, including later dataframe anchors inside repeat blocks.

<div class="admonition warning">
<p class="admonition-title">Some merges can't be shifted</p>
<p>A merged region that genuinely cannot be moved clear of dataframe output still raises <code>ValueError</code>, including inside repeat sections. When that happens, redesign that part of the template so the merge doesn't straddle the table's growth path.</p>
</div>

See `examples/dataframe_shift/xlsx.py` and `examples/dataframe_shift/pdf.py` for runnable demonstrations.

### 3. Manual Page Breaks

Excel's manual print breaks travel with the template into the schema. During `compile()`, they are re-resolved against the final layout (after tables expand and content shifts), so a break you set on "row 20" still lands where you meant it, even if a table pushed that row down to row 4,000.

- `row_page_breaks` start a new printed page after the given 1-based template row.
- `column_page_breaks` start a new printed page after the given 1-based template column (XLSX only).
- **PDF** uses resolved row breaks as hard page boundaries and ignores column breaks.
- **XLSX** preserves both resolved row and column breaks, in fidelity and streaming.

See `examples/page_break/xlsx.py` and `examples/page_break/pdf.py`.

## Core Concepts

### 1. The Template Grid Is Fixed; the Data Is Not

A template has a static layout: cells at known positions, widths set per column. A dataframe's column count and content aren't known until runtime. Column layout (occupation and alignment) is how you describe the intended relationship between your data columns and the template grid at design time. Collision shifting is how the library enforces that relationship at compile time by moving whatever stands in the way of a table that grows wider or taller than the template originally assumed.

### 2. Shifting Is Metadata-Only

No rows are materialised to perform a shift. The plan in `report.json` is updated to reflect new cell positions, but the dataframe rows stay in Parquet and streaming still reads them in batches afterward. Shifting costs almost nothing at compile time and nothing at export time. The same shifted layout then drives both XLSX and PDF without a second shift pass.

## Troubleshooting

- **`dataframe_options` had no effect.**
  Check the keys: they must be the *resolved output sheet name* then the placeholder key. For dynamic sheets, that's the generated sheet name, not `{{key}}`.
- **`ValueError` on a merge during compile.**
  A merged region overlaps where a table expands. Switch `dataframe_shift` away from `"none"`, or move the merge in the template.
- **A page break landed in the wrong place in PDF.**
  PDF ignores column breaks by design. Only row breaks act as PDF page boundaries.
