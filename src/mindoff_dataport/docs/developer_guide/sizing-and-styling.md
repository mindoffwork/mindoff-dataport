# Sizing & Styling

Styling is something you do in Excel, not in code. You design the workbook (fonts, fills, borders, alignment) and the library extracts those styles during `extract()` and reapplies them faithfully at export time. There is no runtime style API to learn. Sizing is the one dimension you *can* steer at export time, because tables don't have a fixed height until the data arrives.

## Prerequisites

- Template designed in Excel with the desired fonts, fills, borders, and alignment already applied
- Template extracted with `mo_dataport.extract()` (styles are captured at extraction time)

## Implementation

### 1. Sizing Modes

Column widths and row heights are each computed by one of three modes. You set defaults in the template; you can override per export with keyword arguments.

**Column width modes:**

| Mode      | Source                                              | Limitation                                  |
|-----------|-----------------------------------------------------|---------------------------------------------|
| `"fixed"` | Per-column widths stored in the template schema     | Requires widths to be set in the template   |
| `"even"`  | Applies `default_column_width` uniformly            | Ignores per-column template widths          |
| `"hug"`   | Computes width from cell content at render time     | Not available in streaming; not for PDF `dataframe-content` columns |

**Row height modes:**

| Mode      | Source                                              | Limitation                                  |
|-----------|-----------------------------------------------------|---------------------------------------------|
| `"fixed"` | Per-row heights stored in the template schema       | Requires heights to be set in the template  |
| `"even"`  | Applies `default_row_height` uniformly              | Ignores per-row template heights            |
| `"hug"`   | Auto-fits row height to content                     | Not available in streaming                  |

For `dataframe-content` in fixed row mode, generated rows inherit the anchor row's height when no explicit height is stored for those rows. And a useful asymmetry for PDF: row `hug` is supported for `dataframe-content` (each streamed chunk auto-sizes), but column `hug` is not; measuring column widths would require buffering every row first.

**Units and defaults:**

| Parameter              | Unit                       | Default in schema |
|------------------------|----------------------------|-------------------|
| `default_column_width` | Excel character units      | `15.0`            |
| `default_row_height`   | Points                     | `15.0`            |
| `margin` (PDF)         | Points (1pt = 1/72 inch)   | `36`              |

Any of these passed to `export()` override the value stored in the template schema.

### 2. Supported Styling

Everything below is read from the template and reapplied at export. No configuration needed: design it once, get it everywhere.

**Font properties:**

| Property    | Values / Range                            | Notes                                                            |
|-------------|-------------------------------------------|------------------------------------------------------------------|
| `name`      | Any font family name                      | Falls back to Helvetica in PDF unless registered as a custom font |
| `size`      | `float` (points)                          | Default `11.0`                                                   |
| `bold`      | `True` / `False`                          |                                                                  |
| `italic`    | `True` / `False`                          |                                                                  |
| `underline` | `"single"`, `"double"`, `None`            | Rendered in PDF via `<u>` markup                                 |
| `color`     | Hex ARGB string or `theme:<index>:<tint>` | PDF resolves theme colors against the default Office palette     |

**Fill properties:**

| Property       | Values                                            | Notes                                  |
|----------------|---------------------------------------------------|----------------------------------------|
| `pattern_type` | `"solid"`, `None`                                 | Patterned fills are not extracted or rendered |
| `fg_color`     | Hex ARGB, `theme:<index>:<tint>`, or `None`       | The visible color for solid fills      |
| `bg_color`     | Hex ARGB, `theme:<index>:<tint>`, or `None`       | The pattern background                 |

**Alignment properties:**

| Property     | Values                                                  |
|--------------|---------------------------------------------------------|
| `horizontal` | `"left"`, `"center"`, `"right"`, `"centerContinuous"`, `"justify"`, `"distributed"` |
| `vertical`   | `"top"`, `"center"`, `"bottom"`                         |
| `wrap_text`  | `True` / `False`                                        |

In PDF, newline characters become line breaks only when `wrap_text=True`; otherwise they're flattened to spaces.

**Border properties:**

Each cell has four sides (`top`, `bottom`, `left`, `right`), and each side has a `style` and optional `color`. PDF maps Excel border styles to line widths:

| Border Style | Rendered Width (PDF points) |
|--------------|-----------------------------|
| `hair`       | 0.25                        |
| `thin`       | 0.5                         |
| `medium`     | 1.0                         |
| `thick`      | 1.5                         |
| `dashed`     | 0.75                        |
| `dotted`     | 0.5                         |
| `double`     | 1.25                        |

Borders on merged cells are drawn around the **full merged region**, not just the anchor cell. `start`/`end` sides are reading-order-aware, so right-to-left sheets resolve their left/right edges correctly.

**Merged cells and gridlines:**

Merged regions are preserved in both formats: re-applied in XLSX fidelity export, and drawn as `SPAN` commands in PDF. The template's `show_gridlines` setting carries through to XLSX output.

<div class="admonition note">
<p class="admonition-title">Not rendered in PDF</p>
<p><code>text_rotation</code> and diagonal borders are captured in the schema for fidelity but are not drawn by the PDF renderer. They render correctly in XLSX.</p>
</div>

## Troubleshooting

- **`hug` was rejected.**
  `hug` isn't available in streaming mode, and column `hug` isn't available for PDF `dataframe-content`. Use `fixed` or `even` there.
- **A fill didn't show up.**
  Only solid fills are extracted. Patterned fills are skipped by design.
- **A theme color looks slightly off in PDF.**
  PDF resolves theme colors against the default Office palette. For exact control, set an explicit hex color in the template.
