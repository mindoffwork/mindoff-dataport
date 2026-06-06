# Exporting to PDF

The same bundle that produces your Excel file produces your PDF. No second template, no separate styling pass. Pass `format="pdf"` and the library renders with ReportLab, starting each sheet on a fresh page and paginating long tables automatically.

```python
mo_dataport.export(bundle, "report.pdf", format="pdf")
```

PDF always paginates, so there's no `export_mode` to choose; the renderer streams large tables into bounded chunks on its own. Your job is to set the page geometry and, optionally, the fonts.

## Prerequisites

- A compiled `ReportBundle` or a path to a persisted bundle directory (see [Compiling Report Bundles](compiling-bundles.md))
- Font files on disk if using custom fonts (TrueType `.ttf` or OpenType `.otf`)

## Implementation

### 1. Configure PDF Options

| Option                     | Type           | Default      | Description                                                                       |
|----------------------------|----------------|--------------|-----------------------------------------------------------------------------------|
| `page_size`                | `str`          | `"A4"`       | `"A4"`, `"LETTER"`, or `"LEGAL"`                                                  |
| `orientation`              | `str`          | `"portrait"` | `"portrait"` or `"landscape"`                                                     |
| `margin`                   | `float`        | `36`         | Page margin in points (0 or greater), applied to all four sides                   |
| `streaming_chunk_rows`     | `int`          | `50000`      | Rows read per batch for `dataframe-content` and repeat sections                   |
| `fonts`                    | `dict \| None` | `None`       | Custom TrueType / OpenType font families. See [Register Custom Fonts](#3-register-custom-fonts) below |
| `repeat_dataframe_headers` | `bool`         | `False`      | Repeat dataframe header rows across paginated table chunks (opt-in)               |
| `column_width_mode`        | `str`          | schema value | `"fixed"` or `"even"` for `dataframe-content` sheets (no `"hug"` for columns)    |
| `row_height_mode`          | `str`          | schema value | `"fixed"`, `"even"`, or `"hug"` (PDF supports `"hug"` for `dataframe-content` rows) |
| `default_column_width`     | `float`        | schema value | Same as XLSX                                                                      |
| `default_row_height`       | `float`        | schema value | Same as XLSX                                                                      |

<div class="admonition note">
<p class="admonition-title">A note on sizing in PDF</p>
<p>For sheets that render <code>dataframe-content</code>, PDF allows <code>row_height_mode="hug"</code> (it auto-fits each streamed chunk) but <strong>not</strong> <code>column_width_mode="hug"</code>; measuring column widths would mean buffering every row first, which defeats streaming. See <a href="sizing-and-styling.md">Sizing &amp; Styling</a>.</p>
</div>

### 2. Repeat Headers Across Pages

By default a dataframe header is written once. When a long table spills across several PDF pages, you often want the column headers to reappear at the top of each page. Set `repeat_dataframe_headers=True`; headers repeat across later chunks wherever a matching `dataframe-header` anchor exists.

```python
mo_dataport.export(
    bundle,
    "report.pdf",
    format="pdf",
    repeat_dataframe_headers=True,
)
```

### 3. Register Custom Fonts

By default, every cell font maps to ReportLab's built-in **Helvetica** family. To render with your own fonts, pass a `fonts` dict. The renderer matches each cell's template `font.name` (case-sensitive) against your keys; unmatched names fall back to Helvetica.

**Shorthand: regular only.** When you only have a regular weight, give a single path. Bold and italic fall back to it.

```python
mo_dataport.export(
    bundle,
    "report.pdf",
    format="pdf",
    fonts={
        "Inter": "/path/to/fonts/Inter-Regular.ttf",
    },
)
```

**Full variant map.** For distinct weights and styles, give a dict per family:

```python
mo_dataport.export(
    bundle,
    "report.pdf",
    format="pdf",
    fonts={
        "Inter": {
            "regular":     "/path/to/fonts/Inter-Regular.ttf",
            "bold":        "/path/to/fonts/Inter-Bold.ttf",
            "italic":      "/path/to/fonts/Inter-Italic.ttf",
            "bold_italic": "/path/to/fonts/Inter-BoldItalic.ttf",
        }
    },
)
```

| Key           | Required | Falls back to                |
|---------------|----------|------------------------------|
| `regular`     | Yes      | (none)                       |
| `bold`        | No       | `regular`                    |
| `italic`      | No       | `regular`                    |
| `bold_italic` | No       | `bold`, then `regular`       |

You can register several families in one call:

```python
fonts={
    "Inter": {...},
    "Roboto Mono": "/path/to/RobotoMono-Regular.ttf",
}
```

Font rules:

- Files must exist on disk when `export()` is called; a missing path raises `ValueError`.
- Every family **must** supply a `regular` file; omitting it raises `ValueError`.
- Fonts are registered with ReportLab once per process; re-registering the same path is a harmless no-op.

### 4. Understand What Renders

PDF reproduces the great majority of Excel styling: fonts, fills (solid), alignment (including `justify`/`distributed`), borders (drawn around full merged regions), strikethrough, superscript/subscript, and indents. A couple of things are captured in the schema but **not** drawn in PDF: `text_rotation` and diagonal borders. The complete matrix is in [Sizing & Styling](sizing-and-styling.md#supported-styling).

## Core Concepts

### 1. Why PDF Has No Export Mode

PDF always paginates. There is no fidelity-vs-streaming choice because the renderer has no alternative to chunked output: a table taller than a page must be split across pages regardless. Chunking (`streaming_chunk_rows`) is always active; it controls how many rows are processed at a time, but it is never optional. You configure the page geometry and fonts; the renderer handles breaking and layout.

### 2. How Automatic and Manual Page Breaks Interact

Two independent mechanisms control where pages break. Automatic height-based breaks happen when a rendered chunk fills the printable page area and the next chunk starts on a new page. Manual row breaks (`row_page_breaks`, set in the template via **Page Layout → Breaks** and re-resolved at compile time) force a new page at a specific row regardless of how full the current page is. Both apply simultaneously: a table can hit a manual break mid-chunk and a height-based break at the end of a full chunk, and the renderer handles both correctly.

## Troubleshooting

- **`ValueError` about a font.**
  Either a font file path doesn't exist, or a family is missing its required `regular` entry.
- **My custom font isn't applied.**
  Matching is case-sensitive against the template cell's `font.name`. Confirm the name in Excel matches your `fonts` key exactly.
- **Headers don't repeat on later pages.**
  Set `repeat_dataframe_headers=True` and make sure the sheet uses a `dataframe-header` anchor (not a bare `dataframe`).
