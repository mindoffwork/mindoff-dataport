# PDF Rendering

The PDF renderer produces a paginated document from the same bundle that drives Excel output. It's built on ReportLab and owned by `pdf_renderer.py`. Where the XLSX renderer offers a choice of modes, PDF has only one job that it always does: lay the report out on fixed-size pages and break cleanly across them.

For the user-facing options, see [Developer Guide → Exporting to PDF](../developer_guide/exporting-pdf.md).

## How PDF Differs from XLSX

PDF and XLSX consume the **same resolved layout and style plan**. The differences are deliberately narrow and deterministic:

- PDF paginates; XLSX does not. Each sheet starts on a new page, and overflow rows flow onto following pages.
- PDF differences from XLSX are limited to supplied font availability and deterministic page scaling; nothing about the layout is reinterpreted.

## Pagination & Streaming Chunks

`dataframe-content` is streamed into **bounded table chunks** rather than built as one giant table. Each chunk is sized to fit the printable page height, and when fixed or even row heights are active, streamed chunks still respect that printable height so a row is never split across a page boundary. This is what lets PDF export arbitrarily large tables without buffering them.

A correctness detail the renderer guards: when flushing chunks by height, multi-row merges that straddle a chunk boundary must not be dropped; they are carried correctly into the next chunk.

### 1. Manual Page Breaks

PDF uses the compile-resolved `row_page_breaks` as **manual page boundaries**, inserting a new page before the affected rows. It ignores `column_page_breaks` (those are an XLSX concern). Combined with automatic height-based pagination, this gives both intentional and overflow page breaks.

### 2. Repeating Headers

By default a dataframe header is written once. With `repeat_dataframe_headers=True`, header rows repeat across later paginated chunks wherever a matching `dataframe-header` anchor exists, so a table that spans pages keeps its column labels at the top of each page. Default remains `False`.

## Sizing in PDF

PDF honors the shared sizing modes with two table-specific rules for `dataframe-content`:

- `column_width_mode="hug"` is **rejected** because measuring column widths would require buffering all rows, defeating streaming.
- `row_height_mode="hug"` is **allowed**; each streamed chunk auto-sizes its row heights.

## Style Rendering Fidelity

ReportLab can't express every Excel style natively, so the renderer maps them carefully:

| Excel style                              | PDF rendering |
|------------------------------------------|---------------|
| `strike`, `vert_align` (super/subscript) | ReportLab paragraph markup |
| `underline`                              | `<u>` markup |
| `indent` / `relative_indent`             | cell padding |
| `justify` / `distributed` alignment      | `TA_JUSTIFY` |
| pattern fills                            | approximated by background color |
| `text_rotation`, diagonal borders        | **captured in schema but not rendered** |

Borders draw around full merged regions, and `start`/`end` sides resolve against `reading_order` so right-to-left sheets get their left/right edges correct. Cell fonts map to ReportLab's built-in **Helvetica** family unless a matching custom font is supplied via the `fonts` option (TrueType/OpenType, registered once per process). Theme colors are resolved against the default Office palette.

## Borders Over Empty Space

A deliberate non-behavior worth calling out: PDF draws **only template borders**. It does not paint a default grid over empty spacer cells, so blank areas of your layout stay blank instead of acquiring stray lines.

## Troubleshooting PDF Rendering

1. **A table row split across a page.** Shouldn't happen; chunks respect printable height. If you see it, check for an unusually tall single row exceeding the page height.
2. **A custom font didn't apply.** Matching is case-sensitive against the template cell's `font.name`; confirm the name and the required `regular` file.
3. **Rotated text or a diagonal border is missing.** Those are captured but not drawn in PDF by design; they render in XLSX.
