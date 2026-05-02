# CLAUDE.md - Canonical Agent Guide

Keep this file short. Prefer the smallest safe change, the fewest words, and the fewest prompts.

## 1) Prompt Contract

Use:

```text
Task: <one sentence>
Scope: <files or module>
Constraints: <compat/perf/api limits>
Tests: <exact pytest target(s)>
Done-When: <observable acceptance>
```

Rules:

- One request, one outcome.
- Do not repeat repo context already here.
- Ask for targeted diffs and tests.
- If unsure, choose the smallest safe change.

## 2) Project Snapshot

- Project: `mindoff_dataport`
- Import package: `mindoff_dataport`
- Flow: `extract_template(.xlsx) -> schema -> compile_report_bundle(...) -> export_report_bundle(...)`
- Main modules: `schema.py`, `extractor.py`, `template_contract.py`, `bundle.py`, `page_breaks.py`, `xlsx_renderer.py`, `pdf_renderer.py`, `xlsx_builder.py`, `style_conversion.py`

## 3) Public API

Stable unless explicitly changed:

- `extract_template(path)` / `extract(path)`
- `get_template_inputs(schema)` / `inputs(schema)`
- `compile_report_bundle(template, data, bundle_path=None, dataframe_options=None, dataframe_shift="both")` / `compile(template, data, bundle_path=None, dataframe_options=None, dataframe_shift="both")`
- `export_report_bundle(bundle_or_path, output_path, format="xlsx", **options)` / `export(bundle_or_path, output_path, format="xlsx", **options)`

Notes:

- `ReportBundle` directory is the canonical intermediate artifact.
- `report.json` resolves scalar/static cells and stores dataframe anchors/repeat plans; it must not expand dataframe rows into cell schemas.
- `dataframe_options` is keyed by resolved sheet name then placeholder key; it stores compact anchor `column_layouts` only.
- `pyarrow>=15.0` is required; dataframe sources are stored as `data/*.parquet`.
- Polars `LazyFrame` is the disk-backed input for larger-than-RAM data; use `pl.scan_parquet(...)` for Parquet inputs.
- `format="xlsx"` and `format="pdf"` are implemented. `format="image"` raises `NotImplementedError`.
- PDF export uses ReportLab, starts each sheet on a new page, and paginates overflow rows vertically.
- Extracted sheet schemas may include `row_page_breaks` / `column_page_breaks` from manual Excel print breaks; compile resolves them against shifted dataframe layout before export.
- PDF uses resolved `row_page_breaks` as manual page boundaries and ignores `column_page_breaks`; XLSX preserves both resolved row and column breaks.
- PDF export supports optional custom TrueType/OpenType fonts via the `fonts` option.
- PDF export draws only template borders; it must not add a default grid over empty spacer cells.
- PDF renders `strike` and `vert_align` (superscript/subscript) via ReportLab paragraph markup; `indent`/`relative_indent` via cell padding; `justify`/`distributed` alignment via `TA_JUSTIFY`; pattern fills approximated by background color. `text_rotation` and diagonal borders are captured in schema but not rendered in PDF.
- PDF dataframe-content export streams rows into bounded table chunks, rejects `column_width_mode="hug"`, and allows `row_height_mode="hug"`.
- `template_contract.py` owns placeholder discovery, input contracts, payload validation, scalar substitution, and sheet payload resolution.
- `xlsx_builder.py` contains XLSX style/sizing helper functions used by `xlsx_renderer.py`.
- `style_conversion.py` contains openpyxl color/border conversion helpers.
- Current input contract is sheet-scoped data, not flat key/value payloads.
- XLSX export supports `export_mode="fidelity" | "streaming"` and streaming returns `list[str]` containing either a single workbook path or a zip path when split output is produced.

## 4) Invariants

- `merged_regions` is authoritative during build, except renderer-owned dataframe `occupation` merges generated from anchor metadata.
- XLSX/PDF must consume the same resolved dataframe layout/style plan; PDF differences are limited to supplied font availability and deterministic page scaling.
- Merged-cell borders must render around the full merged region, not only the anchor cell. XLSX merged-cell materialization must preserve visible outline/diagonal borders without emitting interior (`horizontal`/`vertical`) merge fields that can suppress left/right edges in Excel.
- Renderer-generated dataframe `occupation` merges must apply anchor border styling on every generated row, not only the first row.
- `start`/`end` border sides are LTR/RTL-aware; PDF resolves them against `reading_order` when drawing left/right edges.
- Preserve formulas (`data_only=False`).
- Bundle compilation must not mutate input templates.
- `FillSchema` uses `pattern_type`, `fg_color` (visible color for solid fills), `bg_color` (pattern background). Old `bg_color`-only shape is handled defensively in the builder but canonical form requires all three keys.
- Preserve sheet gridline visibility via `show_gridlines`.
- Builder converts JSON row keys from `str` to `int`.
- Openpyxl styles are immutable; create new style objects.
- Compile `dataframe_shift` controls template cell/merge movement around dataframe output in both normal sheets and repeat records: `"both"`, `"horizontal"`, `"vertical"`, or `"none"`; merges covering dataframe anchors still fail. Streaming still allows one `dataframe-content` placeholder per non-repeat sheet and no `hug`.
- Streaming repeat-sheet rendering must consume dataframe-content-covered row offsets exactly once; it must not emit extra blank rows after repeated dataframe output.

## 5) Sizing

- `fixed`: stored widths/heights.
- `even`: defaults.
- `hug`: content-based sizing.
- Kwargs override schema values.
- Column-dimension spans from templates are expanded to per-column widths on extraction so ranged widths like `C:CV` survive export.
- For `dataframe-content` in fixed row mode, generated rows inherit the anchor row height when explicit heights are not set for those generated row indexes.
- PDF `dataframe-content` allows row `hug` sizing but still forbids column `hug` sizing.
- Streaming forbids `hug`.

## 6) Placeholders

Supported:

- Scalars: `string`, `number`, `date`
- Dataframes: `dataframe`, `dataframe-header`, `dataframe-content`
- Repeats: `repeat-start`, `repeat-end`

Behavior:

- `get_template_inputs` returns a sheet-scoped contract.
- Bundle compilation validates per sheet payload.
- Sheet order follows template order; exact sheet-name placeholders (`{{key}}`) expand to dynamic groups in payload order.
- Output sheet names must stay unique.
- `dataframe` writes headers at the anchor row and content starting on the next row.
- `dataframe-header` writes headers only.
- `dataframe-content` writes rows only.
- XLSX/PDF dataframe anchors support per-column `occupation` and horizontal alignment (`left`, `center`, `right`) via `dataframe_options`.
- Streaming writes `dataframe-content` incrementally from parquet batches.
- Single-sheet repeats use `{{key:repeat-start}}` / `{{key:repeat-end}}` and require an ordered list of record payloads.
- Repeat v1 supports one or more non-overlapping sibling vertical sections per sheet, no nesting, unique repeat keys, static rows before/between/after sections, and merged cells only in fixed repeat/static rows, not `dataframe-content` rows.
- `auto_delete_bundle=True` deletes the bundle directory only after successful export.

## 7) Tests

Default:

```bash
PYTHONPATH=src python -m pytest -q
```

Coverage gate:

```bash
PYTHONPATH=src python -m pytest --cov=mindoff_dataport --cov-branch --cov-report=term-missing:skip-covered --cov-fail-under=90
```

Run in order:

1. Nearest test file(s)
2. `src/tests/test_roundtrip.py::test_roundtrip_schema_identical`
3. Full suite only if multiple modules changed

## 8) Update Policy

- Any behavior/API/schema/placeholder/sizing/test change must update this file in the same change set.
- Keep edits compact and remove stale text.

## 9) Execution Order

1. Read target module and nearest tests.
2. Make the smallest diff.
3. Add or adjust tests.
4. Run targeted tests.
5. Update `CLAUDE.md`.
6. Run the roundtrip guard if the build/serialization path changed.

## 10) Reference Commands

```bash
PYTHONPATH=src python -m pytest -q src/tests/test_bundle.py src/tests/test_public_api.py
PYTHONPATH=src python -m pytest -q src/tests/test_roundtrip.py::test_roundtrip_schema_identical
python examples/xlsx_output.py
```

## 11) Code Organization

Use `# §N. Name` sections in Python files.

**Required order:**

1. Imports (stdlib → third-party → local)
2. Module metadata (`__all__`, `__version__`, etc.)
3. `# §1. Constants & Exceptions`
4. `# §2. Classes and Sub Classes`
5. `# §3. Private Helper Functions`
6. `# §4. Public Functions`
7. `# §5. Entrypoints`

**Rules:**

- Keep subsection comments only for major logic/execution phases
- Prefix private helpers with `_`
- Keep related classes grouped; avoid scattering
- Prefer splitting files over overgrowth (> ~500–800 LOC)
- Use them for meaningful steps like parse, validate, transform, write, batch, finalize.
- Prefer `# §3.1 Parse sheet metadata` style comments when a phase matters.
- Keep tests under `src/tests/` aligned by test filename with implementation modules when practical.
- Cross-reference by section label, for example `template_contract.py §4.2`.

## 12) Git Commit Standard

Use `:<gitmoji_code>: <Verb> <short action-oriented description>`.
