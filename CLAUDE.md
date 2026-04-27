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

- Package: `mindoff_data_export`
- Flow: `extract_template(.xlsx) -> schema -> compile_report_bundle(...) -> export_report_bundle(...)`
- Main modules: `schema.py`, `extractor.py`, `bundle.py`, `xlsx_renderer.py`, `pdf_renderer.py`, `builder.py`, `renderer.py`, `utils.py`

## 3) Public API

Stable unless explicitly changed:

- `extract_template(path)`
- `get_template_inputs(schema)`
- `compile_report_bundle(template, data, bundle_path=None)`
- `export_report_bundle(bundle_or_path, output_path, format="xlsx", **options)`
- `parquet_source(path, *, columns=None, row_count=None)`
- `mode.extract(path)`
- `mode.inputs(template)`
- `mode.compile(template, data, bundle_path=None)`
- `mode.export(bundle_or_path, output_path, format="xlsx", **options)`

Notes:

- `ReportBundle` directory is the canonical intermediate artifact.
- `report.json` resolves scalar/static cells and stores dataframe anchors; it must not expand dataframe rows into cell schemas.
- `pyarrow>=15.0` is required; dataframe sources are stored as `data/*.parquet`.
- `parquet_source(...)` is the disk-backed input for larger-than-RAM data.
- `format="xlsx"` and `format="pdf"` are implemented. `format="image"` raises `NotImplementedError`.
- PDF export uses ReportLab, starts each sheet on a new page, and paginates overflow rows vertically.
- PDF export supports optional custom TrueType/OpenType fonts via the `fonts` option.
- PDF export draws only template borders; it must not add a default grid over empty spacer cells.
- `builder.py` contains XLSX style/sizing helper functions used by `xlsx_renderer.py`.
- Current input contract is sheet-scoped data, not flat key/value payloads.
- XLSX export supports `export_mode="fidelity" | "streaming"` and streaming may return `list[str]`.

## 4) Invariants

- `merged_regions` is authoritative during build.
- Merged-cell borders must render around the full merged region, not only the anchor cell.
- Preserve formulas (`data_only=False`).
- Bundle compilation must not mutate input templates.
- Use `fgColor` for solid fills.
- Preserve sheet gridline visibility via `show_gridlines`.
- Builder converts JSON row keys from `str` to `int`.
- Openpyxl styles are immutable; create new style objects.
- Streaming limits: no `hug`, no merged cells intersecting dataframe-content output, one `dataframe-content` placeholder per sheet.

## 5) Sizing

- `fixed`: stored widths/heights.
- `even`: defaults.
- `hug`: content-based sizing.
- Kwargs override schema values.
- Streaming forbids `hug`.

## 6) Placeholders

Supported:

- Scalars: `string`, `number`, `date`
- Dataframes: `dataframe-headers`, `dataframe-content`

Behavior:

- `get_template_inputs` returns a sheet-scoped contract.
- Bundle compilation validates per sheet payload.
- Sheet order follows template order; dynamic groups follow payload order.
- Output sheet names must stay unique.
- `dataframe-headers` writes headers only.
- `dataframe-content` writes rows only.
- Streaming writes `dataframe-content` incrementally from parquet batches.
- `auto_delete_bundle=True` deletes the bundle directory only after successful export.

## 7) Tests

Default:

```bash
PYTHONPATH=src python -m pytest -q
```

Run in order:

1. Nearest test file(s)
2. `tests/test_roundtrip.py::test_roundtrip_schema_identical`
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
PYTHONPATH=src python -m pytest -q tests/test_bundle.py tests/test_public_api.py
PYTHONPATH=src python -m pytest -q tests/test_roundtrip.py::test_roundtrip_schema_identical
python examples/xlsx_output.py
```

## 11) Code Organization

Use `# §N Name` sections in Python files.

**Required order:**

1. Imports (stdlib → third-party → local)
2. Module metadata (`__all__`, `__version__`, etc.)
3. `# §1 Constants & Exceptions`
4. `# §2 Classes and Sub Classes`
5. `# §3 Private Helper Functions`
6. `# §4 Public Functions`
7. `# §5 Entrypoints`

**Rules:**

- Keep subsection comments only for major logic/execution phases
- Prefix private helpers with `_`
- Keep related classes grouped; avoid scattering
- Prefer splitting files over overgrowth (> ~500–800 LOC)
- Use them for meaningful steps like parse, validate, transform, write, batch, finalize.
- Prefer `# §3.1 Parse sheet metadata` style comments when a phase matters.
- Keep `src/` and `tests/` aligned when practical.
- Cross-reference by section label, for example `renderer.py §4.2`.

## 12) Git Commit Standard

Use `:<gitmoji_code>: <Verb> <short action-oriented description>`.
