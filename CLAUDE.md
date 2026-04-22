# CLAUDE.md - Canonical Agent Guide (Token-Optimized)

Purpose: keep prompts short, edits safe, tests focused, and this file always in sync.

## 1) Prompt Contract (lowest-token default)

Use this compact format for all feature/edit/test requests:

```text
Task: <one sentence>
Scope: <files or module>
Constraints: <compat/perf/api limits>
Tests: <exact pytest target(s)>
Done-When: <observable acceptance>
```

Rules:
- Prefer one request = one clear outcome.
- Avoid repeating repo context already in this file.
- Ask for diffs + targeted tests, not broad rewrites.
- If uncertain, choose smallest safe change.

## 2) Project Snapshot (minimal)

- Package: `mindoff_data_export`
- Core flow: `extract_template(.xlsx) -> schema -> build_template_with_data(..., output_path)`
- Runtime templating: `render_schema`, `build_template_with_data`
- Main modules:
  - `schema.py` TypedDict schemas
  - `extractor.py` xlsx -> schema
  - `builder.py` schema -> xlsx
  - `renderer.py` `{{key:type}}` replacement
  - `streaming.py` bounded-memory export path
  - `utils.py` color/border helpers

## 3) Public API (stable unless explicitly changed)

- `extract_template(path)`
- `build_template_with_data(schema, data, output_path, **sizing_kwargs)`
- `get_template_inputs(schema)`
- `render_schema(schema, data)`

Notes:
- `build_template_with_data` is the only top-level build/export API.
- Low-level workbook reconstruction remains in `mindoff_data_export.builder.build_template` for internal/testing use.

`build_template_with_data` now supports:
- `export_mode="fidelity" | "streaming"` (default: fidelity)
- `streaming_chunk_rows` and `max_rows_per_workbook` when streaming
- return type: `None` (fidelity) or `list[str]` (streaming chunk outputs)

## 4) Non-Negotiable Invariants

- `merged_regions` is authoritative during build.
- Preserve formulas (`data_only=False` behavior).
- `render_schema` must not mutate input.
- Use `fgColor` for solid fills.
- JSON row keys can be `str`; builder converts to `int`.
- Openpyxl styles are immutable; build new style objects.
- Streaming mode is intentionally constrained:
  - no `hug` sizing
  - no merged-cell output
  - one `dataframe-content` placeholder per sheet
  - large `dataframe-content` may split across `*.partNNN.xlsx` files

## 5) Sizing Modes

- `fixed`: use stored widths/heights.
- `even`: use defaults (`default_column_width`, `default_row_height`).
- `hug`: compute from written cell content.
- Function kwargs override per-sheet schema values.
- In streaming mode, `hug` is invalid; only `fixed`/`even` are allowed.

## 6) Placeholder Types

Supported:
- Scalars: `string`, `number`, `date`
- Dataframe: `dataframe-headers`, `dataframe-content`

Behavior:
- `get_template_inputs` returns `{key: type}`.
- `render_schema` validates and resolves placeholders.
- `dataframe-headers` writes header cells only; accepts DataFrame/LazyFrame or list input.
- `dataframe-content` writes row content only; accepts DataFrame/LazyFrame.
- In streaming mode, `dataframe-content` is written incrementally (LazyFrame batches first).

## 7) Test Policy (token-efficient)

Default command:

```bash
PYTHONPATH=src python -m pytest -q
```

Temporary test artifacts:
- Tests create per-test intermediate folders under a writable temp root.
- `tests/conftest.py` auto-falls back to project `.tmp/` if OS temp is not writable.
- Pytest base temp is set to a unique per-run folder to avoid stale ACL/cleanup collisions on Windows.
- On Windows/Python 3.13 test runs, `tests/conftest.py` patches `os.mkdir(mode=0o700)` to a safe mode due ACL access errors.
- Per-test intermediate folders are auto-removed after each test.

Change-scoped first, then broader only if needed:
1. Run nearest test file(s).
2. Run roundtrip guard: `tests/test_roundtrip.py::test_roundtrip_schema_identical`.
3. Run full suite if behavior touched multiple modules.

Opt-in heavy checks:
- None.

## 8) Auto-Update Policy (mandatory for all agents)

When any change affects behavior, API, schema, placeholder rules, sizing logic, or tests, the agent MUST update this `CLAUDE.md` in the same change set.

Required update checklist:
1. Update relevant section(s) above.
2. Keep wording compact; remove stale lines.
3. If new behavior is temporary/experimental, label it clearly.
4. Ensure examples/tests mentioned here still exist.

Fail condition:
- A PR/edit that changes behavior but leaves `CLAUDE.md` stale is incomplete.

## 9) Feature/Edit/Test Execution Order

1. Read target module + nearest tests.
2. Implement minimal diff.
3. Add/adjust tests for behavior.
4. Run targeted tests.
5. Update `CLAUDE.md` per Auto-Update Policy.
6. Run roundtrip guard if serialization/build path changed.

## 10) Reference Commands

```bash
# targeted
PYTHONPATH=src python -m pytest -q tests/test_renderer.py

# roundtrip guard
PYTHONPATH=src python -m pytest -q tests/test_roundtrip.py::test_roundtrip_schema_identical

# examples
python examples/demo.py
```

## 11) Code Organization Philosophy (sectioned, grep-friendly)

Use numbered `§` section comments inside Python modules so structure is predictable and easy to reference in reviews.

Module order:
1. Imports
2. `# §1 Types` (TypedDict, Protocol, aliases, dataclasses)
3. `# §2 Constants`
4. `# §3 Private Helpers`
5. `# §4 Public API` (exported functions/classes)

Rules:
- Keep this order for `src/` and `tests/` modules when practical.
- Use subsection comments only for non-trivial logic (example: `# §3.1 Normalize color token`).
- Prefer small, targeted files; extract helpers only when readability clearly improves.
- Cross-reference as: `builder.py §3` or `renderer.py §4.2`.

## 12) Commit Message Guidelines

Use gitmoji codes in commit subjects that match the change type. Keep the text short and action-oriented. Favor clarity over creativity. Keep subject <= 72 chars.

Format:

```text
:<gitmoji>: <Imperative summary>
```

Recommended mappings:
- `feat` -> `:sparkles:`
- `fix` -> `:bug:`
- `refactor` -> `:recycle:`
- `test` -> `:test_tube:`
- `docs` -> `:memo:`
- `chore` -> `:wrench:`

Examples:
- `:sparkles: Add bulk update validation`
- `:bug: Fix merged border regression`
- `:test_tube: Add renderer coverage`
- `:memo: Update CLAUDE commit rules`

