# Compilation & ReportBundle

Compilation is where a template and a payload become a `ReportBundle`. It is the brain of the pipeline: it validates inputs, resolves what it can immediately, defers what should stay lazy, and produces a plan compact enough to persist and portable enough to render anywhere. Two modules share the work: `template_contract.py` (the contract and validation side) and `bundle.py` (the assembly and layout side).

## Division of Labor

| Module                  | Responsibility |
|-------------------------|----------------|
| **`template_contract.py`** | Placeholder discovery, building the input contract, payload validation, scalar substitution, and resolving each sheet's payload (static, repeat, dynamic) |
| **`bundle.py`**            | Materialising dataframes to Parquet, computing dataframe anchors and column layouts, planning repeat sections, collision shifting, and writing the bundle directory |

## The Contract Side (`template_contract.py`)

Before any data is bound, the contract layer answers "what does this template need?" That is `get_template_inputs`, the sheet-scoped contract you see from `mo_dataport.inputs(...)`. At compile time it then validates the payload against that contract:

- **Type checks**: each scalar value must match its placeholder type; dataframes must be dataframe-like.
- **Repeat payloads**: validated as ordered record lists, including source-backed repeats that must not materialise every record.
- **Sheet resolution**: static sheets, repeat sections, and dynamic `{{key}}` sheet groups are each resolved to concrete per-sheet payloads, in the right order.

Scalar substitution happens here too: a resolved scalar is written straight into its cell, inheriting the template's style. Wrong types and missing keys fail at this stage, before any file is touched. That is the whole point of having a contract.

## The Assembly Side (`bundle.py`)

Once the payload is valid, `bundle.py` builds the artifact.

### 1. Dataframes Become Anchors and Parquet

Each dataframe input is written to `data/*.parquet`, and the plan stores only a compact **anchor**: the column names, the start row/column, the style to apply, and per-column `column_layouts` (occupation widths). The rows themselves never enter `report.json`. Source-backed and lazy inputs are honored; a `LazyFrame` stays disk-backed and is read in batches at export.

### 2. Repeat Planning

Repeat sections are planned, not unrolled. `bundle.py` identifies the content rows, static rows, and merged regions of each block, then stores compact `cell_templates` and a repeat plan. Source-backed repeats use compact representations so a large repeat doesn't balloon the bundle. The repeat rules enforced here:

- One or more **non-overlapping sibling** vertical sections per sheet; no nesting.
- Static rows allowed before, between, and after sections.
- Unique repeat keys per sheet.
- Merged cells permitted in fixed/static rows, but not over `dataframe-content` rows.

### 3. Collision Shifting

When dataframe output will expand into occupied template space, `dataframe_shift` moves the colliding cells and merged regions out of the way (`"both"`, `"horizontal"`, `"vertical"`, or `"none"`). Two things make this safe and cheap:

1. **It's metadata-only.** Cells and merges move in the plan; dataframe rows stay in Parquet, so nothing is materialised just to shift it. Streaming still reads in batches afterward.
2. **It's shared.** The same shifted layout drives both XLSX and PDF, including later dataframe anchors inside repeat blocks.

A merge that genuinely cannot be moved clear of dataframe output raises `ValueError` rather than producing a broken layout. With `"none"`, any overlapping template merge fails fast.

### 4. Page Break Resolution

The manual breaks captured at extraction are re-resolved here against the *final* layout (after dataframe expansion and shifting), so a break set on a template row still lands at the intended logical boundary even when a table pushed that row far down.

## What Lands in the Bundle

The result is the directory described in [The Pipeline](pipeline.md#the-reportbundle-directory): `manifest.json` (version, inputs, sheet metadata, sources, capabilities), `report.json` (resolved cells + anchors + repeat plans), and `data/*.parquet`. `dataframe_options` is stored keyed by resolved sheet name then placeholder key, carrying only compact anchor `column_layouts`.

<div class="admonition note">
<p class="admonition-title">Authoritative merges</p>
<p><code>merged_regions</code> is authoritative during build. The one exception is renderer-owned dataframe <code>occupation</code> merges, which are generated from anchor metadata at render time rather than stored in the plan.</p>
</div>

## Troubleshooting Compilation

1. **`KeyError`**: a required placeholder key is missing; diff your payload against `inputs(schema)`.
2. **`ValueError` on a merge**: a merged region overlaps dataframe output and can't be shifted; change `dataframe_shift` or redesign the template.
3. **A type validation failure**: the offending key is named in the error; check the value against its placeholder type.
4. **A repeat rejected**: re-check the repeat rules above (overlap, nesting, merges over content rows).
