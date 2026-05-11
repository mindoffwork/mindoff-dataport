# Public Benchmark Guide

## Why use mindoff dataport?

- **No manual styling code.** Design the report once in Excel; the library applies your template to any dataset at runtime.
- **Schema-validated inputs.** Errors are caught before any file is written, not after a 30-second export.
- **One template, two formats.** The same Excel template drives both XLSX and PDF output.

The benchmark measures the runtime cost of this template-driven approach and compares it against writing the same output by hand with raw libraries.

---

## Benchmark goal

Show runtime performance and memory efficiency for producing styled XLSX and PDF outputs from a prepared template workflow.

Primary claim:

> Mindoff Dataport is competitive in runtime export performance while preserving template-driven output fidelity.

---

## What is measured

- Runtime: compile + export combined.
- XLSX and PDF output paths.
- Mindoff mode guidance (fidelity, streaming-openpyxl, streaming-xlsxwriter).

## What is not measured

- Template design-time productivity (how fast a team can build or modify report layouts in Excel).
- Template extraction / schema creation - this is a one-time setup step, excluded from timed results.

---

## Method set

**XLSX**

| Method | What it does |
|---|---|
| mindoff fidelity | Template-driven workflow, full style fidelity |
| mindoff streaming-openpyxl | Template-driven workflow, streaming output |
| mindoff streaming-xlsxwriter | Template-driven workflow, streaming with xlsxwriter engine |
| openpyxl - manual styling | Raw openpyxl write with all styles hardcoded in Python code |
| xlsxwriter - manual styling | Raw xlsxwriter write with all styles hardcoded in Python code |

**PDF**

| Method | What it does |
|---|---|
| mindoff export | Template-driven PDF with pagination and print-break handling |
| reportlab - manual styling | Raw ReportLab report with the same visible title, subtitle, table, sizing, and styles hardcoded in Python code |

> **Note on direct-library baselines:** "manual styling" methods hardcode every colour, border, font, row height, and column width directly in Python. They render the same visible benchmark report shape, but no template is read. These are the speed ceiling for pure writing with no design layer. Mindoff adds a compile step on top of this; that step is what converts the Excel template into styled output at runtime.

---

## Output artifact

Running `python examples/benchmark/run.py` produces:

| File | Description |
|---|---|
| `output/results.csv` | Results per method and row count, including std dev, compile time, rows/s, output file path, peak MB per 10K rows, and file MB per 10K rows |
| `output/files/` | Representative XLSX and PDF outputs from the final measured run for each method and row count, grouped by format |

---

## Reliability and fairness

- Each benchmark point reports the **median** value across repeated runs.
- One **warm-up pass** is discarded before measurement starts, to avoid cold Python import / OS file-cache effects.
- Timeout policy is fixed; timed-out points are clearly marked.
- Direct-library baselines render the same visible data, report structure, and style intent as the template-driven output.
- Saved artifacts are refreshed at the start of each benchmark run so manual comparisons do not mix old and new outputs.

---

## Running the benchmark

```bash
# Quick dev run (2 row scales, 1 run, no warm-up)
python examples/benchmark/run.py --quick

# Full publish run (4 XLSX scales up to 500K rows, 5 runs, warm-up enabled)
python examples/benchmark/run.py --full

# Custom override
python examples/benchmark/run.py --runs 3 --timeout 180
```

If the template is missing, regenerate it first:

```bash
python examples/benchmark/create_template.py
```
