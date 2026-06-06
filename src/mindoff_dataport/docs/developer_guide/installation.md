# Installation

Getting `mindoff-dataport` onto your machine takes one command. Adding dataframe support takes one more. This page covers both, plus the handful of things worth knowing before you build your first report.

## Prerequisites

- Python `{{ PYTHON_REQUIRES }}` (tested on {{ PYTHON_SUPPORTED }})
- `pip`

That's it. The core dependencies (`openpyxl`, `xlsxwriter`, `reportlab`, and `pyarrow`) install automatically and cover Excel reading/writing, PDF rendering, and Parquet storage.

## Implementation

### 1. Install the Package

```bash
pip install mindoff-dataport
```

This pulls in everything needed to extract templates and export both XLSX and PDF.

### 2. Add Dataframe Support (Optional)

If you plan to feed tables into your reports (and most reports have a table somewhere), install the `polars` extra:

```bash
pip install "mindoff-dataport[polars]"
```

<div class="admonition info">
<p class="admonition-title">Why is Polars optional?</p>
<p>The core library has no opinion about where your tabular data comes from. Polars is only required when you actually pass a <code>DataFrame</code> or <code>LazyFrame</code> into <code>compile()</code>. Keeping it optional means lighter installs for projects that only render scalar reports.</p>
</div>

### 3. Verify the Install

Open a Python shell and confirm the entry point imports:

```python
from mindoff_dataport import mo_dataport
print(mo_dataport)
```

If that runs without an `ImportError`, you're ready.

## The Four-Step Workflow

Every report follows the same sequence: **extract** the template, **inspect** what it needs, **compile** your data into it, and **export** to a file. The [Quick Start](/#quick-start) walks through all four steps with a full code example. Once you've run your first report, continue with [Templates & Placeholders](templates-and-placeholders.md) to mark up your own Excel files.

## Installing from Source

To work against the latest code or run the examples, clone the repository and install it in editable mode:

```bash
git clone https://github.com/mindoffwork/mindoff-dataport
cd mindoff-dataport
pip install -e ".[dev]"
```

The `dev` extra adds `pytest`, coverage tooling, and Polars so you can run the test suite and every example end to end.

## Troubleshooting

- **`ModuleNotFoundError: No module named 'polars'`**
  You passed a dataframe into `compile()` without the Polars extra. Install it with `pip install "mindoff-dataport[polars]"`.
- **`ImportError` on `mindoff_dataport`**
  Make sure you installed into the same interpreter you're running. A mismatched virtual environment is the usual cause.
