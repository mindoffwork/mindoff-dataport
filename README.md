# Mindoff Data Export

Extract Excel templates, fill typed placeholders, and rebuild `.xlsx` files with layout and styling preserved.

## What This Library Does

- Extract an Excel workbook into a JSON-like schema
- Detect typed template inputs from placeholders (`{{key:type}}`)
- Render runtime data into the schema safely
- Build final `.xlsx` outputs with styling, merges, formulas, and sizing behavior

## Install

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -e .
```

Optional (for demo and DataFrame examples):

```bash
pip install polars
```

## Quick Start

```python
from mindoff_data_export import (
    extract_template,
    get_template_inputs,
    build_template_with_data,
)

schema = extract_template("template.xlsx")
required_inputs = get_template_inputs(schema)

build_template_with_data(
    schema,
    data={
        "customer_name": "Acme Industries",
        "invoice_number": 1024,
    },
    output_path="filled.xlsx",
    column_width_mode="fixed",
    row_height_mode="fixed",
)
```

## Documentation

- [Setup and Usage](docs/setup.md)
- [Template Authoring](docs/template-authoring.md)
- [API Reference](docs/api-reference.md)
- [Architecture Notes](docs/architecture.md)
- [Testing Guide](docs/testing.md)

## Demo

```bash
python examples/demo.py
```

## Current Scope

- Input template: `.xlsx`
- Output export: `.xlsx`
- Public API: `extract_template`, `build_template`, `build_template_with_data`, `get_template_inputs`, `render_schema`
