# API Reference

`mindoff-dataport` has a deliberately small public surface: four functions that mirror the four steps of building a report. This page documents each one in full. The reference below is generated directly from the docstrings in the source, so it always matches the version you have installed.

## Import Alias

The recommended entry point bundles all four functions under one namespace:

```python
from mindoff_dataport import mo_dataport

mo_dataport.extract(...)   # extract_template
mo_dataport.inputs(...)    # get_template_inputs
mo_dataport.compile(...)   # compile_report_bundle
mo_dataport.export(...)    # export_report_bundle
```

Every function is also importable at the top level under both its full name and a short alias:

```python
from mindoff_dataport import (
    extract_template,        # alias: extract
    get_template_inputs,     # alias: inputs
    compile_report_bundle,   # alias: compile
    export_report_bundle,    # alias: export
)
```

## Functions

{{ DATAPORT_PUBLIC_API }}

## Supporting Types

### `ReportBundle`

The canonical intermediate artifact produced by `compile()` and consumed by `export()`. It can live in memory or be persisted to a directory and reloaded later:

```python
from mindoff_dataport import ReportBundle

bundle = ReportBundle.load("saved_bundle")   # reload a persisted bundle
mo_dataport.export(bundle, "report.xlsx")
```

The on-disk layout is documented in [Architecture → The Pipeline](../architecture/pipeline.md#the-reportbundle-directory).

### `repeat_records(...)`

A helper that wraps source-backed repeat payloads (an ordered set of scalar record columns plus constant dataframe payloads) so large repeat sections don't have to materialise every record in memory.

```python
from mindoff_dataport import repeat_records

records = repeat_records(scalar_records, constants={"line_items": shared_df})
```

See the [Repeat Sections recipe](recipes.md#2-repeat-sections-per-customer-blocks) for context on when to reach for it.
