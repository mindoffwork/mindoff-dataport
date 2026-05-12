from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

import mindoff_dataport as mo_dataport

HERE = Path(__file__).resolve().parent
OUT = HERE / "output"
TEMPLATE = HERE / "template.xlsx"
DATA = HERE / "data.parquet"

# Per-column layout: occupation (column span) and alignment for headers vs rows.
DATAFRAME_OPTIONS = {
    "Column Layout": {
        "headers": {
            "columns": {
                "Employee Name": {"occupation": 2, "alignment": "center"},
                "Department": {"occupation": 2, "alignment": "center"},
                "Amount": {"occupation": 1, "alignment": "center"},
            }
        },
        "rows": {
            "columns": {
                "Employee Name": {"occupation": 2, "alignment": "left"},
                "Department": {"occupation": 2, "alignment": "center"},
                "Amount": {"occupation": 1, "alignment": "right"},
            }
        },
    }
}


def main() -> None:
    started = perf_counter()
    OUT.mkdir(exist_ok=True)

    schema = mo_dataport.extract(str(TEMPLATE))
    rows = pl.scan_parquet(DATA).select(["Employee Name", "Department", "Amount"])

    # Template uses dataframe-header + dataframe-content split anchors.
    # dataframe_options controls per-column occupation and alignment independently.
    bundle = mo_dataport.compile(
        schema,
        {
            "Column Layout": {
                "report_title": "Dataframe Column Options",
                "headers": rows,
                "rows": rows,
            }
        },
        dataframe_options=DATAFRAME_OPTIONS,
    )

    xlsx_out = mo_dataport.export(
        bundle,
        str(OUT / "output.xlsx"),
        export_mode="streaming",
        streaming_chunk_rows=2,
    )
    mo_dataport.export(
        bundle,
        str(OUT / "output.pdf"),
        format="pdf",
        streaming_chunk_rows=2,
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE}")
    print(f"Parquet:  {DATA}")
    print(f"XLSX:     {xlsx_out[0]}")
    print(f"PDF:      {OUT / 'output.pdf'}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
