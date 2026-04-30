from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

import mindoff_dataport as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
DATA_PARQUET = HERE / "data.parquet"
OUTPUT_PDF = HERE / "output.pdf"


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET).select(
        ["Employee Name", "Department", "Amount"]
    )
    dataframe_options = {
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

    bundle = mo_dataport.compile(
        schema,
        {
            "Column Layout": {
                "report_title": "Dataframe Column Occupation",
                "headers": rows,
                "rows": rows,
            }
        },
        dataframe_options=dataframe_options,
    )
    mo_dataport.export(
        bundle,
        str(OUTPUT_PDF),
        format="pdf",
        streaming_chunk_rows=2,
        orientation="portrait",
        row_height_mode="fixed",
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_PDF}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
