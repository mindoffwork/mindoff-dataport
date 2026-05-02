from __future__ import annotations

import datetime as dt
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
    rows = pl.scan_parquet(DATA_PARQUET)
    bundle = mo_dataport.compile(
        schema,
        {
            "Page Break Demo": {
                "generated_on": dt.date(2026, 5, 2),
                "items": rows,
                "footer_note": "The manual row break is resolved after dataframe expansion for PDF export.",
            }
        },
        dataframe_shift="vertical",
    )
    mo_dataport.export(
        bundle,
        str(OUTPUT_PDF),
        format="pdf",
        streaming_chunk_rows=2,
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
