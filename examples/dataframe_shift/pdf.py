from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
DATA_PARQUET = HERE / "data.parquet"
OUTPUT_PDF = HERE / "output.pdf"


def _report_payload(rows: pl.LazyFrame) -> list[dict]:
    reports = [
        {
            "customer_name": "Alpha Team",
            "region": "North",
            "line_items": rows,
            "note": "Shift right of dataframe",
            "footer": "Shift below dataframe",
        },
        {
            "customer_name": "Beta Team",
            "region": "South",
            "line_items": rows,
            "note": "Repeat block stays aligned",
            "footer": "Footer follows dataframe rows",
        }
    ]
    return reports


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET).select(["Employee", "Amount"])
    bundle = mo_dataport.compile(
        schema,
        {"Shift Demo": {"reports": _report_payload(rows)}},
        dataframe_shift="both",
    )
    mo_dataport.export(
        bundle,
        str(OUTPUT_PDF),
        format="pdf",
        streaming_chunk_rows=1,
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_PDF}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
