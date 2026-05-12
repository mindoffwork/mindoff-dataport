from __future__ import annotations

import datetime as dt
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


def _report_payload(rows: pl.LazyFrame) -> list[dict]:
    return [
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
        },
    ]


def main() -> None:
    started = perf_counter()
    OUT.mkdir(exist_ok=True)

    schema = mo_dataport.extract(str(TEMPLATE))
    rows = pl.scan_parquet(DATA).select(["Employee", "Amount"])

    # shift="both": dataframe expands right and down inside repeat blocks
    bundle = mo_dataport.compile(
        schema,
        {"Shift Demo": {"reports": _report_payload(rows)}},
        dataframe_shift="both",
    )
    xlsx_out = mo_dataport.export(
        bundle,
        str(OUT / "output.xlsx"),
        export_mode="streaming",
        streaming_chunk_rows=1,
    )
    mo_dataport.export(
        bundle,
        str(OUT / "output.pdf"),
        format="pdf",
        streaming_chunk_rows=1,
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE}")
    print(f"Parquet:  {DATA}")
    print(f"XLSX:     {xlsx_out[0]}")
    print(f"PDF:      {OUT / 'output.pdf'}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
