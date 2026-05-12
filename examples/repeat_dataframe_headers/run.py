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


def _repeat_payload(rows: pl.LazyFrame) -> list[dict]:
    customers = rows.select("customer").unique().sort("customer").collect()["customer"].to_list()
    payload: list[dict] = []
    for customer in customers:
        customer_rows = rows.filter(pl.col("customer") == customer).select(["sku", "item", "qty"])
        payload.append({"customer_name": customer, "line_items": customer_rows})
    return payload


def main() -> None:
    started = perf_counter()
    OUT.mkdir(exist_ok=True)

    schema = mo_dataport.extract(str(TEMPLATE))
    rows = pl.scan_parquet(DATA)
    summary = rows.select(["customer", "sku", "qty"])

    bundle = mo_dataport.compile(
        schema,
        {
            "Combined Header Demo": {
                "summary_headers": summary,
                "summary_rows": summary,
            },
            "Repeat Header Demo": {
                "reports": _repeat_payload(rows),
            },
        },
        dataframe_shift="vertical",
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
        repeat_dataframe_headers=True,
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE}")
    print(f"Parquet:  {DATA}")
    print(f"XLSX:     {xlsx_out[0]}")
    print(f"PDF:      {OUT / 'output.pdf'} (repeat_dataframe_headers=True)")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
