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


def _report_payload(rows: pl.LazyFrame) -> list[dict]:
    customers = rows.select("customer").unique().sort("customer").collect()
    reports = []
    for customer in customers["customer"].to_list():
        customer_rows = rows.filter(pl.col("customer") == customer)
        meta = customer_rows.select(["customer", "region"]).limit(1).collect().row(0, named=True)
        reports.append(
            {
                "customer_name": meta["customer"],
                "region": meta["region"],
                "line_items": customer_rows.select(["sku", "item", "qty", "unit_price"]),
            }
        )
    return reports


def main() -> None:
    started = perf_counter()
    OUT.mkdir(exist_ok=True)

    schema = mo_dataport.extract(str(TEMPLATE))
    rows = pl.scan_parquet(DATA)
    bundle = mo_dataport.compile(
        schema,
        {"Repeat Report": {"reports": _report_payload(rows)}},
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
