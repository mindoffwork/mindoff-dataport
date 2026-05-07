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
OUTPUT_XLSX = HERE / "output.xlsx"


def _ensure_assets() -> None:
    if TEMPLATE_XLSX.exists() and DATA_PARQUET.exists():
        return
    from create_template import main as create_assets

    create_assets()


def _repeat_payload(rows: pl.LazyFrame) -> list[dict]:
    customers = rows.select("customer").unique().sort("customer").collect()["customer"].to_list()
    payload: list[dict] = []
    for customer in customers:
        customer_rows = rows.filter(pl.col("customer") == customer).select(["sku", "item", "qty"])
        payload.append({"customer_name": customer, "line_items": customer_rows})
    return payload


def main() -> None:
    started = perf_counter()
    _ensure_assets()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET)
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
    mo_dataport.export(bundle, str(OUTPUT_XLSX), export_mode="streaming", streaming_chunk_rows=1)

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_XLSX}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
