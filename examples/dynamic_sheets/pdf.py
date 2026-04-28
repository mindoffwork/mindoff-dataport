from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mode as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
DATA_PARQUET = HERE / "data.parquet"
OUTPUT_PDF = HERE / "output.pdf"


def _region_payload(rows: pl.LazyFrame) -> dict[str, dict]:
    regions = rows.select("region").unique().sort("region").collect()
    payload = {}
    for region in regions["region"].to_list():
        region_rows = rows.filter(pl.col("region") == region)
        owner = region_rows.select("owner").limit(1).collect().item()
        payload[f"{region} Sheet"] = {
            "region_name": region,
            "owner": owner,
            "sales_rows": region_rows.select(["product", "units", "revenue"]),
        }
    return payload


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET)

    bundle = mo_dataport.compile(
        schema,
        {"region_sheet": _region_payload(rows)},
    )
    mo_dataport.export(
        bundle,
        str(OUTPUT_PDF),
        format="pdf",
        streaming_chunk_rows=2,
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_PDF}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
