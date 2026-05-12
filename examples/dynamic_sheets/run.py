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
    OUT.mkdir(exist_ok=True)

    schema = mo_dataport.extract(str(TEMPLATE))
    rows = pl.scan_parquet(DATA)
    bundle = mo_dataport.compile(schema, {"region_sheet": _region_payload(rows)})

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
