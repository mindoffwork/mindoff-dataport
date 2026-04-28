from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path
from time import perf_counter

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mode as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
DATA_PARQUET = HERE / "data.parquet"
OUTPUT_XLSX = HERE / "output.xlsx"


def _summary_payload(rows: pl.LazyFrame) -> list[dict]:
    regions = rows.select("region").unique().sort("region").collect()
    summaries = []
    for region in regions["region"].to_list():
        region_rows = rows.filter(pl.col("region") == region)
        totals = region_rows.select(["product", "units", "revenue"])
        summaries.append({"region": region, "totals": totals})
    return summaries


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET)

    bundle = mo_dataport.compile(
        schema,
        {
            "Multi Repeat Report": {
                "customers": [
                    {"customer_name": "Acme Industries", "tier": "Enterprise"},
                    {"customer_name": "Globex Retail", "tier": "Growth"},
                    {"customer_name": "Initech Labs", "tier": "Pilot"},
                ],
                "summaries": _summary_payload(rows),
            }
        },
    )
    with tempfile.TemporaryDirectory(prefix="mindoff_multi_repeat_xlsx_") as tmp_dir:
        outputs = mo_dataport.export(
            bundle,
            str(Path(tmp_dir) / "output.xlsx"),
            export_mode="streaming",
            streaming_chunk_rows=2,
        )
        shutil.copyfile(outputs[0], OUTPUT_XLSX)

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_XLSX}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
