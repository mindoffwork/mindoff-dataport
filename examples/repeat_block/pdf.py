from __future__ import annotations

import shutil
import sys
import tempfile
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

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET)

    bundle = mo_dataport.compile(
        schema,
        {"Repeat Report": {"reports": _report_payload(rows)}},
    )
    with tempfile.TemporaryDirectory(prefix="mindoff_repeat_pdf_") as tmp_dir:
        tmp_pdf = Path(tmp_dir) / "output.pdf"
        mo_dataport.export(
            bundle,
            str(tmp_pdf),
            format="pdf",
            streaming_chunk_rows=2,
        )
        shutil.copyfile(tmp_pdf, OUTPUT_PDF)

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_PDF}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
