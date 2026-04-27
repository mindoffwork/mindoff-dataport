from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path
from time import perf_counter

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from mindoff_dataport import mode as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
DATA_PARQUET = HERE / "data.parquet"
OUTPUT_DIR = HERE / "output"
OUTPUT_PDF = OUTPUT_DIR / "styled_parquet_output.pdf"
BUNDLE_DIR = OUTPUT_DIR / "report_bundle_pdf"


def main() -> None:
    started = perf_counter()
    OUTPUT_DIR.mkdir(exist_ok=True)

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET)
    sheet_name = schema["sheets"][0]["name"]

    bundle = mo_dataport.compile(
        schema,
        {
            sheet_name: {
                "report_title": "Directory-backed parquet PDF export",
                "prepared_for": "Mindoff QA",
                "run_date": dt.date.today(),
                "status": "Rendering parquet batches to PDF",
                "line_item_headers": rows,
                "line_items": rows,
            }
        },
        bundle_path=str(BUNDLE_DIR),
    )
    mo_dataport.export(
        bundle,
        str(OUTPUT_PDF),
        format="pdf",
        streaming_chunk_rows=2,
        auto_delete_bundle=True,
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_PDF}")
    print(f"Bundle removed: {not BUNDLE_DIR.exists()}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
