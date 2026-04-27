from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path
from time import perf_counter

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from mindoff_dataport import mode

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
DATA_PARQUET = HERE / "data.parquet"
OUTPUT_XLSX = HERE / "styled_parquet_output.xlsx"
BUNDLE_DIR = HERE / "report_bundle"


def main() -> None:
    started = perf_counter()

    schema = mode.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET)
    columns = rows.collect_schema().names()
    sheet_name = schema["sheets"][0]["name"]

    bundle = mode.compile(
        schema,
        {
            sheet_name: {
                "report_title": "Directory-backed parquet export",
                "prepared_for": "Mindoff QA",
                "run_date": dt.date.today(),
                "status": "Streaming parquet batches via pyarrow",
                "line_item_headers": columns,
                "line_items": mode.parquet_source(str(DATA_PARQUET), columns=columns),
            }
        },
        bundle_path=str(BUNDLE_DIR),
    )
    outputs = mode.export(
        bundle,
        str(OUTPUT_XLSX),
        export_mode="streaming",
        streaming_chunk_rows=2,
        auto_delete_bundle=True,
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {outputs[0]}")
    print(f"Bundle removed: {not BUNDLE_DIR.exists()}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
