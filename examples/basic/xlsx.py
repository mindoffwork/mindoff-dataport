from __future__ import annotations

import datetime as dt
import sys
import shutil
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


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET).select(
        ["product", "region", "units", "revenue"]
    )

    bundle = mo_dataport.compile(
        schema,
        {
            "Sales Summary": {
                "report_title": "Basic Sales Summary",
                "generated_on": dt.date(2026, 4, 28),
                "sales_rows": rows,
            }
        },
    )
    with tempfile.TemporaryDirectory(prefix="mindoff_basic_xlsx_") as tmp_dir:
        outputs = mo_dataport.export(
            bundle,
            str(Path(tmp_dir) / "output.xlsx"),
            export_mode="streaming",
            streaming_chunk_rows=2,
            column_width_mode="fixed",
            row_height_mode="fixed",
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
