from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path
from time import perf_counter
from zipfile import ZipFile

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
DATA_PARQUET = HERE / "data.parquet"
OUTPUT_XLSX = HERE / "output.xlsx"
OUTPUT_ZIP = HERE / "output.zip"


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    rows = pl.scan_parquet(DATA_PARQUET)
    bundle = mo_dataport.compile(schema, {"Split Demo": {"rows": rows}})
    with tempfile.TemporaryDirectory(prefix="mindoff_split_workbooks_") as tmp_dir:
        outputs = mo_dataport.export(
            bundle,
            str(Path(tmp_dir) / "output.xlsx"),
            export_mode="streaming",
            streaming_chunk_rows=3,
            max_rows_per_workbook=4,
        )
        shutil.copyfile(outputs[0], OUTPUT_ZIP)

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_ZIP}")
    with ZipFile(OUTPUT_ZIP) as zip_file:
        print(f"Parts:    {', '.join(zip_file.namelist())}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
