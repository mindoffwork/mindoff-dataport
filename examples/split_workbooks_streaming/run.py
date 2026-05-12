from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter
from zipfile import ZipFile

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

import mindoff_dataport as mo_dataport

HERE = Path(__file__).resolve().parent
OUT = HERE / "output"
TEMPLATE = HERE / "template.xlsx"
DATA = HERE / "data.parquet"


def main() -> None:
    started = perf_counter()
    OUT.mkdir(exist_ok=True)

    schema = mo_dataport.extract(str(TEMPLATE))
    rows = pl.scan_parquet(DATA)
    bundle = mo_dataport.compile(schema, {"Split Demo": {"rows": rows}})
    outputs = mo_dataport.export(
        bundle,
        str(OUT / "output.xlsx"),
        export_mode="streaming",
        streaming_chunk_rows=3,
        max_rows_per_workbook=4,
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE}")
    print(f"Parquet:  {DATA}")
    print(f"Output:   {outputs[0]}")
    with ZipFile(outputs[0]) as zip_file:
        print(f"Parts:    {', '.join(zip_file.namelist())}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
