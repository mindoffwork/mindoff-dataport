from __future__ import annotations

import datetime as dt
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


def main() -> None:
    started = perf_counter()
    OUT.mkdir(exist_ok=True)

    schema = mo_dataport.extract(str(TEMPLATE))
    sales = pl.scan_parquet(DATA).select(
        ["Product", "Region", "Units", "Revenue", "Growth", "Margin"]
    )
    bundle = mo_dataport.compile(
        schema,
        {
            "Q1 Review": {
                "report_title": "Quarterly Business Review — Q1 2026",
                "department": "Finance & Operations Division",
                "generated_on": dt.date(2026, 4, 30),
                "kpi_revenue": "€ 2.955M",
                "kpi_growth": "+14.6%",
                "kpi_margin": "30.0%",
                "sales": sales,
            }
        },
    )

    mo_dataport.export(
        bundle,
        str(OUT / "output_openpyxl.xlsx"),
        format="xlsx",
        export_mode="streaming",
        streaming_engine="openpyxl",
        column_width_mode="fixed",
        row_height_mode="fixed",
    )
    mo_dataport.export(
        bundle,
        str(OUT / "output_xlsxwriter.xlsx"),
        format="xlsx",
        export_mode="streaming",
        streaming_engine="xlsxwriter",
        column_width_mode="fixed",
        row_height_mode="fixed",
    )
    mo_dataport.export(
        bundle,
        str(OUT / "output.pdf"),
        format="pdf",
        column_width_mode="fixed",
        row_height_mode="fixed",
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE}")
    print(f"Parquet:  {DATA}")
    print(f"XLSX (openpyxl):   {OUT / 'output_openpyxl.xlsx'}")
    print(f"XLSX (xlsxwriter): {OUT / 'output_xlsxwriter.xlsx'}")
    print(f"PDF:               {OUT / 'output.pdf'}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
