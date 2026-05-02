"""Style showcase — PDF export.

Demonstrates every supported style field (font, fill, alignment, borders)
via a themed quarterly report template.

Prerequisites:
    python examples/style_showcase/create_template.py
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path
from time import perf_counter

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

import mindoff_dataport as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
DATA_PARQUET = HERE / "data.parquet"
OUTPUT_PDF = HERE / "output.pdf"


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    sales = pl.scan_parquet(DATA_PARQUET).select(
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
        str(OUTPUT_PDF),
        format="pdf",
        column_width_mode="fixed",
        row_height_mode="fixed",
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Parquet:  {DATA_PARQUET}")
    print(f"Output:   {OUTPUT_PDF}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
