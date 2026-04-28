from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter

import reportlab

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mode as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
OUTPUT_PDF = HERE / "output.pdf"
VERA_FONT = Path(reportlab.__file__).parent / "fonts" / "Vera.ttf"


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    bundle = mo_dataport.compile(schema, {"Custom Font": {"title": "PDF Custom Font"}})
    mo_dataport.export(
        bundle,
        str(OUTPUT_PDF),
        format="pdf",
        fonts={"Vera": {"regular": str(VERA_FONT)}},
    )

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Font:     {VERA_FONT}")
    print(f"Output:   {OUTPUT_PDF}")
    print(f"Bundle:   {bundle.path}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
