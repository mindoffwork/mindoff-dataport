from __future__ import annotations

import shutil
import sys
from pathlib import Path
from time import perf_counter

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mode as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
BUNDLE_PATH = HERE / "report_bundle"
OUTPUT_PDF = HERE / "output.pdf"


def main() -> None:
    started = perf_counter()
    if BUNDLE_PATH.exists():
        shutil.rmtree(BUNDLE_PATH)

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    mo_dataport.compile(
        schema,
        {"Bundle Demo": {"name": "Acme Industries", "status": "Compiled to disk"}},
        bundle_path=str(BUNDLE_PATH),
    )
    mo_dataport.export(str(BUNDLE_PATH), str(OUTPUT_PDF), format="pdf")

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Bundle:   {BUNDLE_PATH}")
    print(f"Output:   {OUTPUT_PDF}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
