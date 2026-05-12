from __future__ import annotations

import shutil
import sys
from pathlib import Path
from time import perf_counter

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

import mindoff_dataport as mo_dataport

HERE = Path(__file__).resolve().parent
OUT = HERE / "output"
BUNDLE = OUT / "bundle"
TEMPLATE = HERE / "template.xlsx"


def main() -> None:
    started = perf_counter()
    OUT.mkdir(exist_ok=True)
    if BUNDLE.exists():
        shutil.rmtree(BUNDLE)

    schema = mo_dataport.extract(str(TEMPLATE))
    mo_dataport.compile(
        schema,
        {"Bundle Demo": {"name": "Acme Industries", "status": "Compiled to disk"}},
        bundle_path=str(BUNDLE),
    )
    mo_dataport.export(str(BUNDLE), str(OUT / "output.xlsx"))
    mo_dataport.export(str(BUNDLE), str(OUT / "output.pdf"), format="pdf")

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE}")
    print(f"Bundle:   {BUNDLE}")
    print(f"XLSX:     {OUT / 'output.xlsx'}")
    print(f"PDF:      {OUT / 'output.pdf'}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
