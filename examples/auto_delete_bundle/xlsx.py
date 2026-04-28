from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mode as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"
OUTPUT_XLSX = HERE / "output.xlsx"


def main() -> None:
    started = perf_counter()

    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    bundle = mo_dataport.compile(
        schema,
        {"Cleanup Demo": {"name": "Acme Industries", "status": "Deleted after export"}},
    )
    bundle_path = Path(bundle.path)
    mo_dataport.export(bundle, str(OUTPUT_XLSX), auto_delete_bundle=True)

    elapsed = perf_counter() - started
    print(f"Template:      {TEMPLATE_XLSX}")
    print(f"Output:        {OUTPUT_XLSX}")
    print(f"Bundle exists: {bundle_path.exists()}")
    print(f"Elapsed:       {elapsed:.2f}s")


if __name__ == "__main__":
    main()
