from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mode as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"


def main() -> None:
    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    bundle = mo_dataport.compile(schema, {"Image Reserved": {"name": "Acme"}})
    try:
        mo_dataport.export(bundle, str(HERE / "output.png"), format="image")
    except NotImplementedError as exc:
        print(f"image export is reserved: {exc}")


if __name__ == "__main__":
    main()
