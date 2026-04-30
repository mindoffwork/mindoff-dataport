from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"


def _show_error(label: str, func) -> None:
    try:
        func()
    except Exception as exc:
        print(f"{label}: {type(exc).__name__}: {exc}")


def main() -> None:
    schema = mo_dataport.extract(str(TEMPLATE_XLSX))

    _show_error(
        "missing sheet",
        lambda: mo_dataport.compile(schema, {}),
    )
    _show_error(
        "missing field",
        lambda: mo_dataport.compile(schema, {"Validation": {}}),
    )
    _show_error(
        "wrong repeat payload",
        lambda: mo_dataport.compile(
            schema,
            {"Validation": {"reports": {"name": "Acme"}}},
        ),
    )


if __name__ == "__main__":
    main()
