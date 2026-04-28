from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mode as mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"


def _export(mode_name: str, **options) -> None:
    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    bundle = mo_dataport.compile(
        schema,
        {
            "Sizing Demo": {
                "short_text": "Tiny",
                "long_text": "A much longer value that hug sizing can use",
            }
        },
    )
    output = HERE / f"output_{mode_name}.xlsx"
    mo_dataport.export(bundle, str(output), **options)
    print(f"{mode_name}: {output}")


def main() -> None:
    started = perf_counter()

    _export("fixed", column_width_mode="fixed", row_height_mode="fixed")
    _export(
        "even",
        column_width_mode="even",
        row_height_mode="even",
        default_column_width=24,
        default_row_height=26,
    )
    _export("hug", column_width_mode="hug", row_height_mode="hug")

    elapsed = perf_counter() - started
    print(f"Template: {TEMPLATE_XLSX}")
    print(f"Elapsed:  {elapsed:.2f}s")


if __name__ == "__main__":
    main()
