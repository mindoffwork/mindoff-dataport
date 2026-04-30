from __future__ import annotations

import pprint
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from mindoff_dataport import mo_dataport

HERE = Path(__file__).resolve().parent
TEMPLATE_XLSX = HERE / "template.xlsx"


def main() -> None:
    schema = mo_dataport.extract(str(TEMPLATE_XLSX))
    pprint.pp(mo_dataport.inputs(schema))


if __name__ == "__main__":
    main()
