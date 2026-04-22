"""
Actual demo flow:
1) Start from a real .xlsx template with typed placeholders.
2) Extract schema directly from the template.
3) Discover required variables + types from the extracted schema.
4) Build a typed dictionary of runtime values (including Polars DataFrames).
5) Render and rebuild customized .xlsx outputs without writing intermediate JSON files.

Run from project root:
    python examples/demo.py
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path
from time import perf_counter
from typing import Any

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import polars as pl

from mindoff_data_export import build_template_with_data, extract_template, get_template_inputs

EXAMPLES_DIR = Path(__file__).parent
INPUT_TEMPLATE = EXAMPLES_DIR / "input" / "customer_statement_template.xlsx"
OUT_DIR = EXAMPLES_DIR / "output" / "real_usage_case"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def _default_for_type(type_name: str) -> Any:
    if type_name == "string":
        return ""
    if type_name == "number":
        return 0
    if type_name == "int":
        return 0
    if type_name == "float":
        return 0.0
    if type_name == "boolean":
        return False
    if type_name == "date":
        return dt.date.today()
    if type_name in ("dataframe-headers", "dataframe-content"):
        return pl.DataFrame()
    raise ValueError(f"Unsupported variable type: {type_name}")


def build_typed_payload(
    variables_by_type: dict[str, str],
    overrides: dict[str, Any],
) -> dict[str, Any]:
    payload = {key: _default_for_type(var_type) for key, var_type in variables_by_type.items()}
    payload.update(overrides)
    return payload


def get_polars_line_items() -> pl.DataFrame:
    return pl.DataFrame(
        {
            "Item": ["Monthly subscription", "Setup fee", "Support package"],
            "Qty": [12, 1, 1],
            "Unit Price": [1299.50, 4999.00, 2499.00],
        }
    ).with_columns((pl.col("Qty") * pl.col("Unit Price")).alias("Total"))


def _build_to_path(
    schema: dict[str, Any],
    data: dict[str, Any],
    path: Path,
    **kwargs: Any,
) -> Path:
    try:
        build_template_with_data(schema, data, str(path), **kwargs)
        return path
    except PermissionError:
        # Common on Windows when target workbook is open in Excel.
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = path.with_name(f"{path.stem}_{stamp}{path.suffix}")
        build_template_with_data(schema, data, str(fallback), **kwargs)
        return fallback


def main() -> None:
    started = perf_counter()
    if not INPUT_TEMPLATE.exists():
        raise FileNotFoundError(f"Template not found: {INPUT_TEMPLATE}")

    print(f"[1/5] Extracting schema from template: {INPUT_TEMPLATE}")
    schema = extract_template(str(INPUT_TEMPLATE))
    print("[2/5] Discovering typed placeholders")
    variables_by_type = get_template_inputs(schema)

    # Core runtime data: scalar + Polars DataFrame
    print("[3/5] Building typed runtime payload (includes Polars DataFrame)")
    data = build_typed_payload(
        variables_by_type,
        overrides={
            "customer_name": "Acme Industries Pvt Ltd",
            "invoice_number": 1024,
            "invoice_amount": 24999.75,
            "is_paid": True,
            "invoice_date": dt.date(2026, 4, 22),
            "line_items": get_polars_line_items(),
        },
    )

    # Customization 1: Use template sizing exactly as-is.
    print("[4/5] Generating workbook variants")
    out_fixed = _build_to_path(
        schema,
        data,
        OUT_DIR / "customer_statement_fixed.xlsx",
    )

    # Customization 2: Auto-fit columns to content and use even row height.
    out_hug = _build_to_path(
        schema,
        data,
        OUT_DIR / "customer_statement_hug_columns.xlsx",
        column_width_mode="hug",
        row_height_mode="even",
        default_row_height=20.0,
    )

    # Customization 3: DataFrame from LazyFrame + even columns + hug rows.
    lazy_data = dict(data)
    lazy_data["line_items"] = get_polars_line_items().lazy()
    out_lazy = _build_to_path(
        schema,
        lazy_data,
        OUT_DIR / "customer_statement_lazyframe.xlsx",
        column_width_mode="even",
        default_column_width=18.0,
        row_height_mode="hug",
    )

    elapsed = perf_counter() - started
    print("[5/5] Done")
    print(f"Template: {INPUT_TEMPLATE}")
    print(f"Variables discovered: {variables_by_type}")
    print(f"Generated: {out_fixed}")
    print(f"Generated: {out_hug}")
    print(f"Generated: {out_lazy}")
    print(f"Elapsed: {elapsed:.2f}s")
    print("")
    print("Customizations demonstrated:")
    print("- scalar placeholders + typed validation")
    print("- Polars DataFrame/LazyFrame placeholder expansion")
    print("- Polars LazyFrame support")
    print("- column_width_mode: fixed | hug | even")
    print("- row_height_mode: fixed | even | hug")
    print("- default_column_width / default_row_height overrides")


if __name__ == "__main__":
    main()
