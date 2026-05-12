"""Guard script for fidelity XLSX speed against direct openpyxl baseline."""

from __future__ import annotations

import csv
import importlib.util
import sys
import tempfile
from pathlib import Path


HERE = Path(__file__).resolve().parent
SCENARIO = "xlsx-fidelity-vs-openpyxl"
ROW_LABELS = ("1K", "10K")


def _read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        raise FileNotFoundError(f"Expected benchmark output file was not found: {path}")
    with path.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def _fidelity_speed_categories(rows: list[dict[str, str]]) -> dict[str, str]:
    categories: dict[str, str] = {}
    for row in rows:
        if row.get("scenario") != SCENARIO:
            continue
        row_label = row.get("rows", "")
        if row_label in ROW_LABELS:
            raw = row.get("speed_category", "").strip()
            categories[row_label] = raw.split(" ", 1)[0] if raw else ""
    return categories


def _elapsed_map(rows: list[dict[str, str]]) -> dict[tuple[str, str], float]:
    result: dict[tuple[str, str], float] = {}
    for row in rows:
        method = row.get("method", "")
        row_count = row.get("rows", "")
        elapsed = row.get("elapsed_s", "")
        if not method or not row_count or not elapsed:
            continue
        try:
            result[(method, row_count)] = float(elapsed)
        except ValueError:
            continue
    return result


def _assert_speed_win(comparison_rows: list[dict[str, str]], result_rows: list[dict[str, str]]) -> None:
    categories = _fidelity_speed_categories(comparison_rows)
    elapsed_by_key = _elapsed_map(result_rows)
    failures: list[str] = []

    for row_label, row_count in (("1K", "1000"), ("10K", "10000")):
        category = categories.get(row_label, "")
        mindoff_elapsed = elapsed_by_key.get(("mindoff fidelity (XLSX)", row_count))
        baseline_elapsed = elapsed_by_key.get(("openpyxl - manual styling", row_count))
        if category != "winner" or mindoff_elapsed is None or baseline_elapsed is None:
            failures.append(
                f"{row_label}: category={category or 'missing'}, "
                f"mindoff_elapsed={mindoff_elapsed}, baseline_elapsed={baseline_elapsed}"
            )

    if failures:
        joined = "\n".join(failures)
        raise AssertionError(
            "Fidelity speed guard failed. Expected 'winner' for 1K and 10K.\n"
            f"{joined}"
        )


def _run_quick_benchmark_in_temp() -> tuple[Path, Path]:
    root = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(root / "src"))
    import mindoff_dataport as mo_dataport
    spec = importlib.util.spec_from_file_location("benchmark_run", HERE / "run.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load benchmark runner module.")
    bench = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = bench
    spec.loader.exec_module(bench)

    temp_root = Path(tempfile.mkdtemp(prefix="mindoff_benchmark_guard_"))
    bench.OUTPUT_DIR = temp_root
    bench.ARTIFACT_DIR = temp_root / "files"
    bench.ACTIVE_XLSX_ROW_COUNTS = [1_000, 10_000]
    bench.ACTIVE_PDF_ROW_COUNTS = [1_000, 10_000]
    bench.ACTIVE_RUNS_PER_POINT = 1
    bench.ACTIVE_RUN_TIMEOUT_S = 90

    schema = mo_dataport.extract(str(bench.TEMPLATE_PATH))
    results = bench.run_benchmarks(schema)
    results_csv = bench.save_csv(results)
    comparison_csv = bench.save_comparison_csv(results)
    return results_csv, comparison_csv


def main() -> None:
    results_csv, comparison_csv = _run_quick_benchmark_in_temp()
    comparison_rows = _read_csv(comparison_csv)
    result_rows = _read_csv(results_csv)
    _assert_speed_win(comparison_rows, result_rows)
    print("Fidelity speed guard passed: winner at 1K and 10K.")


if __name__ == "__main__":
    main()
