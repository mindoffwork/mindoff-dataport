"""Public-facing benchmark: runtime compile+export only.

This benchmark intentionally excludes template extraction/schema creation from
timed measurements. The template is extracted once, then each benchmark run
measures compile + export only.

Produces:
    output/results.csv
    output/files/
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import statistics
import shutil
import sys
import threading
import tracemalloc
from dataclasses import dataclass
from datetime import date, timedelta
from functools import partial
from pathlib import Path
from time import perf_counter, perf_counter_ns

import polars as pl


try:
    from openpyxl import Workbook as _OXLWorkbook
    from openpyxl.styles import Alignment as _Align
    from openpyxl.styles import Border as _Border
    from openpyxl.styles import Color as _Color
    from openpyxl.styles import Font as _Font
    from openpyxl.styles import PatternFill as _Fill
    from openpyxl.styles import Side as _Side

    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False

try:
    import xlsxwriter as _xlsxwriter

    _HAS_XLSXWRITER = True
except ImportError:
    _HAS_XLSXWRITER = False

try:
    from reportlab.lib import colors as _rl_colors
    from reportlab.lib.pagesizes import A4 as _RL_A4
    from reportlab.platypus import SimpleDocTemplate as _RLDoc
    from reportlab.platypus import Table as _RLTable
    from reportlab.platypus import TableStyle as _RLStyle

    _HAS_REPORTLAB = True
except ImportError:
    _HAS_REPORTLAB = False

try:
    import psutil as _psutil

    _PROC = _psutil.Process()
    _HAS_PSUTIL = True
except ImportError:
    _HAS_PSUTIL = False

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
import mindoff_dataport as mo_dataport

# Thread-local store: mindoff bench functions write compile_s here so _timed_runs can read it.
_BENCH_CTX = threading.local()


# Â§1. Constants & Configuration
HERE = Path(__file__).resolve().parent
TEMPLATE_PATH = HERE / "benchmark_template.xlsx"
OUTPUT_DIR = HERE / "output"
ARTIFACT_DIR = OUTPUT_DIR / "files"

XLSX_ROW_COUNTS = [1_000, 10_000, 100_000, 500_000]
PDF_ROW_COUNTS = [1_000, 5_000, 10_000]

RUN_TIMEOUT_S = 120
RUNS_PER_POINT = 5

_COLS = ["id", "name", "category", "value", "date"]

_HUMAN_ROW_LABELS: dict[int, str] = {
    1_000: "1K",
    5_000: "5K",
    10_000: "10K",
    100_000: "100K",
    500_000: "500K",
    1_000_000: "1M",
}

# Runtime-configurable benchmark parameters (set by CLI in main()).
ACTIVE_XLSX_ROW_COUNTS = XLSX_ROW_COUNTS
ACTIVE_PDF_ROW_COUNTS = PDF_ROW_COUNTS
ACTIVE_RUN_TIMEOUT_S = RUN_TIMEOUT_S
ACTIVE_RUNS_PER_POINT = RUNS_PER_POINT


# Â§2. Data Classes
@dataclass
class BenchResult:
    method: str
    fmt: str
    rows: int
    elapsed_s: float
    peak_mb: float
    file_mb: float
    status: str = "completed"
    runs: int = ACTIVE_RUNS_PER_POINT
    compile_s: float = 0.0
    elapsed_stdev_s: float = float("nan")
    peak_stdev_mb: float = float("nan")
    output_file: str = ""


# Â§3. Private Helper Functions
def _is_valid_number(value: float) -> bool:
    return math.isfinite(value) and not math.isnan(value)


def _format_number(value: float, places: int = 1) -> str:
    if not _is_valid_number(value):
        return "-"
    return f"{value:,.{places}f}"


def _format_rows(n: int) -> str:
    return _HUMAN_ROW_LABELS.get(n, f"{n:,}")


def _make_parquet(n_rows: int) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUTPUT_DIR / f"bench_data_{n_rows}.parquet"
    if path.exists():
        return path
    cats = ["Alpha", "Beta", "Gamma", "Delta"]
    base = date(2024, 1, 1)
    df = pl.DataFrame(
        {
            "id": pl.Series(list(range(1, n_rows + 1)), dtype=pl.Int32),
            "name": [f"Record {i:07d}" for i in range(1, n_rows + 1)],
            "category": [cats[i % 4] for i in range(n_rows)],
            "value": [round(100.0 + (i * 0.37 % 900), 2) for i in range(n_rows)],
            "date": [
                (base + timedelta(days=i % 365)).isoformat() for i in range(n_rows)
            ],
        }
    )
    df.write_parquet(str(path))
    print(f"  [data] generated {n_rows:,} rows -> {path.name}")
    return path


def _measure(fn) -> tuple[float, float]:
    if _HAS_PSUTIL:
        peak_rss = [_PROC.memory_info().rss]
        stop_flag = threading.Event()

        def _poll() -> None:
            while not stop_flag.is_set():
                try:
                    peak_rss[0] = max(peak_rss[0], _PROC.memory_info().rss)
                except Exception:
                    pass
                stop_flag.wait(0.05)

        baseline_rss = _PROC.memory_info().rss
        poller = threading.Thread(target=_poll, daemon=True)
        poller.start()
        t0 = perf_counter()
        fn()
        elapsed = perf_counter() - t0
        stop_flag.set()
        poller.join(timeout=1)
        peak_mb = (peak_rss[0] - baseline_rss) / 1024 / 1024
        if peak_mb < 0:
            peak_mb = peak_rss[0] / 1024 / 1024
        return elapsed, peak_mb

    if tracemalloc.is_tracing():
        tracemalloc.stop()
    tracemalloc.start()
    t0 = perf_counter()
    fn()
    elapsed = perf_counter() - t0
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return elapsed, peak / 1024 / 1024


def _export_file_mb(tmp_dir: Path, ext: str) -> float:
    return sum(f.stat().st_size for f in tmp_dir.glob(f"*.{ext}")) / 1024 / 1024


def _safe_artifact_stem(method: str, rows: int) -> str:
    safe_method = re.sub(r"[^A-Za-z0-9]+", "_", method).strip("_").lower()
    return f"{rows:09d}_{safe_method or 'method'}"


def _artifact_path(method: str, fmt: str, rows: int, ext: str) -> Path:
    return ARTIFACT_DIR / fmt / f"{_safe_artifact_stem(method, rows)}.{ext}"


def _copy_export_outputs(tmp_dir: Path, ext: str, keep_output_path: Path) -> list[str]:
    exported = sorted(tmp_dir.glob(f"*.{ext}"))
    if not exported:
        return []
    keep_output_path.parent.mkdir(parents=True, exist_ok=True)
    if len(exported) == 1:
        shutil.copy2(exported[0], keep_output_path)
        return [str(keep_output_path)]

    saved: list[str] = []
    for index, source in enumerate(exported, start=1):
        target = keep_output_path.with_name(
            f"{keep_output_path.stem}.part{index:03d}{keep_output_path.suffix}"
        )
        shutil.copy2(source, target)
        saved.append(str(target))
    return saved


def _run_once(fn, ext: str, keep_output_path: Path | None = None) -> dict[str, object]:
    """Execute fn once in a timeout-guarded thread; return state dict."""
    state: dict[str, object] = {}

    def _target() -> None:
        try:
            scratch_root = OUTPUT_DIR / "_scratch"
            scratch_root.mkdir(parents=True, exist_ok=True)
            tmp_dir = scratch_root / f"run_{perf_counter_ns()}"
            tmp_dir.mkdir(parents=True, exist_ok=True)
            out = tmp_dir / f"output.{ext}"
            _BENCH_CTX.compile_s = 0.0
            elapsed, peak_mb = _measure(partial(fn, out))
            file_mb = _export_file_mb(tmp_dir, ext)
            compile_s = getattr(_BENCH_CTX, "compile_s", 0.0)
            saved_outputs = (
                _copy_export_outputs(tmp_dir, ext, keep_output_path)
                if keep_output_path is not None
                else []
            )
            state.update(
                elapsed=elapsed,
                peak_mb=peak_mb,
                file_mb=file_mb,
                compile_s=compile_s,
                saved_outputs=saved_outputs,
            )
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception as exc:
            state["error"] = str(exc)

    thread = threading.Thread(target=_target, daemon=True)
    thread.start()
    thread.join(timeout=ACTIVE_RUN_TIMEOUT_S)
    if thread.is_alive():
        state["timeout"] = True
    return state


def _timed_runs(
    fn, ext: str, keep_output_path: Path | None = None
) -> tuple[float, float, float, float, float, float, str, str]:
    elapsed_vals: list[float] = []
    peak_vals: list[float] = []
    file_vals: list[float] = []
    compile_vals: list[float] = []

    # Discard one warm-up pass to avoid cold Python import / OS file-cache effects.
    if ACTIVE_RUNS_PER_POINT >= 3:
        _run_once(fn, ext)

    for run_idx in range(ACTIVE_RUNS_PER_POINT):
        state = _run_once(
            fn,
            ext,
            keep_output_path
            if run_idx == ACTIVE_RUNS_PER_POINT - 1
            else None,
        )
        if state.get("timeout"):
            return (
                float("nan"),
                float("nan"),
                float("nan"),
                float("nan"),
                0.0,
                0.0,
                "timeout",
                f"timed out after {ACTIVE_RUN_TIMEOUT_S}s",
            )
        if "error" in state:
            return (
                float("nan"),
                float("nan"),
                float("nan"),
                float("nan"),
                0.0,
                0.0,
                "failed",
                str(state["error"]),
            )
        elapsed_vals.append(float(state["elapsed"]))
        peak_vals.append(float(state["peak_mb"]))
        file_vals.append(float(state["file_mb"]))
        compile_vals.append(float(state.get("compile_s", 0.0)))

    stdev_e = statistics.stdev(elapsed_vals) if len(elapsed_vals) >= 2 else float("nan")
    stdev_m = statistics.stdev(peak_vals) if len(peak_vals) >= 2 else float("nan")
    return (
        statistics.median(elapsed_vals),
        stdev_e,
        statistics.median(peak_vals),
        stdev_m,
        statistics.median(file_vals),
        statistics.median(compile_vals),
        "completed",
        "",
    )


# Â§4. Benchmark Targets


def _payload(df_path: Path, n_rows: int) -> dict[str, dict[str, object]]:
    return {
        "Benchmark": {
            "report_title": "Benchmark Report",
            "generated_on": date.today(),
            "row_count": n_rows,
            "bench_data": pl.scan_parquet(str(df_path)),
        }
    }


def _bench_mindoff_xlsx_fidelity(schema, df_path: Path, n_rows: int, out: Path) -> None:
    bundle_path = out.parent / "bundle"
    t0 = perf_counter()
    bundle = mo_dataport.compile(
        schema, _payload(df_path, n_rows), bundle_path=str(bundle_path)
    )
    _BENCH_CTX.compile_s = perf_counter() - t0
    mo_dataport.export(
        bundle, str(out), export_mode="fidelity", auto_delete_bundle=False
    )


def _bench_mindoff_xlsx_streaming_openpyxl(
    schema, df_path: Path, n_rows: int, out: Path
) -> None:
    bundle_path = out.parent / "bundle"
    t0 = perf_counter()
    bundle = mo_dataport.compile(
        schema, _payload(df_path, n_rows), bundle_path=str(bundle_path)
    )
    _BENCH_CTX.compile_s = perf_counter() - t0
    mo_dataport.export(
        bundle,
        str(out),
        export_mode="streaming",
        streaming_engine="openpyxl",
        streaming_chunk_rows=10_000,
        auto_delete_bundle=False,
    )


def _bench_mindoff_xlsx_streaming_xlsxwriter(
    schema, df_path: Path, n_rows: int, out: Path
) -> None:
    bundle_path = out.parent / "bundle"
    t0 = perf_counter()
    bundle = mo_dataport.compile(
        schema, _payload(df_path, n_rows), bundle_path=str(bundle_path)
    )
    _BENCH_CTX.compile_s = perf_counter() - t0
    mo_dataport.export(
        bundle,
        str(out),
        export_mode="streaming",
        streaming_engine="xlsxwriter",
        streaming_chunk_rows=10_000,
        auto_delete_bundle=False,
    )


def _bench_openpyxl_direct(df_path: Path, n_rows: int, out: Path) -> None:
    df = pl.read_parquet(str(df_path))
    wb = _OXLWorkbook()
    ws = wb.active
    ws.title = "Benchmark"
    for col, w in zip("ABCDE", [12, 28, 18, 14, 16]):
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "Benchmark Report"
    c.font = _Font(name="Calibri", size=14, bold=True, color="FFFFFFFF")
    c.fill = _Fill(patternType="solid", fgColor=_Color(rgb="FF1E3A5F"))
    c.alignment = _Align(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36
    ws.merge_cells("A2:C2")
    ws["A2"].value = f"Generated: {date.today()}"
    ws["A2"].font = _Font(name="Calibri", size=10, color="FFFFFFFF")
    ws["A2"].fill = _Fill(patternType="solid", fgColor=_Color(rgb="FF2563EB"))
    ws["A2"].alignment = _Align(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("D2:E2")
    ws["D2"].value = f"Rows: {n_rows}"
    ws["D2"].font = _Font(name="Calibri", size=10, color="FFFFFFFF")
    ws["D2"].fill = _Fill(patternType="solid", fgColor=_Color(rgb="FF2563EB"))
    ws["D2"].alignment = _Align(horizontal="right", vertical="center")
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 8
    hdr_font = _Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    hdr_fill = _Fill(patternType="solid", fgColor=_Color(rgb="FF1E3A5F"))
    hdr_align = _Align(horizontal="center", vertical="center")
    hdr_border = _Border(
        left=_Side(border_style="thin", color="FFFFFFFF"),
        right=_Side(border_style="thin", color="FFFFFFFF"),
        top=_Side(border_style="thin", color="FFFFFFFF"),
        bottom=_Side(border_style="thin", color="FFFFFFFF"),
    )
    for col_idx, hdr in enumerate(_COLS, 1):
        cell = ws.cell(row=4, column=col_idx, value=hdr)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = hdr_border
    ws.row_dimensions[4].height = 22
    dat_font = _Font(name="Calibri", size=10, color="FF1E3A5F")
    dat_border = _Border(
        left=_Side(border_style="thin", color="FFCCCCCC"),
        right=_Side(border_style="thin", color="FFCCCCCC"),
        top=_Side(border_style="thin", color="FFCCCCCC"),
        bottom=_Side(border_style="thin", color="FFCCCCCC"),
    )
    dat_align = _Align(horizontal="center", vertical="center")
    for row_idx, row in enumerate(df.iter_rows(), 5):
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = dat_font
            cell.border = dat_border
            cell.alignment = dat_align
    wb.save(str(out))


def _bench_xlsxwriter_direct(df_path: Path, n_rows: int, out: Path) -> None:
    df = pl.read_parquet(str(df_path))
    with _xlsxwriter.Workbook(str(out)) as wb:
        ws = wb.add_worksheet("Benchmark")
        title_fmt = wb.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#1E3A5F",
                "font_size": 14,
                "align": "center",
                "valign": "vcenter",
            }
        )
        sub_l_fmt = wb.add_format(
            {
                "font_color": "#FFFFFF",
                "bg_color": "#2563EB",
                "font_size": 10,
                "align": "left",
                "valign": "vcenter",
                "indent": 1,
            }
        )
        sub_r_fmt = wb.add_format(
            {
                "font_color": "#FFFFFF",
                "bg_color": "#2563EB",
                "font_size": 10,
                "align": "right",
                "valign": "vcenter",
            }
        )
        hdr_fmt = wb.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#1E3A5F",
                "font_size": 10,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": "#FFFFFF",
            }
        )
        dat_fmt = wb.add_format(
            {
                "font_color": "#1E3A5F",
                "font_size": 10,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": "#CCCCCC",
            }
        )
        ws.set_column("A:A", 12)
        ws.set_column("B:B", 28)
        ws.set_column("C:C", 18)
        ws.set_column("D:D", 14)
        ws.set_column("E:E", 16)
        ws.merge_range("A1:E1", "Benchmark Report", title_fmt)
        ws.set_row(0, 36)
        ws.merge_range("A2:C2", f"Generated: {date.today()}", sub_l_fmt)
        ws.merge_range("D2:E2", f"Rows: {n_rows}", sub_r_fmt)
        ws.set_row(1, 22)
        ws.set_row(2, 8)
        for col_idx, hdr in enumerate(_COLS):
            ws.write(3, col_idx, hdr, hdr_fmt)
        ws.set_row(3, 22)
        for row_idx, row in enumerate(df.iter_rows(), 4):
            for col_idx, val in enumerate(row):
                ws.write(row_idx, col_idx, val, dat_fmt)


def _bench_mindoff_pdf(schema, df_path: Path, n_rows: int, out: Path) -> None:
    bundle_path = out.parent / "bundle"
    t0 = perf_counter()
    bundle = mo_dataport.compile(
        schema, _payload(df_path, n_rows), bundle_path=str(bundle_path)
    )
    _BENCH_CTX.compile_s = perf_counter() - t0
    mo_dataport.export(
        bundle,
        str(out),
        format="pdf",
        streaming_chunk_rows=200,
        auto_delete_bundle=False,
    )


def _bench_reportlab_direct(df_path: Path, _n_rows: int, out: Path) -> None:
    df = pl.read_parquet(str(df_path))
    raw_col_widths = [w * 7.0 for w in [12.0, 28.0, 18.0, 14.0, 16.0]]
    available_width = _RL_A4[0] - 72
    scale = min(available_width / sum(raw_col_widths), 1.0)
    col_widths = [w * scale for w in raw_col_widths]
    row_heights = [36.0, 22.0, 8.0, 22.0] + [18.0] * len(df)
    data = [
        ["Benchmark Report", "", "", "", ""],
        [f"Generated: {date.today()}", "", "", f"Rows: {_n_rows}", ""],
        ["", "", "", "", ""],
        _COLS,
        *[list(row) for row in df.iter_rows()],
    ]
    table = _RLTable(data, colWidths=col_widths, rowHeights=row_heights, repeatRows=0)
    table.setStyle(
        _RLStyle(
            [
                ("SPAN", (0, 0), (-1, 0)),
                ("SPAN", (0, 1), (2, 1)),
                ("SPAN", (3, 1), (4, 1)),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BACKGROUND", (0, 0), (-1, 0), _rl_colors.HexColor("#1E3A5F")),
                ("TEXTCOLOR", (0, 0), (-1, 0), _rl_colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 14),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("BACKGROUND", (0, 1), (-1, 1), _rl_colors.HexColor("#2563EB")),
                ("TEXTCOLOR", (0, 1), (-1, 1), _rl_colors.white),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, 1), 10),
                ("ALIGN", (0, 1), (2, 1), "LEFT"),
                ("ALIGN", (3, 1), (4, 1), "RIGHT"),
                ("LEFTPADDING", (0, 1), (2, 1), 10),
                ("BACKGROUND", (0, 3), (-1, 3), _rl_colors.HexColor("#1E3A5F")),
                ("TEXTCOLOR", (0, 3), (-1, 3), _rl_colors.white),
                ("FONTNAME", (0, 3), (-1, 3), "Helvetica-Bold"),
                ("FONTSIZE", (0, 3), (-1, 3), 10),
                ("GRID", (0, 3), (-1, 3), 0.5, _rl_colors.white),
                ("FONTNAME", (0, 4), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 4), (-1, -1), 10),
                ("TEXTCOLOR", (0, 4), (-1, -1), _rl_colors.HexColor("#1E3A5F")),
                ("GRID", (0, 4), (-1, -1), 0.5, _rl_colors.HexColor("#CCCCCC")),
            ]
        )
    )
    _RLDoc(
        str(out),
        pagesize=_RL_A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36,
    ).build([table])


# Â§5. Benchmark Orchestration
def _xlsx_registry(schema) -> list[tuple[str, object, bool]]:
    return [
        (
            "mindoff fidelity (XLSX)",
            partial(_bench_mindoff_xlsx_fidelity, schema),
            True,
        ),
        (
            "mindoff streaming-openpyxl (XLSX)",
            partial(_bench_mindoff_xlsx_streaming_openpyxl, schema),
            True,
        ),
        (
            "mindoff streaming-xlsxwriter (XLSX)",
            partial(_bench_mindoff_xlsx_streaming_xlsxwriter, schema),
            _HAS_XLSXWRITER,
        ),
        ("openpyxl â€“ manual styling", _bench_openpyxl_direct, _HAS_OPENPYXL),
        ("xlsxwriter â€“ manual styling", _bench_xlsxwriter_direct, _HAS_XLSXWRITER),
    ]


def _pdf_registry(schema) -> list[tuple[str, object, bool]]:
    return [
        ("mindoff export (PDF)", partial(_bench_mindoff_pdf, schema), True),
        ("reportlab â€“ manual styling", _bench_reportlab_direct, _HAS_REPORTLAB),
    ]


def run_benchmarks(schema) -> list[BenchResult]:
    results: list[BenchResult] = []

    def _run_one(name: str, fn, fmt: str, ext: str, n: int, df_path: Path) -> None:
        artifact_path = _artifact_path(name, fmt, n, ext)
        elapsed_s, stdev_e, peak_mb, stdev_m, file_mb, compile_s, status, detail = (
            _timed_runs(partial(fn, df_path, n), ext, keep_output_path=artifact_path)
        )
        results.append(
            BenchResult(
                name,
                fmt,
                n,
                elapsed_s,
                peak_mb,
                file_mb,
                status=status,
                runs=ACTIVE_RUNS_PER_POINT,
                compile_s=compile_s,
                elapsed_stdev_s=stdev_e,
                peak_stdev_mb=stdev_m,
                output_file=str(artifact_path.relative_to(OUTPUT_DIR))
                if status == "completed"
                else "",
            )
        )
        if status == "completed":
            print(
                f"  {name:<42} {n:>10,} rows  median {elapsed_s:7.2f}s  "
                f"{peak_mb:7.1f} MB  saved {artifact_path.relative_to(OUTPUT_DIR)}"
            )
        else:
            print(f"  [{status.upper()}] {name} @ {n:,}: {detail}")

    print("\n=== XLSX runtime benchmarks (compile + export) ===")
    for name, fn, available in _xlsx_registry(schema):
        if not available:
            print(f"  [SKIP] {name}")
            results.append(
                BenchResult(
                    name,
                    "xlsx",
                    0,
                    float("nan"),
                    float("nan"),
                    0.0,
                    status="skipped",
                    runs=0,
                )
            )
            continue
        for n in ACTIVE_XLSX_ROW_COUNTS:
            _run_one(name, fn, "xlsx", "xlsx", n, _make_parquet(n))

    print("\n=== PDF runtime benchmarks (compile + export) ===")
    for name, fn, available in _pdf_registry(schema):
        if not available:
            print(f"  [SKIP] {name}")
            results.append(
                BenchResult(
                    name,
                    "pdf",
                    0,
                    float("nan"),
                    float("nan"),
                    0.0,
                    status="skipped",
                    runs=0,
                )
            )
            continue
        for n in ACTIVE_PDF_ROW_COUNTS:
            _run_one(name, fn, "pdf", "pdf", n, _make_parquet(n))

    return results


# Â§6. Outputs
def save_csv(results: list[BenchResult]) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_DIR / "results.csv"
    with open(out, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(
            [
                "method",
                "format",
                "rows",
                "elapsed_s",
                "elapsed_stdev_s",
                "peak_mb",
                "peak_stdev_mb",
                "compile_s",
                "file_mb",
                "output_file",
                "rows_per_s",
                "peak_mb_per_10k_rows",
                "file_mb_per_10k_rows",
                "status",
                "runs",
            ]
        )
        for r in results:
            rows_per_s = (
                r.rows / r.elapsed_s
                if r.rows > 0 and r.status == "completed" and r.elapsed_s > 0
                else float("nan")
            )
            peak_mb_per_10k = (
                r.peak_mb / (r.rows / 10_000)
                if r.rows > 0 and r.status == "completed"
                else float("nan")
            )
            file_mb_per_10k = (
                r.file_mb / (r.rows / 10_000)
                if r.rows > 0 and r.status == "completed"
                else float("nan")
            )
            w.writerow(
                [
                    r.method,
                    r.fmt,
                    r.rows,
                    r.elapsed_s,
                    r.elapsed_stdev_s,
                    r.peak_mb,
                    r.peak_stdev_mb,
                    r.compile_s,
                    r.file_mb,
                    r.output_file,
                    rows_per_s,
                    peak_mb_per_10k,
                    file_mb_per_10k,
                    r.status,
                    r.runs,
                ]
            )
    return out


# §7. Entrypoint
def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run public benchmark.")
    profile = parser.add_mutually_exclusive_group()
    profile.add_argument(
        "--quick",
        action="store_true",
        help="Fast dev profile (small rows, single run, shorter timeout).",
    )
    profile.add_argument(
        "--full",
        action="store_true",
        help="Publish profile (default).",
    )
    parser.add_argument(
        "--runs", type=int, default=None, help="Override runs per point."
    )
    parser.add_argument(
        "--timeout", type=int, default=None, help="Override timeout seconds per run."
    )
    return parser.parse_args()


def _apply_cli_profile(args: argparse.Namespace) -> None:
    global ACTIVE_XLSX_ROW_COUNTS
    global ACTIVE_PDF_ROW_COUNTS
    global ACTIVE_RUNS_PER_POINT
    global ACTIVE_RUN_TIMEOUT_S

    # Default publish profile.
    ACTIVE_XLSX_ROW_COUNTS = XLSX_ROW_COUNTS
    ACTIVE_PDF_ROW_COUNTS = PDF_ROW_COUNTS
    ACTIVE_RUNS_PER_POINT = RUNS_PER_POINT
    ACTIVE_RUN_TIMEOUT_S = RUN_TIMEOUT_S

    if args.quick:
        ACTIVE_XLSX_ROW_COUNTS = [1_000, 10_000]
        ACTIVE_PDF_ROW_COUNTS = [1_000]
        ACTIVE_RUNS_PER_POINT = 1
        ACTIVE_RUN_TIMEOUT_S = 90

    if args.runs is not None:
        ACTIVE_RUNS_PER_POINT = max(1, args.runs)
    if args.timeout is not None:
        ACTIVE_RUN_TIMEOUT_S = max(30, args.timeout)


def main() -> None:
    args = _parse_args()
    _apply_cli_profile(args)

    if not TEMPLATE_PATH.exists():
        sys.exit(
            f"Template not found: {TEMPLATE_PATH}\nRun first: python examples/benchmark/create_template.py"
        )

    print(f"Template: {TEMPLATE_PATH}")
    print(f"Output:   {OUTPUT_DIR}")
    print(
        "Profile:  "
        f"xlsx_rows={[_format_rows(n) for n in ACTIVE_XLSX_ROW_COUNTS]}, "
        f"pdf_rows={[_format_rows(n) for n in ACTIVE_PDF_ROW_COUNTS]}, "
        f"runs={ACTIVE_RUNS_PER_POINT}, "
        f"warmup={'yes' if ACTIVE_RUNS_PER_POINT >= 3 else 'no'}, "
        f"timeout={ACTIVE_RUN_TIMEOUT_S}s"
    )
    print("Preparing schema once (excluded from timed benchmark runs)...")
    schema = mo_dataport.extract(str(TEMPLATE_PATH))
    if ARTIFACT_DIR.exists():
        shutil.rmtree(ARTIFACT_DIR)

    results = run_benchmarks(schema)
    csv_path = save_csv(results)

    print(f"\nResults CSV:  {csv_path}")


if __name__ == "__main__":
    main()
