"""Command line entry point (no GUI needed).

Examples:
    python cli.py "D:\\Documents" --excel report.xlsx
    python cli.py "D:\\Documents" --no-word --threshold 80
"""

from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docstats.analyzer import (  # noqa: E402
    DEFAULT_COMPLEX_THRESHOLD,
    DEFAULT_FIGURE_MIN_AREA,
    analyze_folder,
)
from docstats.export import export_csv, export_xlsx, page_text  # noqa: E402


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Count tables and figures in Word and PDF documents"
    )
    p.add_argument("folder", help="Folder (or a single file) to scan")
    p.add_argument("--no-recursive", action="store_true", help="Do not scan subfolders")
    p.add_argument(
        "--no-word",
        action="store_true",
        help="Word files: do not use MS Word (faster, but no page numbers / split detection)",
    )
    p.add_argument(
        "--pdf-borderless",
        action="store_true",
        help="PDF: also detect tables without ruling lines (more false positives)",
    )
    p.add_argument(
        "--threshold",
        type=int,
        default=DEFAULT_COMPLEX_THRESHOLD,
        help=f"Cells a table needs to count as complex (default {DEFAULT_COMPLEX_THRESHOLD})",
    )
    p.add_argument(
        "--fig-min-area",
        type=float,
        default=DEFAULT_FIGURE_MIN_AREA * 100,
        metavar="PERCENT",
        help=(
            "Minimum %% of the page a figure must cover to be counted "
            f"(default {DEFAULT_FIGURE_MIN_AREA * 100:.0f} = a quarter of the page, "
            "which drops logos; use 0 to count every figure)"
        ),
    )
    p.add_argument("--excel", metavar="FILE.xlsx", help="Write an Excel report")
    p.add_argument("--csv", metavar="FOLDER", help="Write three CSV files into a folder")
    p.add_argument("--list", action="store_true", help="Print every table with an issue")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    def progress(done: int, total: int, path: str) -> None:
        if done:
            print(f"  [{done}/{total}] {os.path.basename(path)}", flush=True)

    result = analyze_folder(
        args.folder,
        recursive=not args.no_recursive,
        threshold=args.threshold,
        figure_min_area=max(0.0, min(100.0, args.fig_min_area)) / 100,
        use_word=not args.no_word,
        pdf_borderless=args.pdf_borderless,
        progress=progress,
    )
    t = result.totals()

    print()
    print("=" * 62)
    print(f"  Folder              : {result.root}")
    print(f"  Complex threshold   : >= {result.complex_threshold} cells")
    print(
        "  Figure threshold    : "
        + (
            f">= {result.figure_min_area * 100:.0f}% of the page"
            if result.figure_min_area > 0
            else "count every figure"
        )
    )
    print("-" * 62)
    print(
        f"  Files               : {t['files']}  "
        f"(Word: {t['files_word']}, PDF: {t['files_pdf']}, errors: {t['files_error']})"
    )
    print(f"  Pages               : {t['pages']}")
    print(f"  Tables              : {t['tables']}  (top-level: {t['tables_top']})")
    print(f"  Figures             : {t['figures']}  ({t['figures_small']} small ones dropped)")
    print(f"  Complex tables      : {t['complex']}")
    print(f"  Split across pages  : {t['split']}")
    print(f"  Nested tables       : {t['nested_parent']} parents / {t['nested_child']} children")
    print(f"  With merged cells   : {t['merged']}")
    print(f"  Split w/o header    : {t['split_no_header']}")
    print(f"  Total cells         : {t['cells']}")
    print("=" * 62)

    if args.list:
        for tb in result.all_tables():
            if not tb.issue_tags():
                continue
            print(
                f"  {tb.file_name} | {tb.label} | page {page_text(tb)} | "
                f"{tb.rows}x{tb.grid_cols}={tb.cells} cells | {'; '.join(tb.issue_tags())}"
            )

    for doc in result.failed_docs:
        print(f"  [ERROR] {doc.path}: {doc.error}")

    if args.excel:
        print("Saved:", export_xlsx(result, args.excel))
    if args.csv:
        for path in export_csv(result, args.csv):
            print("Saved:", path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
