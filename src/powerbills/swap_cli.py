"""Command-line interface for swap dump Net SOC summaries."""

from __future__ import annotations

import argparse
import time
import zipfile

from .swapdata import (
    SwapDumpError,
    SwapReport,
    aggregate_swap_dump,
    combine_reports,
    discover_dumps,
    write_csv,
    write_summary,
)

DEFAULT_INPUTS = ("Swap Dump/*.xlsx",)
DEFAULT_OUTPUT = "excel/Swap Station Business Metrics.xlsx"
DEFAULT_CSV = "excel/Swap Station Business Metrics.csv"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse the swap summary command line."""
    parser = argparse.ArgumentParser(prog="kplc-swap", description=__doc__)
    parser.add_argument(
        "paths",
        nargs="*",
        default=list(DEFAULT_INPUTS),
        help="swap dump workbooks, directories, or globs (default: Swap Dump/*.xlsx)",
    )
    parser.add_argument(
        "--period",
        help="override the period for every input, as YYYY-MM (default: month in filename)",
    )
    parser.add_argument(
        "--year",
        type=int,
        help="year to use when the period comes from the filename (default: this year)",
    )
    parser.add_argument(
        "--output", default=DEFAULT_OUTPUT, help=f"summary workbook (default: {DEFAULT_OUTPUT})"
    )
    parser.add_argument("--csv", default=DEFAULT_CSV, help=f"summary CSV (default: {DEFAULT_CSV})")
    parser.add_argument("--quiet", action="store_true", help="only print the output paths")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Aggregate swap dumps and write the per-station Net SOC summary."""
    args = parse_args(argv)
    dumps = discover_dumps(args.paths)
    if not dumps:
        print(f"no swap dump workbooks matched: {' '.join(args.paths)}")
        return 1

    reports: list[SwapReport] = []
    for dump in dumps:
        started = time.monotonic()
        try:
            report = aggregate_swap_dump(dump, period=args.period, year=args.year)
        except (SwapDumpError, ValueError, zipfile.BadZipFile, OSError) as error:
            print(f"error: {dump} ({error})")
            return 1
        reports.append(report)
        if not args.quiet:
            print(
                f"{dump.name}: sheet={report.sheet} period={report.period} "
                f"rows={report.rows_scanned} swaps={report.swaps} "
                f"stations={len(report.stations)} "
                f"net_soc={report.net_soc:,.0f} ({time.monotonic() - started:.1f}s)"
            )

    rows = combine_reports(reports)
    output = write_summary(rows, reports, args.output)
    csv_path = write_csv(rows, args.csv)
    if not args.quiet:
        print(f"stations={len(rows)} swaps={sum(int(row['Swaps']) for row in rows)}")
    print(f"workbook: {output}")
    print(f"csv: {csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
