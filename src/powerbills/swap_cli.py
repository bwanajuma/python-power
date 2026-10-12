"""Command-line interface for swap dump Net SOC summaries."""

from __future__ import annotations

import argparse
import time
import zipfile
from pathlib import Path

from .swapdata import (
    SwapDumpError,
    SwapReport,
    aggregate_swap_dump,
    combine_reports,
    discover_dumps,
    load_station_names,
    match_stations,
    summary_rows,
    write_csv,
    write_summary,
)

DEFAULT_INPUTS = ("Swap Dump/*.xlsx",)
DEFAULT_TEMPLATE = "excel/Power Bills Template.xlsx"
DEFAULT_OUTPUT = "excel/Swap Net SOC Summary.xlsx"
DEFAULT_CSV = "excel/Swap Net SOC Summary.csv"


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
        "--template",
        default=DEFAULT_TEMPLATE,
        help=f"workbook holding the Matched Station Name list (default: {DEFAULT_TEMPLATE})",
    )
    parser.add_argument(
        "--template-sheet",
        help="worksheet in the template; defaults to its first sheet",
    )
    parser.add_argument(
        "--output", default=DEFAULT_OUTPUT, help=f"summary workbook (default: {DEFAULT_OUTPUT})"
    )
    parser.add_argument("--csv", default=DEFAULT_CSV, help=f"summary CSV (default: {DEFAULT_CSV})")
    parser.add_argument(
        "--no-template",
        action="store_true",
        help="skip station matching and report raw swap station names",
    )
    parser.add_argument("--quiet", action="store_true", help="only print the output paths")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Aggregate swap dumps and write the per-station Net SOC summary."""
    args = parse_args(argv)
    dumps = discover_dumps(args.paths)
    if not dumps:
        print(f"no swap dump workbooks matched: {' '.join(args.paths)}")
        return 1

    known: dict[str, str] = {}
    if not args.no_template:
        template = Path(args.template)
        if not template.is_file():
            print(f"template not found: {template} (use --no-template to skip matching)")
            return 1
        try:
            known = load_station_names(template, args.template_sheet)
        except SwapDumpError as error:
            print(f"error: {error}")
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
            unmatched = sum(1 for _, matched in match_stations(report, known) if matched is None)
            print(
                f"{dump.name}: sheet={report.sheet} period={report.period} "
                f"rows={report.rows_scanned} swaps={report.swaps} "
                f"stations={len(report.stations)} unmatched={unmatched} "
                f"net_soc={report.net_soc:,.0f} ({time.monotonic() - started:.1f}s)"
            )

    rows = combine_reports(reports, known) if known else summary_rows(reports, known)
    output = write_summary(rows, reports, args.output, known or None)
    csv_path = write_csv(rows, args.csv)
    if not args.quiet:
        matched = sum(1 for row in rows if row["Matched Station Name"])
        print(f"rows={len(rows)} matched={matched} unmatched={len(rows) - matched}")
    print(f"workbook: {output}")
    print(f"csv: {csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
