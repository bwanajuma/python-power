"""Command-line interface for the KPLC PDF renamer."""

from __future__ import annotations

import argparse
from pathlib import Path

from .renamer import (
    BillData,
    discover_pdfs,
    extract_bill_data,
    load_station_lookup,
    merge_bill_data,
    ocr_pdf,
    process_files,
)

DEFAULT_INPUT_DIRECTORY = "raw_files"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "paths",
        nargs="*",
        default=[DEFAULT_INPUT_DIRECTORY],
        help=f"PDF files, directories, or glob patterns (default: {DEFAULT_INPUT_DIRECTORY})",
    )
    run_mode = parser.add_mutually_exclusive_group()
    run_mode.add_argument("--apply", action="store_true", help="perform the renames")
    run_mode.add_argument(
        "--extract-only",
        action="store_true",
        help="extract bill fields to the output workbook without rename planning",
    )
    parser.add_argument(
        "--workbook",
        default="excel/Power Bills Template.xlsx",
        help="Excel template containing station names and account numbers",
    )
    parser.add_argument(
        "--lookup-sheet",
        default="Septe",
        help="worksheet containing station and account lookup data",
    )
    parser.add_argument(
        "--output-sheet",
        default="September",
        help="worksheet where extracted bill values are written",
    )
    parser.add_argument(
        "--output",
        default="excel/September KPLC Power Bills.xlsx",
        help="output workbook for extracted bill fields",
    )
    parser.add_argument(
        "--on-conflict",
        choices=("skip", "error"),
        default="skip",
        help="what to do when the destination already exists (default: skip)",
    )
    args = parser.parse_args()

    files = discover_pdfs(args.paths)
    if not files:
        print("skipped: no PDF files found")
        return 0

    workbook = Path(args.workbook)
    try:
        station_lookup = (
            load_station_lookup(workbook, args.lookup_sheet)
            if workbook.exists() and not args.extract_only
            else {}
        )
    except (RuntimeError, ValueError) as error:
        parser.error(str(error))
    ocr_cache: dict[Path, str] = {}

    def cached_ocr(path: Path) -> str:
        if path not in ocr_cache:
            ocr_cache[path] = ocr_pdf(path)
        return ocr_cache[path]

    bills: list[BillData] = []
    has_errors = False
    if args.extract_only:
        for path in files:
            try:
                text = cached_ocr(path)
            except Exception as error:
                print(f"error: {path} ({error})")
                has_errors = True
                continue
            bill = extract_bill_data(text)
            if bill is None:
                print(f"skipped: {path} (account number not found)")
            else:
                bills.append(bill)
                print(f"extracted: {path} ({bill.account_number})")
    else:
        results = process_files(
            files,
            apply=args.apply,
            on_conflict=args.on_conflict,
            station_lookup=station_lookup,
            ocr=cached_ocr,
        )
        for result in results:
            destination = f" -> {result.destination}" if result.destination else ""
            print(f"{result.status}: {result.source}{destination} ({result.message})")
        for text in ocr_cache.values():
            bill = extract_bill_data(text)
            if bill is not None:
                bills.append(bill)
        has_errors = any(result.status == "error" for result in results)
    if bills:
        merge_bill_data(workbook, Path(args.output), bills, args.output_sheet)
        print(f"merged: {len(bills)} bill(s) -> {args.output}")
    return 1 if has_errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
