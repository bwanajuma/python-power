"""OCR and collision-safe renaming for KPLC bill PDFs."""

from __future__ import annotations

import glob
import re
import shutil
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ACCOUNT_LABEL = re.compile(
    r"(?:account|a/c)\s*(?:number|no\.?|#)?\s*[:\-]?\s*([0-9][0-9\s\-]{4,20})",
    re.IGNORECASE,
)
ACCOUNT_NUMBER = re.compile(r"\b[0-9]{6,14}\b")
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')
TARIFF_CLASSIFICATION = re.compile(
    r"Domestic\s+Lifeline(?:\s*\([^)]*\))?"
    r"|Domestic\s+Ordinary\s+[12](?:\s*\([^)]*\))?"
    r"|Electric\s+Vehicles?(?:\s*\(E[- ]?Mobility\))?"
    r"|E[- ]Mobility"
    r"|Small\s+Commercial(?:\s*\(?\s*SC\d+\s*\)?)?",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class RenameResult:
    source: Path
    destination: Path | None
    status: str
    message: str


@dataclass(frozen=True)
class BillData:
    account_number: str
    consumption_period: str | None
    consumption: float | None
    bill: float | None
    due_date: str | None
    issue_date: str | None
    low_rate_consumption: float | None
    high_rate_consumption: float | None
    method_of_charge: str | None = None


def normalize_account_number(value: object) -> str | None:
    """Convert Excel numeric/text account values to comparable digits."""
    if value is None:
        return None
    digits = re.sub(r"\D", "", str(value))
    return digits or None


def resolve_sheet_name(workbook: Any, requested: str) -> str:
    """Use the requested worksheet, or the only worksheet in a changing template."""
    sheet_names = list(workbook.sheetnames)
    if requested in sheet_names:
        return requested
    if len(sheet_names) == 1:
        return sheet_names[0]
    raise ValueError(f"worksheet not found: {requested}; available: {', '.join(sheet_names)}")


def load_station_lookup(workbook: Path, sheet_name: str = "August") -> dict[str, str]:
    """Load account number to station name mappings from an Excel sheet."""
    try:
        from openpyxl import load_workbook
    except ImportError as error:
        raise RuntimeError("Excel support is missing; run `python -m pip install -e .`") from error

    book = load_workbook(workbook, read_only=True, data_only=True)
    try:
        sheet = book[resolve_sheet_name(book, sheet_name)]
        headers = [cell.value for cell in next(sheet.iter_rows())]
        try:
            station_column = headers.index("Station")
            account_column = headers.index("Account Number")
        except ValueError as error:
            raise ValueError(
                "worksheet must contain 'Station+OA1:P387' and 'Account Number' columns"
            ) from error

        lookup: dict[str, str] = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            account = normalize_account_number(row[account_column])
            station = str(row[station_column]).strip() if row[station_column] else ""
            if account and station:
                lookup[account] = station
        return lookup
    finally:
        book.close()


def safe_filename(value: str) -> str:
    """Make a station name safe to use as a PDF filename."""
    cleaned = INVALID_FILENAME_CHARS.sub("_", value).strip().rstrip(".")
    return cleaned or "unnamed"


def extract_account_number(text: str) -> str | None:
    """Return the first account number associated with an account label."""
    for line in text.splitlines():
        match = ACCOUNT_LABEL.search(line)
        if match:
            digits = re.sub(r"\D", "", match.group(1))
            if 6 <= len(digits) <= 14:
                return digits

        if re.search(r"\baccount\b|\ba/c\b", line, re.IGNORECASE):
            number = ACCOUNT_NUMBER.search(line)
            if number:
                return number.group()
    return None


def _extract_decimal(pattern: str, text: str) -> float | None:
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return None
    return float(match.group(1).replace(",", ""))


def extract_bill_data(text: str) -> BillData | None:
    """Extract the requested bill fields from OCR text."""
    account = extract_account_number(text)
    if account is None:
        return None

    period_match = re.search(r"Consumption\s+Period\s*:\s*([^\n\r]+)", text, re.IGNORECASE)
    due_match = re.search(
        r"(?:Date\s+Due|Due\s+Date)\s*:\s*(\d{1,2}/\d{1,2}/\d{4})",
        text,
        re.IGNORECASE,
    )
    issue_match = re.search(
        r"(?:Issue\s+Date|Date\s+of\s+Issue)\s*:\s*(\d{1,2}/\d{1,2}/\d{4})",
        text,
        re.IGNORECASE,
    )
    method_match = re.search(r"Method\s+of\s+Charge\s*:\s*([^\n\r]+)", text, re.IGNORECASE)
    tariff_match = TARIFF_CLASSIFICATION.search(method_match.group(1)) if method_match else None
    return BillData(
        account_number=account,
        consumption_period=period_match.group(1).strip() if period_match else None,
        consumption=_extract_decimal(
            r"Fuel\s+Energy\s+Cost\s*:?\s*(?:K?shs?\.?\s*)?([\d,.]+)", text
        ),
        bill=_extract_decimal(r"ELECTRICITY\s+BILL\s*:\s*K?sh\s*([\d,.]+)", text)
        or _extract_decimal(r"Total\s+Monthly\s+Bill\s+([\d,.]+)", text),
        due_date=due_match.group(1) if due_match else None,
        issue_date=issue_match.group(1) if issue_match else None,
        low_rate_consumption=_extract_decimal(r"LowRateConsumption\s+([\d,.]+)\s*kWh", text),
        high_rate_consumption=_extract_decimal(r"HighRateConsumption\s+([\d,.]+)\s*kWh", text),
        method_of_charge=tariff_match.group(0).strip() if tariff_match else None,
    )


def merge_bill_data(
    template: Path,
    output: Path,
    bills: Iterable[BillData],
    sheet_name: str = "September",
) -> None:
    """Copy the template and merge extracted bill fields into matching rows."""
    try:
        from openpyxl import load_workbook
    except ImportError as error:
        raise RuntimeError("Excel support is missing; run `python -m pip install -e .`") from error

    output.parent.mkdir(parents=True, exist_ok=True)
    if template.resolve() != output.resolve():
        shutil.copy2(template, output)
    book = load_workbook(output)
    try:
        sheet = book[resolve_sheet_name(book, sheet_name)]
        headers: dict[str, int] = {}
        for cell in sheet[1]:
            if cell.value and cell.column is not None:
                headers[str(cell.value)] = cell.column
        required = ("Station", "Account Number")
        missing = [header for header in required if header not in headers]
        if missing:
            raise ValueError(f"worksheet missing columns: {', '.join(missing)}")

        fields = {
            "Consumption Period": "consumption_period",
            "consumption": "consumption",
            "Bill": "bill",
            "Due Date": "due_date",
            "Issue Date": "issue_date",
            "Method of Charge": "method_of_charge",
            "LowRateConsumption": "low_rate_consumption",
            "HighRateConsumption": "high_rate_consumption",
        }
        for header in fields:
            if header not in headers:
                column = sheet.max_column + 1
                sheet.cell(1, column, header)
                headers[header] = column

        rows_by_account: dict[str, int] = {}
        for row in range(2, sheet.max_row + 1):
            account = normalize_account_number(sheet.cell(row, headers["Account Number"]).value)
            if account:
                rows_by_account[account] = row

        for bill in bills:
            row = rows_by_account.get(bill.account_number)
            if row is None:
                row = sheet.max_row + 1
                sheet.cell(row, headers["Account Number"], bill.account_number)
                sheet.cell(row, headers["Station"], bill.account_number)
                rows_by_account[bill.account_number] = row
            for header, attribute in fields.items():
                sheet.cell(row, headers[header], getattr(bill, attribute))
        book.save(output)
    finally:
        book.close()


def ocr_pdf(path: Path) -> str:
    """OCR the account header first, then the full page as a layout fallback."""
    try:
        import fitz
        import pytesseract
        from PIL import Image
    except ImportError as error:
        raise RuntimeError(
            "OCR dependencies are missing; run `python -m pip install -e .`"
        ) from error

    pages: list[str] = []
    with fitz.open(path) as document:
        for page in document:
            pixmap = page.get_pixmap(dpi=250, alpha=False)
            image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            header = image.crop((int(image.width * 0.45), 0, image.width, int(image.height * 0.24)))
            account_text = pytesseract.image_to_string(header, config="--psm 6")
            full_page_text = pytesseract.image_to_string(image)
            pages.append(f"{account_text}\n{full_page_text}")
    return "\n".join(pages)


def discover_pdfs(inputs: Iterable[str]) -> list[Path]:
    """Expand files, directories, and shell-style paths into sorted PDFs."""
    paths: set[Path] = set()
    for value in inputs:
        path = Path(value)
        if path.is_dir():
            paths.update(item for item in path.glob("*.pdf") if item.is_file())
        elif any(char in value for char in "*?[]"):
            paths.update(Path(item) for item in glob.glob(value) if Path(item).is_file())
        elif path.is_file() and path.suffix.lower() == ".pdf":
            paths.add(path)
    return sorted(paths)


def process_files(
    paths: Iterable[Path],
    *,
    apply: bool = False,
    on_conflict: str = "skip",
    ocr: Callable[[Path], str] = ocr_pdf,
    station_lookup: dict[str, str] | None = None,
) -> list[RenameResult]:
    """Extract and optionally apply account-number-based PDF renames."""
    results: list[RenameResult] = []
    planned: set[Path] = set()
    for source in paths:
        try:
            account = extract_account_number(ocr(source))
        except Exception as error:
            results.append(RenameResult(source, None, "error", str(error)))
            continue
        if account is None:
            results.append(RenameResult(source, None, "not-found", "account number not found"))
            continue

        filename = safe_filename(
            station_lookup.get(account, account) if station_lookup else account
        )
        destination = source.with_name(f"{filename}.pdf")
        if destination != source and (destination.exists() or destination in planned):
            message = f"destination exists: {destination.name}"
            if on_conflict == "error":
                raise FileExistsError(message)
            results.append(RenameResult(source, destination, "conflict", message))
            continue

        planned.add(destination)
        if apply and destination != source:
            source.rename(destination)
        status = "renamed" if apply and destination != source else "planned"
        results.append(RenameResult(source, destination, status, account))
    return results
