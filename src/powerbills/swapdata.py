"""Streaming aggregation of large KPLC swap dump workbooks.

One month of swaps is roughly 900k rows and 21M cells; the worksheet XML inside
the .xlsx container expands to about 800 MB. Loading that with openpyxl means
tens of minutes and gigabytes of RAM, so the sheet is never materialised as a
worksheet. Instead the sheet part is streamed out of the zip container in fixed
chunks and only the handful of columns that are aggregated are decoded. Memory
stays flat regardless of dump size, and a month takes well under a minute.
"""

from __future__ import annotations

import calendar
import csv
import glob
import re
import zipfile
from collections.abc import Iterable, Iterator, Sequence
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import BinaryIO
from xml.etree.ElementTree import iterparse

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
EXCEL_EPOCH = datetime(1899, 12, 30)
CHUNK_BYTES = 1 << 23
MONTHS = {name.lower(): number for number, name in enumerate(calendar.month_abbr) if name}

DATE_COLUMN = "Swap DateTime"
STATION_COLUMN = "Station"
NET_SOC_COLUMN = "Net SOC"
REQUIRED_HEADERS = (DATE_COLUMN, STATION_COLUMN, NET_SOC_COLUMN)
METRIC_COLUMNS = (NET_SOC_COLUMN, "Billable SOC", "Swap Cost")
AGGREGATED_HEADERS = (DATE_COLUMN, STATION_COLUMN, *METRIC_COLUMNS)
SUMMARY_HEADERS = (
    "Period",
    "Matched Station Name",
    "Swap Station",
    "Swaps",
    "Net SOC",
    "Billable SOC",
    "Swap Cost",
    "Average Net SOC",
    "First Swap",
    "Last Swap",
)

MONTH_IN_NAME = re.compile(r"jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec", re.I)
VALUE = re.compile(rb"<v>([^<]*)</v>")
INLINE_TEXT = re.compile(rb"<t[^>]*>(.*?)</t>", re.S)
SHEET = re.compile(r"<sheet[^>]*name=\"([^\"]+)\"[^>]*r:id=\"([^\"]+)\"")
ELEMENT = re.compile(r"<([A-Za-z]+)\s([^>]*)>")
ATTRIBUTE = re.compile(r'([A-Za-z:]+)="([^"]*)"')


class SwapDumpError(RuntimeError):
    """Raised when a workbook cannot be read as a swap dump."""


@dataclass
class StationTotals:
    """Accumulated swap metrics for a single station."""

    station: str
    swaps: int = 0
    net_soc: float = 0.0
    billable_soc: float = 0.0
    swap_cost: float = 0.0
    first_swap: datetime | None = None
    last_swap: datetime | None = None

    @property
    def average_net_soc(self) -> float:
        return self.net_soc / self.swaps if self.swaps else 0.0

    def merge(self, other: StationTotals) -> None:
        """Add another period's totals to this one."""
        self.swaps += other.swaps
        self.net_soc += other.net_soc
        self.billable_soc += other.billable_soc
        self.swap_cost += other.swap_cost
        self.first_swap = earliest(self.first_swap, other.first_swap)
        self.last_swap = latest(self.last_swap, other.last_swap)


@dataclass
class SwapReport:
    """Station totals for one workbook, plus the metadata needed to trust them."""

    source: Path
    sheet: str
    period: str | None
    period_start: datetime | None
    period_end: datetime | None
    rows_scanned: int = 0
    rows_in_period: int = 0
    rows_outside_period: int = 0
    observed_start: datetime | None = None
    observed_end: datetime | None = None
    stations: dict[str, StationTotals] = field(default_factory=dict)

    @property
    def net_soc(self) -> float:
        """Total Net SOC across every station."""
        return sum(item.net_soc for item in self.stations.values())

    @property
    def billable_soc(self) -> float:
        """Total Billable SOC across every station."""
        return sum(item.billable_soc for item in self.stations.values())

    @property
    def swap_cost(self) -> float:
        """Total Swap Cost across every station."""
        return sum(item.swap_cost for item in self.stations.values())

    @property
    def swaps(self) -> int:
        """Total in-period swaps across every station."""
        return sum(item.swaps for item in self.stations.values())


def earliest(current: datetime | None, candidate: datetime | None) -> datetime | None:
    """Return the earlier of two optional datetimes."""
    if candidate is None:
        return current
    return candidate if current is None or candidate < current else current


def latest(current: datetime | None, candidate: datetime | None) -> datetime | None:
    """Return the later of two optional datetimes."""
    if candidate is None:
        return current
    return candidate if current is None or candidate > current else current


def column_letters(index: int) -> str:
    """Convert a one-based column index to its spreadsheet letters."""
    if index < 1:
        raise ValueError("column index must be one-based")
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def serial_to_datetime(value: float) -> datetime:
    """Convert an Excel date serial to a datetime, rounding away float drift."""
    return EXCEL_EPOCH + timedelta(seconds=round(value * 86400))


def normalize_station(value: object) -> str:
    """Fold a station name to a key that ignores case and spacing noise."""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def period_bounds(period: str) -> tuple[datetime, datetime]:
    """Return the half-open [start, end) datetime range of a YYYY-MM period."""
    match = re.fullmatch(r"(\d{4})-(\d{1,2})", period.strip())
    if not match:
        raise ValueError(f"period must look like YYYY-MM, got {period!r}")
    year, month = int(match.group(1)), int(match.group(2))
    if not 1 <= month <= 12:
        raise ValueError(f"period month out of range in {period!r}")
    start = datetime(year, month, 1)
    end = datetime(year + 1, 1, 1) if month == 12 else datetime(year, month + 1, 1)
    return start, end


def period_from_filename(path: Path, year: int) -> str | None:
    """Infer a YYYY-MM period from a name such as August_Swap_Data.xlsx."""
    match = MONTH_IN_NAME.search(path.stem)
    if match is None:
        return None
    return f"{year:04d}-{MONTHS[match.group(0).lower()]:02d}"


def _worksheet_parts(book: zipfile.ZipFile) -> Iterator[tuple[str, str]]:
    """Yield (sheet name, worksheet part) pairs in workbook order."""
    try:
        workbook = book.read("xl/workbook.xml").decode("utf-8")
        rels = book.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    except KeyError as error:
        raise SwapDumpError("workbook is missing its sheet index") from error

    names = book.namelist()
    targets = _attributes(rels, "Relationship")
    for sheet in SHEET.finditer(workbook):
        target = targets.get(sheet.group(2))
        if target is None:
            continue
        part = target.lstrip("/") if target.startswith("/") else f"xl/{target.lstrip('./')}"
        if part in names:
            yield sheet.group(1), part


def _attributes(xml: str, element: str) -> dict[str, str]:
    """Return tag attribute maps, keyed by Id, for repeated XML elements."""
    found: dict[str, str] = {}
    for tag, body in ELEMENT.findall(xml):
        if tag != element:
            continue
        attributes = dict(ATTRIBUTE.findall(body))
        if "Id" in attributes:
            found[attributes["Id"]] = attributes.get("Target", "")
    return found


def _shared_strings(book: zipfile.ZipFile) -> list[str]:
    """Load the shared string table, or an empty table when there is none."""
    if "xl/sharedStrings.xml" not in book.namelist():
        return []
    strings: list[str] = []
    with book.open("xl/sharedStrings.xml") as stream:
        for _, element in iterparse(stream, events=("end",)):
            if element.tag == f"{MAIN_NS}si":
                strings.append("".join(node.text or "" for node in element.iter(f"{MAIN_NS}t")))
                element.clear()
    return strings


def _iter_cells(
    stream: BinaryIO, letters: frozenset[str]
) -> Iterator[tuple[str, str, bool, bytes]]:
    """Yield (row, column, is_shared, value) for the wanted columns of a sheet.

    The sheet is read in chunks and a cell straddling a chunk boundary is
    carried into the next chunk, so a truncated match is never decoded. Inline
    strings arrive already decoded, shared strings keep their table index, and
    numbers arrive as raw bytes for the caller to parse.
    """
    alternation = "|".join(sorted(letters)).encode()
    pattern = re.compile(rb'<c r="(' + alternation + rb')(\d+)"([^>]*?)(?:/>|>(.*?)</c>)')
    carry = b""
    while True:
        chunk = stream.read(CHUNK_BYTES)
        buffer = carry + chunk if chunk else carry
        streaming = bool(chunk)
        position = 0
        while True:
            match = pattern.search(buffer, position)
            if match is None:
                break
            if streaming and match.end() >= len(buffer):
                break
            position = match.end()
            attributes = match.group(3)
            body = match.group(4) or b""
            if b't="inlineStr"' in attributes:
                text = INLINE_TEXT.search(body)
                value = text.group(1) if text else b""
                is_shared = False
            else:
                text = VALUE.search(body)
                value = text.group(1) if text else b""
                is_shared = b't="s"' in attributes
            yield match.group(2).decode(), match.group(1).decode(), is_shared, value
        if not streaming:
            return
        carry = buffer[position:]


def _iter_rows(
    stream: BinaryIO, letters: frozenset[str], headers: dict[str, str]
) -> Iterator[tuple[str, dict[str, tuple[bool, bytes]]]]:
    """Yield (row number, raw target cells) for every row after the header."""
    row_number = ""
    row: dict[str, tuple[bool, bytes]] = {}
    for number, letter, is_shared, raw in _iter_cells(stream, letters):
        if number != row_number:
            if row_number and row_number != "1":
                yield row_number, row
            row_number = number
            row = {}
        row[headers[letter]] = (is_shared, raw)
    if row_number and row_number != "1":
        yield row_number, row


def _header_row(stream: BinaryIO, strings: Sequence[str]) -> dict[str, str]:
    """Return the header name to column letter map from row 1.

    Repeated headers such as the three `Swap DateTime` columns keep the first
    column, which is the raw timestamp rather than a derived or truncated one.
    """
    letters = frozenset(column_letters(index) for index in range(1, 65))
    headers: dict[str, str] = {}
    for number, letter, is_shared, raw in _iter_cells(stream, letters):
        if number != "1":
            break
        headers.setdefault(_text(raw, is_shared, strings).strip(), letter)
    return headers


def _candidate_sheets(
    book: zipfile.ZipFile, strings: Sequence[str]
) -> list[tuple[str, str, dict[str, str]]]:
    """Find the worksheets whose header row looks like a swap dump."""
    candidates: list[tuple[str, str, dict[str, str]]] = []
    for name, part in _worksheet_parts(book):
        with book.open(part) as stream:
            headers = _header_row(stream, strings)
        if all(header in headers for header in REQUIRED_HEADERS):
            candidates.append((name, part, headers))
    return candidates


def _text(raw: bytes, is_shared: bool, strings: Sequence[str]) -> str:
    """Decode a cell that is known to hold text."""
    if not raw:
        return ""
    if is_shared:
        index = int(raw)
        return strings[index] if index < len(strings) else ""
    return raw.decode("utf-8", "replace")


def _number(raw: bytes) -> float:
    """Decode a numeric cell, treating blanks and text as zero."""
    try:
        return float(raw)
    except ValueError:
        return 0.0


def aggregate_swap_dump(
    source: Path | str,
    period: str | None = None,
    year: int | None = None,
) -> SwapReport:
    """Stream one swap dump workbook and total Net SOC per station.

    `period` filters on `Swap DateTime` and defaults to the month named in the
    filename, so August_Swap_Data.xlsx becomes the August of `year`. Rows that
    fall outside the period are counted but excluded from every total.
    """
    source = Path(source)
    if not source.is_file():
        raise SwapDumpError(f"swap dump not found: {source}")
    period = period or period_from_filename(source, year or datetime.now().year)
    start, end = period_bounds(period) if period else (None, None)
    report = SwapReport(source=source, sheet="", period=period, period_start=start, period_end=end)

    with zipfile.ZipFile(source) as book:
        strings = _shared_strings(book)
        candidates = _candidate_sheets(book, strings)
        if not candidates:
            raise SwapDumpError(
                f"{source.name}: no worksheet with columns {', '.join(REQUIRED_HEADERS)}"
            )
        report.sheet, part, headers = max(
            candidates, key=lambda candidate: book.getinfo(candidate[1]).file_size
        )
        columns = {headers[name]: name for name in AGGREGATED_HEADERS if name in headers}

        with book.open(part) as stream:
            for _, row in _iter_rows(stream, frozenset(columns), columns):
                if DATE_COLUMN in row:
                    report.rows_scanned += 1
                    _, raw = row[DATE_COLUMN]
                    swap_at = serial_to_datetime(_number(raw))
                    report.observed_start = earliest(report.observed_start, swap_at)
                    report.observed_end = latest(report.observed_end, swap_at)
                    if start is not None and not start <= swap_at < end:  # type: ignore[operator]
                        report.rows_outside_period += 1
                        continue
                    report.rows_in_period += 1
                else:
                    continue

                station_cell = row.get(STATION_COLUMN)
                if station_cell is None:
                    continue
                is_shared, raw_station = station_cell
                station = _text(raw_station, is_shared, strings).strip()
                if not station:
                    continue
                totals = report.stations.setdefault(station, StationTotals(station))
                totals.swaps += 1
                totals.first_swap = earliest(totals.first_swap, swap_at)
                totals.last_swap = latest(totals.last_swap, swap_at)
                if NET_SOC_COLUMN in row:
                    totals.net_soc += _number(row[NET_SOC_COLUMN][1])
                if "Billable SOC" in row:
                    totals.billable_soc += _number(row["Billable SOC"][1])
                if "Swap Cost" in row:
                    totals.swap_cost += _number(row["Swap Cost"][1])
    return report


def load_station_names(workbook: Path | str, sheet_name: str | None = None) -> dict[str, str]:
    """Return normalized name to station name for a workbook's Station column."""
    try:
        from openpyxl import load_workbook
    except ImportError as error:  # pragma: no cover - dependency guard
        raise SwapDumpError("openpyxl is required; run `python -m pip install -e .`") from error

    book = load_workbook(workbook, read_only=True, data_only=True)
    try:
        sheet = book[sheet_name] if sheet_name else book[book.sheetnames[0]]
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, ())
        try:
            index = [str(name).strip() if name else "" for name in header].index(STATION_COLUMN)
        except ValueError as error:
            raise SwapDumpError(
                f"{Path(workbook).name} needs a '{STATION_COLUMN}' column"
            ) from error
        names: dict[str, str] = {}
        for row in rows:
            if len(row) <= index or row[index] is None:
                continue
            name = str(row[index]).strip()
            if name:
                names.setdefault(normalize_station(name), name)
        return names
    finally:
        book.close()


def match_stations(
    report: SwapReport, known: dict[str, str]
) -> list[tuple[StationTotals, str | None]]:
    """Pair each station in a report with its KPLC name, matched names first."""
    pairs = [
        (totals, known.get(normalize_station(totals.station)))
        for totals in report.stations.values()
    ]
    pairs.sort(key=lambda pair: (pair[1] is None, (pair[1] or pair[0].station).lower()))
    return pairs


def _row(
    periods: Sequence[str],
    matched: str,
    swap_station: str,
    swaps: int,
    net_soc: float,
    billable_soc: float,
    swap_cost: float,
    first_swap: datetime | None,
    last_swap: datetime | None,
) -> dict[str, object]:
    """Build one summary row in the shared output column order."""
    return {
        "Period": ", ".join(periods),
        "Matched Station Name": matched,
        "Swap Station": swap_station,
        "Swaps": swaps,
        "Net SOC": round(net_soc, 3),
        "Billable SOC": round(billable_soc, 3),
        "Swap Cost": round(swap_cost, 2),
        "Average Net SOC": round(net_soc / swaps, 3) if swaps else 0.0,
        "First Swap": _stamp(first_swap),
        "Last Swap": _stamp(last_swap),
    }


def summary_rows(reports: Sequence[SwapReport], known: dict[str, str]) -> list[dict[str, object]]:
    """Flatten every report into one output row per station per period."""
    rows: list[dict[str, object]] = []
    for report in reports:
        for totals, matched in match_stations(report, known):
            rows.append(
                _row(
                    periods=[report.period or ""],
                    matched=matched or "",
                    swap_station=totals.station,
                    swaps=totals.swaps,
                    net_soc=totals.net_soc,
                    billable_soc=totals.billable_soc,
                    swap_cost=totals.swap_cost,
                    first_swap=totals.first_swap,
                    last_swap=totals.last_swap,
                )
            )
    return rows


@dataclass
class _Accumulator:
    """Running totals for one station across every period being summarised."""

    periods: list[str]
    matched: str
    swap_station: str
    swaps: int
    net_soc: float
    billable_soc: float
    swap_cost: float
    first_swap: datetime | None
    last_swap: datetime | None

    @classmethod
    def of(cls, period: str, matched: str, totals: StationTotals) -> _Accumulator:
        """Start an accumulator from one station's totals in one period."""
        return cls(
            periods=[period],
            matched=matched,
            swap_station=totals.station,
            swaps=totals.swaps,
            net_soc=totals.net_soc,
            billable_soc=totals.billable_soc,
            swap_cost=totals.swap_cost,
            first_swap=totals.first_swap,
            last_swap=totals.last_swap,
        )

    def add(self, period: str, totals: StationTotals) -> None:
        """Fold another period's totals for the same station into this one."""
        if period not in self.periods:
            self.periods.append(period)
        self.swaps += totals.swaps
        self.net_soc += totals.net_soc
        self.billable_soc += totals.billable_soc
        self.swap_cost += totals.swap_cost
        self.first_swap = earliest(self.first_swap, totals.first_swap)
        self.last_swap = latest(self.last_swap, totals.last_swap)

    def as_row(self) -> dict[str, object]:
        """Render the accumulator as a summary row."""
        return _row(
            periods=self.periods,
            matched=self.matched,
            swap_station=self.swap_station,
            swaps=self.swaps,
            net_soc=self.net_soc,
            billable_soc=self.billable_soc,
            swap_cost=self.swap_cost,
            first_swap=self.first_swap,
            last_swap=self.last_swap,
        )


def combine_reports(
    reports: Sequence[SwapReport], known: dict[str, str]
) -> list[dict[str, object]]:
    """Total every period per matched station into single summary rows."""
    combined: dict[str, _Accumulator] = {}
    for report in reports:
        period = report.period or ""
        for totals, matched in match_stations(report, known):
            key = matched or totals.station
            accumulator = combined.get(key)
            if accumulator is None:
                combined[key] = _Accumulator.of(period, matched or "", totals)
            else:
                accumulator.add(period, totals)

    rows = [accumulator.as_row() for accumulator in combined.values()]
    rows.sort(
        key=lambda row: (not row["Matched Station Name"], str(row["Matched Station Name"]).lower())
    )
    return rows


def _stamp(value: object) -> str:
    """Format a datetime for Excel and CSV output."""
    return value.strftime("%Y-%m-%d %H:%M:%S") if isinstance(value, datetime) else ""


def write_summary(
    rows: Sequence[dict[str, object]],
    reports: Sequence[SwapReport],
    output: Path | str,
    known: dict[str, str] | None = None,
) -> Path:
    """Write the station summary, its exceptions, and an audit trail."""
    try:
        from openpyxl import Workbook
    except ImportError as error:  # pragma: no cover - dependency guard
        raise SwapDumpError("openpyxl is required; run `python -m pip install -e .`") from error

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook(write_only=True)
    sheet = book.create_sheet("Station Summary")
    sheet.append(list(SUMMARY_HEADERS))
    for row in rows:
        sheet.append([row.get(header) for header in SUMMARY_HEADERS])

    unmatched = book.create_sheet("Unmatched Swap Stations")
    unmatched.append(["Period", "Swap Station", "Swaps", "Net SOC"])
    for row in rows:
        if not row["Matched Station Name"]:
            unmatched.append([row["Period"], row["Swap Station"], row["Swaps"], row["Net SOC"]])

    if known is not None:
        seen = {str(row["Matched Station Name"]) for row in rows if row["Matched Station Name"]}
        idle = book.create_sheet("KPLC Stations Without Swaps")
        idle.append(["Matched Station Name"])
        for name in sorted(known.values(), key=str.lower):
            if name not in seen:
                idle.append([name])

    audit = book.create_sheet("Metadata")
    audit.append(["Report", "Value"])
    for report in reports:
        for label, value in (
            ("Source", str(report.source)),
            ("Worksheet", report.sheet),
            ("Period", report.period or "all dates"),
            ("Period start", _stamp(report.period_start)),
            ("Period end (exclusive)", _stamp(report.period_end)),
            ("Rows scanned", report.rows_scanned),
            ("Rows in period", report.rows_in_period),
            ("Rows outside period", report.rows_outside_period),
            ("Stations", len(report.stations)),
            ("Swaps", report.swaps),
            ("Net SOC", round(report.net_soc, 3)),
            ("Billable SOC", round(report.billable_soc, 3)),
            ("Swap Cost", round(report.swap_cost, 2)),
            ("Observed first swap", _stamp(report.observed_start)),
            ("Observed last swap", _stamp(report.observed_end)),
        ):
            audit.append([f"{report.source.name} {label}", value])
    book.save(output)
    return output


def write_csv(rows: Sequence[dict[str, object]], output: Path | str) -> Path:
    """Write the station summary as UTF-8 CSV, ready for database loading."""
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(SUMMARY_HEADERS))
        writer.writeheader()
        for row in rows:
            writer.writerow({header: row.get(header, "") for header in SUMMARY_HEADERS})
    return output


def discover_dumps(inputs: Iterable[str]) -> list[Path]:
    """Expand files, directories, and globs into a sorted list of workbooks."""
    paths: set[Path] = set()
    for value in inputs:
        path = Path(value)
        if path.is_dir():
            paths.update(item for item in path.glob("*.xlsx") if item.is_file())
        elif any(char in value for char in "*?[]"):
            paths.update(
                Path(item)
                for item in sorted(glob.glob(value, recursive=True))
                if Path(item).is_file()
            )
        elif path.is_file() and path.suffix.lower() == ".xlsx":
            paths.add(path)
    return sorted(paths)
