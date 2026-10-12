# KPLC Power Bill Renamer

Two tools:

- `kplc-rename` reads account numbers off scanned KPLC bill PDFs with OCR,
  looks the account up in an Excel workbook, and renames each PDF to
  `<station-name>.pdf`.
- `kplc-swap` sums Net SOC per station in the monthly swap dumps and matches
  the result against the same station list.

## Bill renaming

Finds the first table in each KPLC bill that contains an account number, reads
the account number with OCR, looks it up in the `September
` sheet of the Excel
workbook, and renames the PDF to `<station-name>.pdf`. If the account number is not in the workbook, it uses `<account-number>.pdf`.

The supplied PDFs are image-only scans, so the project uses Tesseract OCR.

## Setup

Install the system OCR and Python dependencies:

```bash
sudo apt install tesseract-ocr
python -m pip install -e .
```

The first command installs the Tesseract OCR engine used to read scanned
account numbers. The second installs this project and its Python dependencies
in editable mode, so changes to the source are reflected immediately.

## Bill usage

Preview planned renames and capture bill fields without renaming PDFs:

```bash
python -m powerbills
```

Apply the renames and capture bill fields:

```bash
python -m powerbills --apply
```

Extract bill fields without planning or applying PDF renames:

```bash
python -m powerbills --extract-only
```

To extract only PDFs in a specific folder:

```bash
python -m powerbills --extract-only raw_files/Harold
```

The command always uses `raw_files` when no input is supplied. The `raw-2`
directory is a backup/test directory and is never scanned automatically. To
run a manual test against it, pass it explicitly: `python -m powerbills raw-2`.

By default the command reads `excel/Power Bills Template.xlsx` and its
`September` sheet. The requested sheet name can be overridden with
`--lookup-sheet` and `--output-sheet`; if the workbook has only one sheet, that
sheet is used automatically when the requested name is absent.

The extract-only command skips station lookup and all rename processing. It
writes captured bill data to `excel/September KPLC Power Bills.xlsx` by default.
Captured fields include `consumption` from the bill's Fuel Energy Cost line,
the bill total, consumption period, low- and high-rate consumption, classified
method of charge, and due date. Use `--output` to choose another output
workbook. The default run previews planned renames; pass
`--apply` to perform them.

You can also pass individual PDFs or a glob. Existing destination files are
never overwritten; use `--on-conflict skip` (the default) or `--on-conflict
error` to stop at the first conflict.

## Swap dump summary

```bash
kplc-swap
```

Defaults to `Swap Dump/*.xlsx`, writing `excel/Swap Net SOC Summary.xlsx` and
`excel/Swap Net SOC Summary.csv`.

Each dump is one month of battery swaps, roughly 900k rows whose worksheet XML
expands to about 800 MB inside the .xlsx container. `kplc-swap` never loads the
worksheet: it streams the sheet part out of the zip container in 8 MB chunks,
decodes only `Swap DateTime`, `Station`, `Net SOC`, `Billable SOC`, and
`Swap Cost`, and keeps one running total per station. A 350 MB pair of dumps
takes about 80 seconds in a few hundred MB of RAM, against roughly 4 minutes
per month with `openpyxl`.

The month comes from the filename, so `August_Swap_Data.xlsx` is summarised as
August. Override it when a filename is unhelpful:

```bash
kplc-swap --period 2026-08 --year 2026 "Swap Dump/August_Swap_Data.xlsx"
```

Station names are matched to the `Station` column of the workbook template,
case- and punctuation-insensitively, and reported as `Matched Station Name`.
Swap stations with no match are listed in `Unmatched Swap Stations`, and KPLC
stations that recorded no swaps in `KPLC Stations Without Swaps`. Pass
`--no-template` to skip matching, `--template` to point at a different station
list, and `--csv` to redirect the CSV.

The workbook has four sheets:

| Sheet | Contents |
| --- | --- |
| `Station Summary` | One row per matched station: swaps, Net SOC, Billable SOC, Swap Cost, average Net SOC, first and last swap |
| `Unmatched Swap Stations` | Swap stations absent from the station list, so nothing is silently dropped |
| `KPLC Stations Without Swaps` | Station-list entries with no swaps in the period |
| `Metadata` | Source file, worksheet, period bounds, rows scanned, rows outside the period, and totals |

Check `Metadata` before trusting a total: `Rows outside period` non-zero means
the dump held dates the filename did not describe, and `Observed last swap`
earlier than the period end means the month is only partly dumped.

## Development

```bash
python -m pip install -e .
python -m pip install pytest ruff
python -m pytest
python -m ruff check src tests
```

## Why not Postgres for the swap dumps

The aggregation is a single group-by over two columns per row, and streaming it
locally takes 80 seconds with no schema, no COPY, and no egress cost. A Neon
table earns its keep once the raw swaps need to be queried, joined, or joined
again by other reports. The CSV this tool writes is already the right shape to
`COPY` into a `station_period_totals` table when that day comes.
