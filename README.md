# SubPay Review — Processing Phase

SubPay Review is a subcontractor pay-application review pipeline. This repository
contains the backend processing module that ingests a single uploaded file and
returns a normalized list of review records.

## Product / Archetype

Archetype: **document / pay-application review**. The pipeline ingests a
subcontractor pay application (AIA G702/G703 style forms, waivers, COIs, W-9s,
lien waivers, etc.), extracts structured fields from mixed file formats, derives
a review status, and emits one record per tracked entity for downstream review
and flagging.

## Files

- `processor.py` — entry point is `process_file(file_bytes: bytes) -> list[dict]`
- `run_demo.py` — hardcoded CSV smoke test
- `run_tests.py` — unit tests for `process_file()`
- `requirements.txt` — Python dependencies
- `README.md` — this file

## Record Shape

`process_file()` always returns a `list[dict]`. Each record contains:

- `title` — primary tracked entity, typically the vendor/subcontractor name
- `status` — one of the allowed status strings (`Name:severity`, e.g. `Pending_review:info`)
- `details` — dict of extracted fields (never contains a nested `due_date`)
- `due_date` — top-level ISO-8601 date string, or empty string if none found

Every record's `status` is guaranteed to be present in `processor.ALLOWED_STATUSES`.

## Supported Inputs

`process_file()` accepts raw uploaded bytes and tries in order:

1. PDF text extraction with `pdfplumber`
2. Excel (`.xlsx`) extraction with `openpyxl`
3. UTF-8 text / CSV fallback

It never rejects an unknown format — it always falls back to text/CSV decoding.

## What the Poller Expects as Input

The poller hands `process_file()` the **raw bytes of a single uploaded file**
exactly as received from storage — no filename, extension, or content-type is
required. The processor sniffs the format itself. Expected inputs include:

- AIA G702/G703 pay application PDFs
- Excel schedules of values and pay-app workbooks
- CSV exports of pay-app line items
- Plain-text field dumps (`key: value` or comma-separated lines)

## Usage

```python
from processor import process_file

records = process_file(open("pay_app.xlsx", "rb").read())
for record in records:
    print(record["title"], record["status"], record["due_date"])
```

## Running Locally

```bash
pip install -r requirements.txt
python3 run_demo.py
python3 run_tests.py
```

`run_demo.py` runs a hardcoded CSV smoke test and exits 0 on success.
`run_tests.py` runs the unit test suite via `unittest`.
Railway: subcontractorpayappreview
