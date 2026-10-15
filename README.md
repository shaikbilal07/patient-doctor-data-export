# PostgreSQL to Excel Exporter

A Python program that joins the `patient_details` and `doctor` tables in PostgreSQL and exports the result to a formatted Excel file (`.xlsx`) or a CSV file (`.csv`, `.csv.gz`).

## What it does

- Connects to PostgreSQL with psycopg2 in read-only mode. The password is read from the `PGPASSWORD` environment variable and is never written to logs.
- Runs the query in `patient_doctor_query.sql`: a `LEFT JOIN` on `patient_details.doctor_id = doctor.doctor_id`, so patients without a doctor are kept.
- Streams rows with a server-side cursor and `fetchmany()`, so the full result is never loaded into memory.
- Writes a readable Excel file:
  - only the query's columns, with clear headers;
  - frozen header row and filters;
  - formatted dates and numbers;
  - column widths limited, with long text wrapped.
- Checks the free disk space and Excel's limits before starting.
- Writes to `<file>.part` first, checks that the written row count equals the query's row count, and only then renames the file.

## Excel limits and 2 TB of data

An Excel worksheet holds at most 1,048,576 rows, including the header. Larger results are split across worksheets and workbooks. Results above 5,000,000 rows (`--max-xlsx-rows`) are refused, because Excel cannot open them in practice.

A 2 TB result is billions of rows and cannot be stored in Excel. Export it to `.csv.gz` instead, which has no row limit and keeps every value exactly.

An interrupted export cannot be resumed. The partial file is deleted, and the export must be started again.

## Requirements

- Python 3.12 or newer
- PostgreSQL

## Setup (Windows PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:PGPASSWORD = "YOUR_PASSWORD"
```

## Usage

```powershell
python export_to_excel.py --output exports\patients_doctors.xlsx
python export_to_excel.py --output exports\patients_doctors.csv.gz
python export_to_excel.py --output exports\sample.xlsx --limit 1000
python export_to_excel.py --dbname YOUR_DATABASE --output D:\export\patients_doctors.csv.gz --log-file export.log
```

Default connection: host `127.0.0.1`, port `5432`, database `hospital_export_test`, user `postgres`. You can change these with `--host`, `--port`, `--dbname` and `--user`. Run `python export_to_excel.py --help` to see all options.

## Tests

The tests run against a real PostgreSQL database and need `PGPASSWORD`. The edge-case tests create a temporary schema and drop it again afterwards. The patient and doctor tables are only read.

```powershell
$env:PGPASSWORD = "YOUR_PASSWORD"
python -m pytest -v
```

| Variable | Default | Purpose |
|---|---|---|
| `EXPORT_TEST_HOST` | `127.0.0.1` | database host |
| `EXPORT_TEST_PORT` | `5432` | database port |
| `EXPORT_TEST_DB` | `hospital_export_test` | database with `patient_details` and `doctor` |
| `EXPORT_TEST_USER` | `postgres` | database user |
| `EXPORT_SHOWCASE_DB` | `loadtest` | database with the optional 50,000-row `showcase_records` table; the test is skipped if that table is missing |

## Files

| File | Purpose |
|---|---|
| `export_to_excel.py` | The exporter |
| `patient_doctor_query.sql` | The join query (one `SELECT`, no `LIMIT`) |
| `test_export_to_excel.py` | Tests |
| `requirements.txt` | Python dependencies |
