# Invoice Processing Bot

Automated bot that extracts structured data from PDF invoices, writes it to a master Excel file and a Google Sheet, then emails a batch summary.

Built as part of an Applied AI portfolio project — combines OCR (computer vision applied to documents) with classical text parsing, data validation, and end-to-end automation.

## Features

- Reads PDF invoices from a watched folder
- Dual extraction strategy: `pdfplumber` for native PDFs, Tesseract OCR fallback for scanned ones
- Pulls structured fields (invoice number, date, vendor, total, currency) via regex heuristics
- Validates and normalises via Pydantic
- Appends to a styled local Excel master and a live Google Sheet
- Sends HTML email summary with the Excel attached
- Structured logging with rotation
- Designed to run on schedule (LaunchAgent / cron)

## Tech stack

- `pdfplumber` + `pytesseract` + `pdf2image` — PDF and OCR extraction
- `pydantic` — data validation
- `openpyxl` — local Excel output
- `gspread` + `google-auth` — Google Sheets sync
- `smtplib` + Gmail App Password — email confirmations

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Install system dependencies (macOS)
brew install tesseract poppler

# Configure
cp .env.example .env
# Fill in GMAIL_ADDRESS, GMAIL_APP_PASSWORD, RECIPIENT_EMAIL
# Add gcp-credentials.json from Google Cloud Console
# Add GOOGLE_SHEET_ID from your Sheet's URL
```

## Usage

Drop PDFs into `invoices/inbox/` and run:

```bash
python main.py
```

The bot will extract data, write to Excel + Google Sheets, move processed files, and email a summary.

## Architecture
```
main.py                 ← orchestrator
src/extractor.py        ← PDF → Pydantic Invoice model
src/excel_writer.py     ← local master.xlsx
src/sheets_writer.py    ← live Google Sheet sync
src/email_sender.py     ← HTML batch summary email
```
## What I'd improve

- Extract line items as a structured list (currently empty in v1)
- Add an LLM fallback for vendor name extraction on unusual invoice layouts
- Support multi-currency totals in the email summary
- Add a web dashboard for browsing processed invoices

## License

MIT
