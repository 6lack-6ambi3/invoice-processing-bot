"""
Append parsed invoices to a Google Sheet.

Uses a Google Cloud service account for headless authentication - 
no browser-based OAuth flow, perfect for scheduled bots.
"""

from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Iterable

import gspread
from gspread.utils import ValueInputOption
from google.oauth2.service_account import Credentials

from src.extractor import Invoice

logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive",
          ]

HEADERS = [
    "Invoice #", "Date", "Vendor", "Total Amount", "Currency",
    "Line Items", "Source File", "Extraction Method", "Processed Timestamp"
]

def get_client() -> gspread.Client:
    """Authenticate with Google Sheets API using service account credentials."""
    creds_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "gcp-credentials.json")
    creds_path = Path(creds_path)

    if not creds_path.is_absolute():
        creds_path = Path(__file__).parent.parent / creds_path

    if not creds_path.exists():
        raise FileNotFoundError(
            f"Google credentials not found at {creds_path}. "
            f"Download a service account JSON from Google Cloud Console."
        )
    
    creds = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)
    return gspread.authorize(creds)

def _ensure_headers(worksheet: gspread.Worksheet) -> None:
    """Write the header row if the sheet is empty."""
    first_row = worksheet.row_values(1)
    if not first_row:
        worksheet.append_row(HEADERS, value_input_option=ValueInputOption.user_entered)
        worksheet.format("A1:I1", {
            "textFormat": {"bold": True, "foregroundColor": 
                           {"red": 1, "green": 1, "blue": 1}},
                           "backgroundColor": {"red": 0.114, "green": 0.62, "blue": 0.46},
                           "horizontalAlignment": "CENTER",})
        worksheet.freeze(rows=1)
        logger.info("Initialized worksheet with headers in Google Sheet.")

def append_invoices_to_sheet(invoices: Iterable[Invoice], sheet_name: str | None = None) -> int:
    """
    Append each invoice as a new row in the configured Google Sheet.
    Returns the number of rows/invoices appended.
    """

    sheet_name = sheet_name or os.getenv("GOOGLE_SHEET_NAME", "Invoice Master")

    client = get_client()
    try:
        spreadsheet = client.open(sheet_name)
    except gspread.SpreadsheetNotFound:
        raise RuntimeError(
            f"Sheet '{sheet_name}' not found. Make sure it exists and is "
            f"shared with your service account email."
        )
    
    worksheet = spreadsheet.sheet1
    _ensure_headers(worksheet)

    rows = []
    for inv in invoices:
        rows.append([
            inv.invoice_number,
            inv.invoice_date.isoformat() if inv.invoice_date else "",
            inv.vendor_name or "",
            float(inv.total_amount),
            inv.currency,
            "; ".join(f"{li.description}: {li.amount}" for li in inv.line_items),
            inv.source_file,
            inv.extraction_method,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ])
        logger.info("Appended to Google Sheets: %s - %s %s",
                    inv.invoice_number, inv.total_amount, inv.currency)

    if rows:
        worksheet.append_rows(rows, value_input_option=ValueInputOption.user_entered)

    return len(rows)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    from dotenv import load_dotenv
    load_dotenv()

    from src.extractor import parse_invoice
    inbox = Path(__file__).parent.parent / "invoices" / "inbox"
    pdfs = sorted(inbox.glob("*.pdf"))

    if not pdfs:
        inbox = Path(__file__).parent.parent / "invoices" / "processed"
        pdfs = sorted(inbox.glob("*.pdf"))

    invoices = [parse_invoice(p) for p in pdfs[:3]]
    n = append_invoices_to_sheet(invoices)
    print(f"Appended {n} rows/invoices to Google Sheet.")