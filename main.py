"""
Invoice processing bot - main orchestrator.

Workflow on each run:
1. Scan invoices/inbox/ for PDF files.
2. Extract structured data from each (pdfplumber -> OCF fallback)
3. Append results to master spreadsheet (Google Sheets)
4. Move processed PDFs to invoices/processed/
5. Send a summary email with the Excel attached (SMTP)

Designed to be run on a schedule (e.g. via cron / LaunchAgent) or triggered by new files in the inbox directory - idempotent and safe to invoke repeatedly without side effects.
"""

from __future__ import annotations
import logging
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

from src.excel_writer import append_invoices_to_excel, move_processed_pdf
from src.extractor import parse_invoice
from src.email_sender import send_summary_email

from src.sheets_writer import append_invoices_to_sheet

#-------PATHS & CONFIG-------
PROJECT_ROOT = Path(__file__).parent
INBOX_DIR = PROJECT_ROOT / "invoices" / "inbox"
PROCESSED_DIR = PROJECT_ROOT / "invoices" / "processed"
OUTPUT_DIR = PROJECT_ROOT / "output" / "master.xlsx"
LOG_FILE = PROJECT_ROOT / "logs" / "invoice_bot.log"

#-------LOGGING SETUP-------
def setup_logging() -> None:
    """Configure both console and rotating-file logging."""
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    file_handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=5, encoding="utf-8")
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.handlers.clear()  # Remove any default handlers
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

#-------MAIN FUNCTION-------
def run() -> int:
    """Excute one full batch. Returns process exit code."""
    log = logging.getLogger("bot")
    start_time = datetime.now()
    log.info("=" * 60)
    log.info("Starting invoice processing run")

    pdfs = sorted(INBOX_DIR.glob("*.pdf"))
    if not pdfs:
        log.info("Nothing to process - no PDFs found in %s", INBOX_DIR)
        return 0
    log.info("Found %d PDF(s) to process", len(pdfs))

    invoices = []
    successful_pdfs = []
    failed: list[tuple[str, str]] = []

    for pdf in pdfs:
        try:
            invoice = parse_invoice(pdf)
            invoices.append(invoice)
            successful_pdfs.append(pdf)
            log.info("Parsed %s - %s %s", pdf.name, invoice.total_amount, invoice.currency)
        except Exception as e:
            failed.append((pdf.name, str(e)))
            log.error("Failed %s - %s", pdf.name, e)

    if invoices:
        try:
            n = append_invoices_to_excel(invoices, OUTPUT_DIR)
            log.info("Appended %d invoice(s) to %s", n, OUTPUT_DIR)
            n = append_invoices_to_sheet(invoices)
            log.info("Wrote %d row(s)/invoice(s) to Google Sheet", n)
        except Exception as e:
            log.error("Failed to append invoices to Excel - %s", e)
            log.exception("Google Sheets write failed: %s - local Excel still saved", e)
            return 2
    else:
        log.warning("No succcessful extractions - skipping Excel write")

    if successful_pdfs:
        try:
            move_processed_pdf(successful_pdfs, PROCESSED_DIR)
            log.info("Moved %d PDF(s) to processed/", len(successful_pdfs))
        except Exception as e:
            log.exception("File moved failed: %s", e)

    try:
        sent = send_summary_email(invoices, failed, OUTPUT_DIR)
        if sent:
            log.info("Summary email sent successfully.")
        else:
            log.warning("Summary email was not sent (check .env credentials).")
    except Exception as e:
        log.exception("Email sending failed: %s", e)

    elapsed = (datetime.now() - start_time).total_seconds()
    log.info("Run completed - %d ok, %d failed, in %.1f seconds", len(invoices), len(failed), elapsed)
    log.info("=" * 60)

    return 1 if failed and not invoices else 0

if __name__ == "__main__":
    setup_logging()
    raise SystemExit(run())