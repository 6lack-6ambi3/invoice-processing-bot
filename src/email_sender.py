"""
Send HTML email summaries of invoice batch runs via Gmail SMTP.

Reads credentials from .env (GMAIL_ADDRESS, GMAIL_APP_PASSWORD, RECIPIENT_EMAIL).
and sends a clean HTML summary with the Excel master attached.
"""

from __future__ import annotations
import logging
from operator import inv
import os
import smtplib
import ssl
from datetime import datetime
from decimal import Decimal
from email.message import EmailMessage
from pathlib import Path
from typing import Sequence
from unicodedata import name
from dotenv import load_dotenv
from src.extractor import Invoice

load_dotenv()  # Load environment variables from .env
logger = logging.getLogger(__name__)

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465

# ----HTML Rendering
def _render_html(successful: Sequence[Invoice],
                 failed: Sequence[tuple[str, str]],) -> str:
    """Build a clean HTML email body summarizing the batch."""
    total_value = sum((inv.total_amount for inv in successful), start=Decimal("0"))
    currency = successful[0].currency if successful else "USD"

    success_rows = "\n".join(
        f"""
        <tr>
            <td>{inv.invoice_number}</td>
            <td>{inv.vendor_name or 'N/A'}</td>
            <td>{inv.invoice_date.isoformat() if inv.invoice_date else "-"}</td>
            <td style="text-align:right;font-variant-numeric:tabular-nums;">{inv.total_amount:,.2f} {currency}</td>
            <td style="color:#888;font-size:12px;">{inv.extraction_method}</td>
        </tr>
        """
        for inv in successful
    )

    failed_section = ""
    if failed:
        failed_rows = "\n".join(
            f"<tr><td>{name}</td><td style='color:#c0392b;'>{err}</td></tr>"
            for name, err in failed
        )
        failed_section = f"""
        <h3 style="color:#c0392b;margin-top:24px;">Failed Extractions ({len(failed)})</h3>
        <table style="border-collapse:collapse;width:100%;font-size:13px;">
            <thead>
                <tr>
                    <th style="border:1px solid #ccc;padding:8px;">Invoice Number</th>
                    <th style="border:1px solid #ccc;padding:8px;">Error</th>
                </tr>
            </thead>
            <tbody>
                {failed_rows}
            </tbody>
        </table>
        """

    return f"""
    <html>
        <body style="font-family:-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        color:#333; max-width:680px;margin:0 auto;padding:24px;">
            <h2 style="color:#1D9E75;margin:0 0 4px;">Invoice batch processed</h2>
            <p style="color:#666; margin: 0 0 24px;font-size:14px;">{datetime.now().strftime("%A, %d %B %Y at %H:%M")}</p>

            <div style="display:grid;grid-template-columns:repeat(3, 1fr);gap:12px;margin-bottom:24px;">
                <div style="background:#f0f4f8;padding:16px;border-radius:8px;">
                    <div style="font-size:22px;font-weight:500;">{len(successful)}</div>
                    <div style="font-size:12px;color:#888;">Processed</div>
                </div>
                <div style="background:#f4f4f0;padding:14px;border-radius:8px;">
                    <div style="font-size:22px;font-weight:500;">{currency} {total_value:,.2f}</div>
                    <div style="font-size:12px;color:#888;">Total Value</div>
                </div>
                <div style="background:#f4f4f0;padding:14px;border-radius:8px;">
                    <div style="font-size:22px;font-weight:500;color:{'#c0392b' if failed else '#1D9E75'};">
                        {len(failed)}
                    </div>
                    <div style="font-size:12px;color:#888;">Failed</div>
                </div>
            </div>

            <h3 style="margin:0 0 8px;">Successfully processed ({len(successful)})</h3>
            <table style="border-collapse:collapse;width:100%;font-size:13px;">
                <thead>
                    <tr style="background:#e1f5ee;">
                        <th style="padding:8px;text-align:left;">Invoice #</th>
                        <th style="padding:8px;text-align:left;">Date</th>
                        <th style="padding:8px;text-align:left;">Vendor</th>
                        <th style="padding:8px;text-align:left;">Total</th>
                        <th style="padding:8px;text-align:left;">Method</th>
                    </tr>
                </thead>
                <tbody>{success_rows}</tbody>
            </table>

            {failed_section}

            <p style="color:#888;font-size:12px;margin-top:32px; border-top:1px solid #eee;padding-top:12px;">
            Sent automatically by the invoice processing bot.
            The full master spreadsheet is attached for reference.
            </p>
        </body>
    </html>"""

def _render_plaintext(successful: Sequence[Invoice], failed: Sequence[tuple[str, str]],) -> str:
    """Plaintext fallback for email clients that do not render HTML."""
    total_value = sum((inv.total_amount for inv in successful), start=Decimal("0"))
    currency = successful[0].currency if successful else "USD"

    lines = [
        "Invoice batch processed",
        f"Run at {datetime.now().strftime('%Y-%m-%d at %H:%M')}",
        "",
        f"Successfully processed: {len(successful)}",
        f"Total value: {currency} {total_value:,.2f}",
        f"Failed extractions: {len(failed)}",
        "",
        "Details of successful extractions:",
    ]
    for inv in successful:
        date = inv.invoice_date.isoformat() if inv.invoice_date else "-"
        lines.append(f" - Invoice #{inv.invoice_number} ({date}): {currency} {inv.total_amount:,.2f}")

    if failed:
        lines += ["", "Failed Extractions:"]
        for name, err in failed:
            lines.append(f" - {name}: {err}")
    return "\n".join(lines)

def send_summary_email(successful: Sequence[Invoice], failed: Sequence[tuple[str, str]], excel_path: Path) -> bool:
    """Send an HTML summary email with the master Excel attachment."""
    sender = os.getenv("GMAIL_ADDRESS")
    password = os.getenv("GMAIL_APP_PASSWORD")
    recipient = os.getenv("RECIPIENT_EMAIL", sender)

    if not sender or not password:
        logging.error("Missing Gmail credentials in environment variables.")
        return False
    
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = (f"Invoice Batch Processed - {len(successful)} successful" + (f", {len(failed)} failed" if failed else ""))

    msg.set_content(_render_plaintext(successful, failed), subtype="plain")
    msg.add_alternative(_render_html(successful, failed), subtype="html")

    excel_path = Path(excel_path)
    if excel_path.is_file():
        with open(excel_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=excel_path.name,)
        
    else:
        logger.warning(f"Excel file not found at {excel_path}, sending email without attachment.")

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
            server.login(sender, password)
            server.send_message(msg)
        logger.info(f"Summary email sent to {recipient}.")
        return True
    except smtplib.SMTPAuthenticationError:
        logger.error("Gmail auth failed. Check GMAIL_APP_PASSWORD - must be a 16-char"
                     "App Password, not your real password.")
        
        return False
    
    except Exception as e:
        logger.error("Email send failed: %s", e)
        return False
    
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    fake = Invoice(
        invoice_number="TEST_001",
        invoice_date=datetime.now().date(),
        vendor_name="Smoke Test Vendor Ltd",
        total_amount=Decimal("1234.56"),
        currency="EUR",
        line_items=[],
        source_file='smoke_test.pdf',
        extraction_method="manual",
    )
    fake_failed = [("TEST_002", "OCR failed - low confidence"), ("TEST_003", "File not found")]
    excel_path = Path(__file__).parent.parent / "output" / "master.xlsx"
    ok = send_summary_email([fake], fake_failed, excel_path)
    print("Email send result:", "Success" if ok else "Failure")