"""
Invoice extraction with pdfplumber primary + Tesseract OCR fallback.
Strategy:
1. Try to extract text with pdfplumber - fast and accurate for native PDFs.
2. If extracted text is empty or too short, fallback to OCR.
3. Run regex heuristics on the text to pull out structured fields.
4. Validate and normalise via the Pydantic Invoice model.
"""

from __future__ import annotations
from email.mime import text
import logging
import re
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Optional
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from pydantic import BaseModel, Field, field_validator

logger = logging.getLogger(__name__)

MIN_NATIVE_TEXT_CHARS = 50

# ----- Pydantic Data Model -----
class LineItem(BaseModel):
    description: str
    amount: Decimal

class Invoice(BaseModel):
    """
    Validate and normalised invoice records
    """
    invoice_number: str = Field(..., min_length=1)
    invoice_date: Optional[date] = None
    vendor_name : Optional[str] = None
    total_amount: Decimal = Field(..., ge=0)
    currency: str = "USD"
    line_items: list[LineItem] = Field(default_factory=list)

    #BookKeeping
    source_file: str
    extraction_method: str

    @field_validator("vendor_name")
    @classmethod
    def strip_vendor(cls, v: Optional[str]) -> Optional[str]:
        return v.strip() if v else v
    
# ----- Extraction Logic -----
def extract_invoice_data(pdf_path: Path) -> str:
    """
    Native PDF text extraction. Fast but only works if PDF has a text layer.
    """
    text_parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)
    return "\n".join(text_parts).strip()

def ocr_extract_invoice_data(pdf_path: Path, dpi: int = 300) -> str:
    """
    OCR fallback for scanned PDFs. Slower but works on any image-based PDF.
    """
    images = convert_from_path(str(pdf_path), dpi=dpi)
    text_parts = [pytesseract.image_to_string(img) for img in images]
    return "\n".join(text_parts).strip()

def extract_text(pdf_path: Path) -> tuple[str, str]:
    """
    Try pdfplumber first; fallback to OCR if text is missing/too short.
    Returns (text, method_used).
    """
    text = extract_invoice_data(pdf_path)
    if len(text) >= MIN_NATIVE_TEXT_CHARS:
        logger.info("Extracted %d chars with pdfplumber from %s", len(text), pdf_path.name)
    
    return text, "pdfplumber"
    
    logger.info("pdfplumber returned only %d chars for %s - falling back to ORC", len(text), pdf_path.name,)
    text = ocr_extract_invoice_data(pdf_path)
    return text, "ocr"

# ----- Field Extraction Heuristics -----
INVOICE_NUMBER_PATTERNS = [
    # "Invoice Number: 4HirpOI" or "Invoice #: 1234"
    r"invoice\s*(?:number|no\.?|#)\s*[:\-]?\s*([A-Z0-9][A-Z0-9\-]*)",
    # "INV-001" or "INV: 12345"
    r"\binv\s*[:\-#]?\s*([A-Z0-9][A-Z0-9\-]*)",
    # "# 1" but only when standalone (not "Date: May 4")
    r"^#\s*([A-Z0-9][A-Z0-9\-]*)\s*$",
]

DATE_PATTERNS = [
    #ISO: YYYY-MM-DD
    (r"(\d{4})-(\d{1,2})-(\d{1,2})", "ymd"),
    #US: MM/DD/YYYY or MM-DD-YYYY
    (r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", "mdy"),
    #Long: January 1, 2020
    (r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", "dmy_long"),
    (r"([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})", "mdy_long"),
]

TOTAL_PATTERNS = [
    # "Total: $376.00" or "Total $376.00" or "Total 376.00"
    r"(?<!sub)(?<!ship)total\s*(?:due|amount)?\s*[:\-]?\s*[$£€]?\s*([\d,]+\.\d{2})",
    r"amount\s*due\s*[:\-]?\s*[$£€]?\s*([\d,]+\.\d{2})",
    r"grand\s*total\s*[:\-]?\s*[$£€]?\s*([\d,]+\.\d{2})",
    r"balance\s*due\s*[:\-]?\s*[$£€]?\s*([\d,]+\.\d{2})",
]

CURRENCY_PATTERNS = {"$": "USD", "£": "GBP", "€": "EUR"}

MONTHS = {m : i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}

def extract_invoice_number(text: str) -> Optional[str]:
    for pattern in INVOICE_NUMBER_PATTERNS:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            return match.group(1).strip()
    return None

def find_date(text: str) -> Optional[date]:
    """
    Try multiple date formats; return the first valid match.
    """
    for pattern, fmt in DATE_PATTERNS:
        for match in re.finditer(pattern, text):
            try:
                g = match.groups()
                if fmt == "ymd":
                    y, m, d = int(g[0]), int(g[1]), int(g[2])
                elif fmt == "mdy":
                    m, d, y = int(g[0]), int(g[1]), int(g[2])
                elif fmt == "dmy_long":
                    d, m, y = int(g[0]), MONTHS[g[1][:3].lower()], int(g[2])
                elif fmt == "mdy_long":
                    m, d, y = MONTHS[g[0][:3].lower()], int(g[1]), int(g[2])
                if 2000 <= y <= 2100 and 1 <= m <= 12 and 1 <= d <= 31:
                    return date(y, m, d)
            except (ValueError, KeyError) as e:
                continue
    return None

def extract_total_amount(text: str) -> tuple[Optional[Decimal], str]:
    """Returns (amount, currency)"""
    currency = "USD"
    for symbol, code in CURRENCY_PATTERNS.items():
        if symbol in text:
            currency = code
            break

    for pattern in TOTAL_PATTERNS:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            cleaned = match.group(1).replace(",", "")
            try:
                return Decimal(cleaned), currency
            except InvalidOperation:
                continue
    return None, currency

def find_vendor_name(text: str) -> Optional[str]:
    """
    Heuristic: Look for lines starting with "From:", "Vendor:", "Supplier:" etc.
    """
    skip_words = {"invoice", "bill", "receipt", "statement", "tax invoice"}
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    for line in lines[:5]:
        cleaned = re.sub(r"\b(invoice|bill|receipt|statement)\b", "", line, flags=re.IGNORECASE).strip()
        low = cleaned.lower()

        if not cleaned or len(cleaned) < 3:
            continue
        if cleaned.startswith("#") or re.match(r"^\d", cleaned):
            continue
        if low in skip_words:
            continue
        return cleaned[:60]
    return None

# ----- Main Extraction Function -----
def parse_invoice(pdf_path: Path) -> Invoice:
    """
    Extract text from a PDF, parse fields and return a validated Invoice.

    Raises:
        ValueError: if required fields are missing or invalid.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.is_file():
        raise ValueError(f"File not found: {pdf_path}")
    
    text, method = extract_text(pdf_path)
    if not text:
        raise ValueError(f"No text extracted from {pdf_path.name} using {method}")
    
    invoice_num = extract_invoice_number(text)
    invoice_date = find_date(text)
    vendor = find_vendor_name(text)
    total_amount, currency = extract_total_amount(text)

    if not invoice_num:
        raise ValueError(f"Invoice number not found in {pdf_path.name}")
    if total_amount is None:
        raise ValueError(f"Total amount not found in {pdf_path.name}")
    
    return Invoice(
        invoice_number=invoice_num,
        invoice_date=invoice_date,
        vendor_name=vendor,
        total_amount=total_amount,
        currency=currency,
        line_items=[],  # Line item extraction not implemented in this version
        source_file=str(pdf_path.name),
        extraction_method=method,
    )

#----CLI Interface for Testing ----
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    inbox = Path(__file__).parent.parent / "invoices" / "inbox"
    pdfs = sorted(inbox.glob("*.pdf"))

    if not pdfs:
        print(f"No PDF files found in {inbox}. Please add some test invoices.")
        raise SystemExit(0)
    
    print(f"Found {len(pdfs)} PDF(s) in {inbox}. Extracting data...\n")
    for pdf in pdfs:
        try:
            invoice = parse_invoice(pdf)
            print(f"Extracted from {pdf.name}:")
            print(f"  Invoice Number: {invoice.invoice_number}")
            print(f"  Invoice Date: {invoice.invoice_date}")
            print(f"  Vendor Name: {invoice.vendor_name}")
            print(f"  Total Amount: {invoice.total_amount} {invoice.currency}")
            print(f"  Extraction Method: {invoice.extraction_method}\n")
        except Exception as e:
            print(f"Error processing {pdf.name}:\n ERROR: {e}\n")