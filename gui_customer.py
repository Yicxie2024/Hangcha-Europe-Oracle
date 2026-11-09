"""Conservative customer-name recognition for Hangcha sales-order PDFs.

The customer is the addressee in the *left-hand address block* above the
Salesorder heading.  Never infer a customer from the employee name, customer
PO, recipient contact, or the Hangcha issuer/footer.  Return None if the
geometry and address evidence do not support one unambiguous addressee.

This module is read-only and intentionally independent of the employee's
folders and Excel workbook.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

_COUNTRIES = re.compile(
    r"(?i)\b(?:germany|deutschland|netherlands|nederland|belgium|belgique|belgi[eë]|"
    r"france|italy|italia|spain|espa[ñn]a|poland|polska|austria|[öo]sterreich|"
    r"switzerland|schweiz|suisse|luxembourg|united kingdom|great britain|"
    r"denmark|danmark|sweden|sverige|norway|norge|finland|suomi|"
    r"czech republic|czechia|slovakia|slovenia|hungary|romania|"
    r"portugal|ireland|croatia|greece)\b"
)
_POSTAL = re.compile(r"\b\d{4,5}\b|\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b", re.I)
_BAD_FIRST = re.compile(
    r"(?i)\b(?:hangcha\s+europe|sales\s*order|salesorder|invoice|payment|"
    r"customer\s+order|order\s+no|orderdatum|delivery\s+(?:note|order)|"
    r"attn\.?|attention|kontakt|contact|phone|telephone|fax|"
    r"email|e-mail|www\.|https?://|iban|bank|vat|ust\.?[- ]?id|"
    r"delivery\s+ad(?:dress)?|delivery\s+to|shipping\s+address|ship\s+to|lieferadresse)\b"
)
_ADDRESS_START = re.compile(
    r"(?i)^(?:\d+\s|str(?:aße|asse|\.)\b|street\b|road\b|"
    r"avenue\b|rue\b|postcode\b|postfach\b|p\.?o\.?\s*box\b)"
)
_LABEL = re.compile(r"(?i)^(?:bill(?:ed)?\s*to|sold\s*to|customer|rechnung\s*an|kunde)\s*:\s*$")


def _clean_line(line: str) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", line)).strip()


def _name_from_address_block(text: str) -> str | None:
    lines = [_clean_line(x) for x in text.splitlines()]
    lines = [x for x in lines if x and x not in {'....', '...'}]
    if lines and _LABEL.match(lines[0]):
        lines.pop(0)
    if len(lines) < 3:
        return None
    name = lines[0]
    if (len(name) < 3 or len(name) > 110 or not re.search(r'[A-Za-zÀ-ÿ]', name)
            or _BAD_FIRST.search(name) or _ADDRESS_START.search(name)
            or _POSTAL.search(name) or _COUNTRIES.fullmatch(name)):
        return None
    # A two-line company name is not guessed. Those layouts require a
    # separately verified template to avoid appending a contact or city.
    subsequent = lines[1:7]
    if not any(_POSTAL.search(x) for x in subsequent):
        return None
    if not any(_COUNTRIES.search(x) for x in subsequent):
        return None
    return name


def extract_order_customer(path: Path) -> str | None:
    """Return official PDF addressee, or None if extraction is not trustworthy.

    Two observed Hangcha Salesorder templates place this address in the left
    column between the issuer logo and the order heading.  Do not use these
    coordinates for arbitrary other order layouts.
    """
    import pdfplumber

    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            return None
        page = pdf.pages[0]
        # Prevent an invoice in an arbitrary folder being treated as an order.
        # Dedupe layered/overprinted characters before extracting the address.
        clean = page.dedupe_chars(tolerance=1.5)
        layout = clean.extract_text() or ''
        if not re.search(r'\bSales\s*order\b|\bSalesorder\b', layout, re.I):
            return None
        if not re.search(r'\bOrder\s+No\.?\b', layout, re.I) or 'Orderdatum' not in layout:
            return None
        w, h = clean.width, clean.height
        # The verified template has TWO address columns: the addressee on the
        # left and a separate "Delivery address" on the right.  v1.5.8 used
        # .67*w as the crop limit, which included the right-hand heading and
        # merged two unrelated lines.  For example, contract 266788 was
        # misidentified as "Delivery ad" instead of BUW Stapler-Service GmbH.
        # Never allow right-column text into the billing/customer block.
        left_bound, right_bound = .075*w, .55*w
        address = clean.crop((left_bound, .135*h, right_bound, .305*h))
        address_text = address.extract_text(x_tolerance=2, y_tolerance=2) or ''
        customer = _name_from_address_block(address_text)
        if not customer:
            return None
        # Detect a possibly truncated long legal name at the right crop edge.
        # If the candidate touches that edge, refuse to guess a partial name;
        # scanner will mark the PDF as needing review instead.
        first_line = next((x for x in address_text.splitlines() if x.strip()), '')
        candidate_chars = [c for c in address.chars
                           if abs(c.get('top', 0) - min((x.get('top', 0) for x in address.chars), default=0)) < 4]
        if (first_line.strip() == customer and candidate_chars
                and max(c.get('x1', 0) for c in candidate_chars) >= right_bound - 3):
            return None
        return customer


def extract_invoice_customer(path: Path) -> str | None:
    """Extract the invoice bill-to addressee solely from the PDF address block.

    Do not confuse the right-side trade/payment summary with the customer, and
    do not fall back to a filename, folder or a customer's purchase-order label.
    If this verified coordinate template does not fit, return None.
    """
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            return None
        page = pdf.pages[0].dedupe_chars(tolerance=1.5)
        text = page.extract_text() or ''
        if not re.search(r'\bInvoice\s+No\.?\b', text, re.I) or not re.search(r'\bInvoice\s+date\b', text, re.I):
            return None
        w, h = page.width, page.height
        left, right = .075*w, .55*w
        address = page.crop((left, .17*h, right, .265*h))
        address_text = address.extract_text(x_tolerance=2, y_tolerance=2) or ''
        name = _name_from_address_block(address_text)
        if not name:
            return None
        first_top = min((c.get('top', 0) for c in address.chars), default=0)
        first_chars = [c for c in address.chars if abs(c.get('top', 0) - first_top) < 4]
        if first_chars and max(c.get('x1', 0) for c in first_chars) >= right - 3:
            return None
        return name


def extract_delivery_recipient(path: Path) -> str | None:
    """Read the *receiving company* from a Delivery order/note PDF itself.

    The delivery recipient is not necessarily the purchaser on the related
    sales order or invoice. Use the same verified left-address region as the
    order, but do not infer a recipient from a filename or parent directory.
    Unknown/changed layouts return None for review, not a guessed name.
    """
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            return None
        page = pdf.pages[0].dedupe_chars(tolerance=1.5)
        text = page.extract_text() or ''
        if not re.search(r'(?im)^\s*(?:delivery\s*(?:order|note)|lieferschein)\s*$', text):
            return None
        if not re.search(r'(?i)\bDelivery\s+No\.?\b|\bLieferschein\s*(?:Nr\.?|Nummer)', text):
            return None
        w, h = page.width, page.height
        left, right = .075*w, .55*w
        # Verified on 266954: customer address is above the "Delivery order"
        # heading, with Hangcha issuer text above and a footer far below.
        address = page.crop((left, .135*h, right, .28*h))
        addr_text = address.extract_text(x_tolerance=2, y_tolerance=2) or ''
        recipient = _name_from_address_block(addr_text)
        if not recipient:
            return None
        # Decline if the first line is clipped at the right boundary.
        first_top = min((c.get('top', 0) for c in address.chars), default=0)
        first_chars = [c for c in address.chars if abs(c.get('top', 0) - first_top) < 4]
        if first_chars and max(c.get('x1', 0) for c in first_chars) >= right - 3:
            return None
        return recipient
