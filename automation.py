"""Hangcha Europe: deterministic local PDF -> Excel weekly summary automation.
Requires Windows, installed Microsoft Excel and packages from requirements.txt.
This application never edits a PDF. Existing Excel cells are never overwritten.
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
import re
import shutil
import sqlite3
import sys
import traceback
import unicodedata
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, replace
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

BASE = Path(__file__).resolve().parent
STATE = BASE / "state"
REPORTS = STATE / "reports"
BACKUPS = STATE / "backups"
LOCK = STATE / "running.lock"
CONFIG_FILE = BASE / "config.json"
VERSION = "1.2.0"
EXCEL_PREFIX = "业务周汇总表_XYC_更新至"
WK_DIR = re.compile(r"^到\s*(\d{1,2})[.\-](\d{1,2})[.\-](\d{4})$", re.I)
WK_LABEL = re.compile(r"^\s*KW\s*(\d{1,2})(?:\s*[（(]\s*(\d{4})\s*[)）])?\s*$", re.I)
DATE_TOKEN = r"\d{1,2}[-./]\d{1,2}[-./]\d{4}"
AMOUNT_TOKEN = r"(?:\d{1,3}(?:[.\s]\d{3})*|\d+),\d{2}"
DOC_DANGER = ("作废",)  # Explicit user workflow flag only, never a business-field source
DELIVERY_FOLDERS = {"发货单", "送货单", "delivery note", "delivery notes", "lieferschein", "lieferscheine"}
XL_UP = -4162
XL_SHIFT_DOWN = -4121

EXCEL_HEADERS = {
    "订单": ["下单日期", "合同号", "总金额", "Netto", "客户", "业务员", "备注"],
    "发货": ["客户名称", "合同号", "发票号", "发票金额", "Netto", "到货期", "发货日期", "付款期限", "发票发送日期", "客户付款日期", "业务员", "备注"],
}



@dataclass(frozen=True)
class Record:
    kind: str
    contract: str
    document_date: date
    net: Decimal
    gross: Decimal
    vat: Decimal | None
    customer: str
    salesperson: str
    file: str
    sha: str
    invoice: str = ""
    payment_days: int = 30
    skip_order_invoice_amount_check: bool = False  # 有误: skip order AND existing Excel amount comparisons (keep other checks)

    @property
    def key(self):
        return (self.kind, self.contract if self.kind == "order" else self.contract + ":" + self.invoice)

    @property
    def week(self):
        d = self.document_date.isocalendar()
        return d.year, d.week


def normalize_id(value):
    if value is None:
        return ""
    if isinstance(value, (int, float)) and int(value) == value:
        return str(int(value))
    return str(value).strip().replace(" ", "").upper()


def cell_amount(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        return Decimal(str(value)).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def euro(value: str) -> Decimal:
    s = re.sub(r"[^\d,.]", "", value).replace(".", "").replace(",", ".")
    return Decimal(s).quantize(Decimal("0.01"))


def parse_date(value: str) -> date:
    return datetime.strptime(value.replace("/", "-").replace(".", "-"), "%d-%m-%Y").date()


def get_match(pattern, text, label):
    m = re.search(pattern, text, flags=re.I | re.M)
    if not m:
        raise ValueError(f"未识别到 {label}")
    return m


def money_to_decimal(token):
    """Read EUR-style or English-style decimal amounts safely."""
    value = re.sub(r"[^0-9,.-]", "", token)
    if value.count('-') > 1 or ('-' in value and not value.startswith('-')):
        raise ValueError(f"不是有效金额：{token!r}")
    # Last separator before exactly two digits is the decimal mark.
    m = re.fullmatch(r"(-?)([0-9][0-9,\.\s]*)([,.])([0-9]{2})", value)
    if not m:
        raise ValueError(f"不是有效的两位小数金额：{token!r}")
    whole = re.sub(r"[,.\s]", "", m.group(2))
    return Decimal(f"{m.group(1)}{whole}.{m.group(4)}").quantize(Decimal("0.01"))


def summary_excerpt(text, max_lines=3):
    """Short local diagnostic. Only summary lines, never full PDF text or bank details."""
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]
    hints = [i for i,line in enumerate(lines) if re.search(r"(?i)(exclud(?:ing)?\s*vat|excl\.?\s*vat|net(?:to)?\s*total)", line)]
    if not hints:
        return "未找到 Excluding VAT / Netto 汇总标题"
    i = hints[-1]
    safe = []
    for x in lines[i:i+max_lines]:
        if re.search(r"\b(?:IBAN|Steuernummer|Ust\.-?IdNr|Geschäftsführer|BIC)\b", x, re.I):
            break
        if x: safe.append(x[:180])
    return " || ".join(safe)[:500]


def extract_totals(text):
    """Read net/gross directly from the *labelled final PDF summary*.

    3 explicitly displayed values: Netto / VAT / Total (VAT is only validation).
    2 explicitly displayed values: Netto / Total, if there is no VAT column.
    Never infer Netto from a tax rate or from Total. Never assume 19% VAT.
    """
    amount = r"-?(?:\d{1,3}(?:[.,\s]\d{3})+|\d+)[,.]\d{2}"
    euro_amount = re.compile(rf"(?<![\w\d])(?:€|EUR)\s*({amount})|(?<![\d])({amount})\s*(?:€|EUR)\b", re.I)
    plain_amount = re.compile(rf"(?<![\d\w])({amount})(?!\d)")
    lines = [re.sub(r"[\u00a0\u202f]", " ", ln) for ln in text.splitlines()]
    header_rows = []
    for i, line in enumerate(lines):
        low = re.sub(r"\s+", " ", line.lower())
        has_net = bool(re.search(r"exclud(?:ing)?\s*vat|excl\.?\s*vat|net(?:to)?\s*(?:total|amount)?\b", low))
        has_total = bool(re.search(r"\btotal\b|\bgesamt(?:betrag)?\b|\bgrand total\b", low))
        if has_net and has_total:
            without_net_label = re.sub(r"exclud(?:ing)?\s*vat|excl\.?\s*vat", "", low)
            has_vat = bool(re.search(r"\bvat\b|\bmwst\b|\btax\b", without_net_label))
            # Some headers say 'Excluding VAT ... Total' but do not actually
            # carry a VAT tax column. Only VAT outside Netto's own label counts.
            header_rows.append((i, has_vat))
    if not header_rows:
        raise ValueError("未找到包含明确 Netto / Total 标签的金额汇总区：" + summary_excerpt(text))
    candidates = []
    for i, has_vat in header_rows:
        for j in range(i + 1, min(i + 4, len(lines))):
            line = lines[j]
            if re.search(r"\b(?:IBAN|Steuernummer|Ust\.-?IdNr|Geschäftsführer|BIC)\b", line, re.I):
                break
            vals = [a or b for a, b in euro_amount.findall(line)]
            if not vals:
                vals = plain_amount.findall(line)
            if has_vat:
                if len(vals) != 3:
                    continue
                try:
                    net, vat, gross = (money_to_decimal(x) for x in vals)
                except ValueError:
                    continue
                if net < 0 or vat < 0 or gross < 0:
                    continue
                if abs(net + vat - gross) <= Decimal('0.02'):
                    candidates.append((net, vat, gross))
            else:
                if len(vals) != 2:
                    continue
                try:
                    net, gross = (money_to_decimal(x) for x in vals)
                except ValueError:
                    continue
                if net < 0 or gross < 0:
                    continue
                # No VAT amount is available; do not infer or invent it.
                candidates.append((net, None, gross))
    uniq = set(candidates)
    if len(uniq) != 1:
        raise ValueError(f"付款汇总金额匹配到 {len(uniq)} 组有效结果，暂不自动录入；汇总片段：{summary_excerpt(text)}")
    return next(iter(uniq))


def extract_invoice_header(text: str) -> tuple[str, date]:
    """Read values from an Invoice No. / Invoice date header row.

    Additional named columns (e.g. Customer No.) are allowed *between* the
    invoice ID and the date.  Do not use the customer number, delivery number,
    customer-order date or free-form dates as substitutes for these fields.
    Only rows immediately below an explicit invoice header are considered.
    Multi-page PDFs must agree on one (invoice number, invoice date) pair.
    """
    lines = text.splitlines()
    found = set()
    header_line = re.compile(r"\bInvoice\s+No\.?\b", re.I)
    date_label = re.compile(r"\bInvoice\s+date\b", re.I)
    value_row = re.compile(r"^\s*\d+\s*-\s*\d+\s+(H\d{5,12})\b(?P<rest>.*)$", re.I)
    dates = re.compile(rf"(?<!\d)({DATE_TOKEN})(?!\d)")
    for i, line in enumerate(lines):
        if not (header_line.search(line) and date_label.search(line)):
            continue
        # Invoice No. must precede Invoice date in this supported template.
        if header_line.search(line).start() >= date_label.search(line).start():
            continue
        # The numeric data line is directly below the column heading (allow
        # one extraction-related empty line).  Do not search arbitrary prose.
        for candidate in lines[i+1:i+4]:
            if not candidate.strip():
                continue
            match = value_row.match(candidate)
            if not match:
                break
            invoice = normalize_id(match.group(1))
            tail = match.group('rest').strip()
            date_matches = dates.findall(tail)
            # An extra date in a different column makes this ambiguous;
            # never guess which one is the invoice date.
            if len(date_matches) != 1:
                raise ValueError('发票页眉包含多个日期或缺少 Invoice date，需人工复核')
            # No intervening date is permitted; non-date tokens may represent
            # Customer No., VAT No., etc. but are not financial identifiers.
            prefix = tail[:dates.search(tail).start()].strip()
            if len(prefix.split()) > 8:
                raise ValueError('发票页眉额外字段过多，需人工复核')
            found.add((invoice, parse_date(date_matches[0])))
            break
    if len(found) != 1:
        raise ValueError('PDF 内发票页眉缺失，或含多个不同发票号/日期')
    return next(iter(found))

def extract_pdf(path: Path, kind: str, customer: str, default_agent: str, *, customer_selected_by_scanner: bool = False):
    import pdfplumber
    if "作废" in unicodedata.normalize("NFKC", path.name):
        raise ValueError("该 PDF 已由文件名明确标记作废，不进入业务记录")
    stat_before = (path.stat().st_size, path.stat().st_mtime_ns)
    if stat_before[0] == 0:
        raise ValueError("PDF 文件为空")
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    with pdfplumber.open(path) as pdf:
        pages = [(p.extract_text() or "") for p in pdf.pages]
    if stat_before != (path.stat().st_size, path.stat().st_mtime_ns):
        raise ValueError("PDF 正在修改，已跳过本次处理")
    text = "\n".join(pages)
    if len(text.strip()) < 50:
        raise ValueError("PDF 没有足够可提取文字（可能需要 OCR）")
    if re.search(r"(?i)\bcredit\s+note\b|\bstorno\b|\bgutschrift\b", text):
        raise ValueError("PDF 正文可能是贷项/作废文件，请人工复核")
    net, vat, gross = extract_totals(text)
    payment = re.search(r"Payment\s*\(\s*(\d+)\s*days\s*\)", text, re.I)
    if not payment:
        payment = re.search(r"Payment\s*:\s*Bank\s+transfer\.?\s*(\d+)\s+days", text, re.I)
    if kind == "order":
        if "Order No." not in text or "Orderdatum" not in text:
            raise ValueError("不符合当前已验证的 Salesorder 模板")
        header_regex = rf"(?m)^\s*\d+\s*-\s*\d+\s+(\d{{5,9}})\s+.*?({DATE_TOKEN})\s*$"
        all_headers = re.findall(header_regex, text, re.I|re.M)
        if len(set(all_headers)) != 1:
            raise ValueError("PDF 内订单页眉缺失，或含多个不同订单号/日期")
        contract, day = all_headers[0][0], parse_date(all_headers[0][1])
        if not customer_selected_by_scanner:
            from gui_customer import extract_order_customer
            customer = extract_order_customer(path) or '未确定'
        # Scanner-supplied customer always comes from a verified PDF or an
        # exact contract link; never use the legacy caller's folder label.
        return Record(kind, contract, day, net, gross, vat, customer, default_agent, str(path), digest)
    if "Invoice No." not in text or "Invoice date" not in text:
        raise ValueError("不符合当前已验证的 Invoice 模板")
    invoice, day = extract_invoice_header(text)
    contract = get_match(r"Ordernumber\s*:\s*(\d{5,9})", text, "发票关联合同号").group(1)
    if not payment:
        raise ValueError("未识别到付款期限 Payment (xx days)")
    if not customer_selected_by_scanner:
        from gui_customer import extract_invoice_customer
        customer = extract_invoice_customer(path) or '未确定'
    exempt = "有误" in unicodedata.normalize("NFKC", path.name)
    return Record(kind, contract, day, net, gross, vat, customer, default_agent, str(path), digest,
                  invoice=invoice, payment_days=int(payment.group(1)),
                  skip_order_invoice_amount_check=exempt)


def classify_pdf(path: Path, root: Path):
    """Legacy CLI discovery: paths delimit scan scope; PDF content decides kind.

    Kept compatible with the old weekly tree, but never returns the customer
    directory name as customer identity. No filename/folder type inference.
    """
    try:
        parts = path.relative_to(root).parts
    except ValueError:
        return None
    if len(parts) < 4 or not WK_DIR.fullmatch(parts[0]):
        return None
    if any(x.casefold() in {'muster', '模板', 'backups', 'backup', 'archive', 'archiv'} for x in parts[:-1]):
        return None
    if '作废' in unicodedata.normalize('NFKC', path.name):
        return None
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return None
            body = pdf.pages[0].dedupe_chars(tolerance=1.5).extract_text() or ''
    except Exception:
        return None
    if re.search(r'(?mi)^\s*(?:credit\s*note|gutschrift|credit\s*memo)\b', body):
        return None
    from gui_scanner import content_hint
    kind = content_hint(body)
    if kind not in ('order', 'invoice'):
        return None
    return kind, '未确定', parts[0]


def discover(root, include_ignored=False):
    rows, ignored_cancelled = [], []
    for p in sorted(root.rglob('*')):
        if not p.is_file() or p.suffix.lower() != '.pdf':
            continue
        if '作废' in unicodedata.normalize('NFKC', p.name):
            if include_ignored:
                ignored_cancelled.append(str(p))
            continue
        cls = classify_pdf(p, root)
        if cls is not None:
            rows.append((p, *cls))
    return (rows, ignored_cancelled) if include_ignored else rows


def discover_delivery_notes(root: Path):
    """Recognize delivery notes from PDF headings, not a directory name."""
    from gui_scanner import content_hint
    found = []
    for path in sorted(root.rglob('*.pdf')):
        if '作废' in unicodedata.normalize('NFKC', path.name):
            continue
        try:
            import pdfplumber
            with pdfplumber.open(path) as pdf:
                if pdf.pages and content_hint(pdf.pages[0].dedupe_chars(tolerance=1.5).extract_text() or '') == 'delivery':
                    found.append((path, '未确定'))
        except Exception:
            continue
    return found


def explicit_contract_refs(text: str) -> set[str]:
    """Find only explicitly LABELLED sales-order/contract numbers.

    Never interpret a delivery-note number or customer's own purchase order
    as Hangcha's order number. No +/-1 proximity inference.
    """
    label = (r'Ordernumber|(?:Sales\s+)?Order\s*(?:No\.?|Number|#)'
             r'|Auftragsnummer|Auftrags(?:-|\s*)Nr\.?|Auftrag\s*(?:Nr\.?|Nummer)'
             r'|合同号|订单号|销售订单号')
    pat = re.compile(rf'(?<![\w])(?:{label})\s*[:=#.\-\s]*([0-9]{{5,9}})\b', re.I)
    result = set()
    for line in text.splitlines():
        for m in pat.finditer(line):
            # 'Customer order number' and 'Your order no.' are external PO IDs.
            before = line[max(0, m.start()-25):m.start()]
            if re.search(r'(?i)(?:customer|your|kunden)\s*$', before):
                continue
            result.add(m.group(1))
    return result


def delivery_order_refs(text: str) -> set[str]:
    """Only the explicitly labelled internal Ordernumber links a delivery PDF.

    'Delivery No.' is a separate shipment identifier. 'Customer order number'
    is the buyer's purchase order. Neither may substitute for Ordernumber,
    even when their numeric values resemble an order number.
    """
    pat = re.compile(r'(?i)(?<![\w])Ordernumber\s*[:：]\s*([0-9]{5,9})\b')
    refs = set()
    for line in text.splitlines():
        for match in pat.finditer(line):
            preceding = line[max(0,match.start()-35):match.start()]
            if re.search(r'(?i)(?:customer|your|kunden)\s*$', preceding):
                continue
            refs.add(match.group(1))
    return refs


def delivery_header_fields(text: str) -> tuple[str, str]:
    """Read (Delivery No., departure date) from the *labelled* delivery header.

    Never use a neighbouring Ordernumber, Customer order number or arbitrary
    date as a substitute. Duplicate pages may repeat identical fields; if
    pages conflict, return no values and let the scanner request review.
    The returned date is an independent shipping fact, NOT the Excel invoice
    date, and must never be used to add an Excel 发货 record.
    """
    entries: set[tuple[str, str]] = set()
    lines = text.splitlines()
    header = re.compile(r'(?i)\bPage\b.*\bDelivery\s+No\.?(?=\s|$)')
    departure = re.compile(r'(?i)\bDate\s+of\s+departure\b')
    values = re.compile(rf'^\s*\d+\s*-\s*\d+\s+(\d{{5,9}})\b.*?({DATE_TOKEN})(?:\s|$)', re.I)
    for i, line in enumerate(lines):
        if not (header.search(line) and departure.search(line)):
            continue
        # The next non-empty line contains the header values, never an
        # unrelated line much further down in the document.
        for next_line in lines[i + 1:i + 4]:
            if not next_line.strip():
                continue
            m = values.match(next_line)
            if m:
                try:
                    when = parse_date(m.group(2)).strftime('%d.%m.%Y')
                except ValueError:
                    break
                entries.add((m.group(1), when))
            break
    return next(iter(entries)) if len(entries) == 1 else ('', '')


def delivery_document_number(text: str) -> str:
    """Return only a reliably identified Delivery No., not the contract."""
    return delivery_header_fields(text)[0]


def pdf_delivery_contract_refs(path: Path) -> set[str]:
    """Read delivery PDF and extract only its internal Ordernumber labels."""
    import pdfplumber
    old = (path.stat().st_size, path.stat().st_mtime_ns)
    if old[0] <= 0:
        raise ValueError('PDF 文件为空')
    with pdfplumber.open(path) as pdf:
        body = '\n'.join(page.dedupe_chars(tolerance=1.5).extract_text() or '' for page in pdf.pages)
    if old != (path.stat().st_size, path.stat().st_mtime_ns):
        raise ValueError('PDF 读取期间发生修改')
    return delivery_order_refs(body)


def pdf_contract_refs(path: Path) -> set[str]:
    """Lightweight order reference extraction; independent of PDF total parsing."""
    import pdfplumber
    state = (path.stat().st_size, path.stat().st_mtime_ns)
    if state[0] <= 0:
        raise ValueError('PDF 文件为空')
    with pdfplumber.open(path) as pdf:
        text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
    if state != (path.stat().st_size, path.stat().st_mtime_ns):
        raise ValueError('PDF 读取期间发生修改')
    refs = explicit_contract_refs(text)
    return refs



def pdf_order_identity(path: Path):
    """Read a Salesorder's contract and Orderdatum WITHOUT requiring amounts.

    This is for warnings only: a malformed PDF must still never be auto-added
    to the Excel workbook.
    """
    import pdfplumber
    old = (path.stat().st_size, path.stat().st_mtime_ns)
    if old[0] <= 0:
        return None
    with pdfplumber.open(path) as pdf:
        text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
    if old != (path.stat().st_size, path.stat().st_mtime_ns):
        return None
    if 'Order No.' not in text or 'Orderdatum' not in text:
        return None
    header = rf'(?m)^\s*\d+\s*-\s*\d+\s+(\d{{5,9}})\s+.*?({DATE_TOKEN})\s*$'
    matches = set(re.findall(header, text, re.I | re.M))
    if len(matches) != 1:
        return None
    contract, day = matches.pop()
    try:
        return contract, parse_date(day)
    except ValueError:
        return None


def days_before_next_month(d: date):
    if d.month == 12:
        first = date(d.year + 1, 1, 1)
    else:
        first = date(d.year, d.month + 1, 1)
    return (first - d).days


def missing_month_end_invoices(records, orders_existing, invoices_existing,
                               delivery_notes, invoice_files, today: date,
                               warning_days=5, fallback_orders=None):
    """Read-only missing-invoice alerts; never synthesize invoice/Excel entries.

    Delivery note must explicitly reference a sales order. Both the invoice
    folder and existing '发货' sheet must be empty for this contract. A PDF with
    unreadable amount still counts as an invoice if its Ordernumber is present.
    """
    warnings, notices = [], []
    orders = {}
    for r in records:
        if r.kind == 'order':
            orders[r.contract] = (r.document_date, r.customer, r.file)
    for contract, ordered, customer, source in (fallback_orders or []):
        orders.setdefault(contract, (ordered, customer, source))
    for contract, rows in orders_existing.items():
        if contract in orders:
            continue  # PDF order date takes precedence over manually entered date
        dated = [(x.get('order_date'), x.get('customer', ''), f'Excel 订单行 {x.get("row", "?")}')
                 for x in rows if isinstance(x.get('order_date'), date)]
        if dated:
            dated.sort(key=lambda x:x[0])
            orders[contract] = dated[0]
    eligible = {c:(d, customer, src) for c,(d,customer,src) in orders.items()
                if d <= today and 1 <= days_before_next_month(d) <= warning_days}
    if not eligible:
        return warnings, notices, len(delivery_notes)

    # '发货' existing records are authoritative manual entries for this contract.
    shipped_excel = {c for c, _inv in invoices_existing}
    has_invoice = {r.contract for r in records if r.kind == 'invoice'}
    unknown_invoices = defaultdict(list)
    for path, customer in invoice_files:
        try:
            refs = pdf_contract_refs(path)
        except Exception as exc:
            unknown_invoices[customer.casefold()].append(str(path))
            continue
        if refs:
            has_invoice.update(refs)
        else:
            unknown_invoices[customer.casefold()].append(str(path))

    shipment_refs = defaultdict(list)
    unresolved_delivery_notes = defaultdict(list)
    for path, customer in delivery_notes:
        try:
            refs = pdf_delivery_contract_refs(path)
        except Exception:
            refs = set()
        if refs:
            for contract in refs:
                shipment_refs[contract].append(str(path))
        else:
            unresolved_delivery_notes[customer.casefold()].append(str(path))

    for contract, (ordered, customer, src) in sorted(eligible.items()):
        if contract in has_invoice or contract in shipped_excel:
            continue
        proof = shipment_refs.get(contract, [])
        if not proof:
            if unresolved_delivery_notes.get(customer.casefold()):
                notices.append(f'【发货单无法关联】合同 {contract}（订单日期 {ordered:%d.%m.%Y}）：'
                               f'客户 {customer} 的发货单未识别出明确 Ordernumber/合同号，无法判断是否触发缺票提醒；'
                               f'请提供发货单模板核对：{unresolved_delivery_notes[customer.casefold()][0]}')
            continue
        # With PDF-only customer identities, some unreadable invoices have no
        # reliable customer. Treat them as globally ambiguous, not as proof of
        # absence for one named customer's contract.
        ambiguous_invoice = list(unknown_invoices.get(customer.casefold(), []))
        if customer.casefold() != '未确定':
            ambiguous_invoice += unknown_invoices.get('未确定', [])
        if ambiguous_invoice:
            warnings.append(f'【月末开票状态待核对】合同 {contract}（订单日期 {ordered:%d.%m.%Y}，距次月'
                            f'{days_before_next_month(ordered)} 天）：已找到关联合同的发货单，但同一客户存在'
                            f'无法确认合同归属的发票 PDF，不直接断言缺票；发货表无对应合同记录。'
                            f'发货单：{proof[0]}；待核查发票：{ambiguous_invoice[0]}')
        else:
            warnings.append(f'【月末已发货但缺发票】合同 {contract}（订单日期 {ordered:%d.%m.%Y}，距次月'
                            f'{days_before_next_month(ordered)} 天）：已找到发货单，未找到有效对应发票 PDF，'
                            f'发货子表也无该合同记录；请检查是否漏开票/漏归档。发货单：{proof[0]}')
    return warnings, notices, len(delivery_notes)


def find_workbook(root):
    """A missing summary is valid: create one from PDFs on --apply.

    Multiple matching files are ALWAYS an error, to protect manual edits.
    """
    matches = [p for p in root.iterdir() if p.is_file() and p.suffix.lower() == ".xlsx"
               and p.name.startswith(EXCEL_PREFIX) and not p.name.startswith("~$")]
    if len(matches) > 1:
        raise RuntimeError(f"根目录发现 {len(matches)} 份 '{EXCEL_PREFIX}*.xlsx'，请先保留唯一一份；程序不会猜测哪份为准")
    return matches[0] if matches else None


def empty_sheet(ws):
    """Formatted-only worksheets count as empty; formulas and manual data do not."""
    return int(ws.Application.WorksheetFunction.CountA(ws.UsedRange)) == 0


def make_header(ws, kind):
    """Initialize an empty worksheet without changing any populated cells."""
    cols = EXCEL_HEADERS[kind]
    if not empty_sheet(ws):
        raise RuntimeError(f"工作表 {kind} 含有内容，不能自动重建表头")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, len(cols))).Value2 = (tuple(cols),)
    h = ws.Range(ws.Cells(1, 1), ws.Cells(1, len(cols)))
    h.Font.Name = 'Calibri'
    h.Font.Bold = True
    h.Font.Color = 0xFFFFFF
    h.Interior.Color = 0x5B3B1C  # muted dark blue (Excel uses BGR)
    h.RowHeight = 25
    h.HorizontalAlignment = -4108  # xlCenter
    widths = ([17, 16, 15, 15, 22, 14, 28] if kind == 'order'
              else [24, 16, 18, 16, 16, 17, 17, 15, 18, 18, 14, 28])
    for c, width in enumerate(widths, 1):
        ws.Columns(c).ColumnWidth = width
    # Plain-text date columns; all newly generated dates remain text.
    for col in ([1] if kind == 'order' else [6, 7, 9]):
        ws.Columns(col).NumberFormat = '@'


def prepare_workbook(wb, from_scratch=False):
    """Use intact existing sheets, create absent or blank sheets safely.

    Only reuse truly empty sheets. Never clear or rename non-empty manual sheets.
    Return True when one or more sheets were initialized.
    """
    changed = False
    existing = [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
    all_empty_at_start = all(empty_sheet(ws) for ws in existing)
    for name, cols in EXCEL_HEADERS.items():
        sheet = next((x for x in existing if str(x.Name) == name), None)
        if sheet is None:
            unused = next((x for x in existing if str(x.Name) not in EXCEL_HEADERS
                           and empty_sheet(x)), None)
            if unused is not None:
                unused.Name = name
                sheet = unused
            else:
                sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                sheet.Name = name
            existing = [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
            changed = True
        values = [str(sheet.Cells(1, col).Value or '').strip() for col in range(1, len(cols)+1)]
        if values == cols:
            continue
        if not empty_sheet(sheet):
            raise RuntimeError(f"工作表 {name} 有内容但表头不符合要求。现有内容不会被清除；预期：{cols}，实际：{values}")
        make_header(sheet, name)
        changed = True
    # A genuinely empty source workbook (including Excel's default Sheet2/3)
    # should become exactly the two requested sheets. Nonempty other sheets
    # are always preserved and cannot be deleted automatically.
    for other in list(existing):
        if (str(other.Name) not in EXCEL_HEADERS and empty_sheet(other)
                and (from_scratch or (all_empty_at_start and re.fullmatch(
                    r'(?:Sheet|Tabelle|Feuil|Foglio|Hoja)\d+', str(other.Name), re.I)))):
            other.Delete()
            changed = True
    verify_template(wb)
    return changed


def verify_template(wb):
    for name, columns in EXCEL_HEADERS.items():
        ws = wb.Worksheets(name)
        actual = [str(ws.Cells(1, i + 1).Value or '').strip() for i in range(len(columns))]
        if actual != columns:
            raise RuntimeError(f"工作表 {name} 标题已变化，停止写入。预期：{columns}；实际：{actual}")


def week_identity(day: date):
    """ISO-8601 week with the appropriate ISO week-year, not a fixed 2026."""
    calendar = day.isocalendar()
    return calendar.year, calendar.week


def week_label(day: date, current: date | None = None):
    """Ordinary current-year KW labels; explicit year on historical/future rows."""
    wk_year, wk = week_identity(day)
    today_year = week_identity(current or date.today())[0]
    return f'KW  {wk:02d}' if wk_year == today_year else f'KW  {wk:02d} ({wk_year})'


def date_text(day: date):
    return day.strftime('%d.%m.%Y')


def excel_date(x):
    if x is None or x == "":
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    if isinstance(x, (float, int)):
        from datetime import timedelta
        return date(1899, 12, 30) + timedelta(days=int(x))
    try:
        return parse_date(str(x))
    except ValueError:
        return None


def last_data_row(ws, key_col):
    return max(1, int(ws.Cells(ws.Rows.Count, key_col).End(XL_UP).Row))


def sheet_snapshot(ws, kind):
    end = last_data_row(ws, 1)
    # Identify year via a date in the section. Headers deliberately remain KW NN in original template.
    rows = {}
    headers = []
    for r in range(2, end + 1):
        first = ws.Cells(r, 1).Value
        wm = WK_LABEL.match(str(first).strip()) if first is not None else None
        if wm:
            headers.append([r, int(wm.group(1)), int(wm.group(2)) if wm.group(2) else None])
            continue
        contract = normalize_id(ws.Cells(r, 2).Value)
        if not re.fullmatch(r"\d{5,9}", contract):
            continue
        if kind == "order":
            rows.setdefault(contract, []).append({"row": r, "gross": cell_amount(ws.Cells(r, 3).Value),
                "net": cell_amount(ws.Cells(r, 4).Value), "customer": str(ws.Cells(r, 5).Value or "").strip(),
                "order_date": excel_date(ws.Cells(r, 1).Value2)})
        else:
            inv = normalize_id(ws.Cells(r, 3).Value)
            if not inv:
                continue
            rows.setdefault((contract, inv), []).append({"row": r, "gross": cell_amount(ws.Cells(r, 4).Value),
                "net": cell_amount(ws.Cells(r, 5).Value), "customer": str(ws.Cells(r, 1).Value or "").strip()})
    # Merged KW section consumes TWO rows even if only first row has a value.
    if headers:
        end = max(end, max(x[0] + 1 for x in headers))
    # Determine section year from existing row dates, using first date available in the section.
    for i, h in enumerate(headers):
        begin, wk, yr = h
        after = headers[i + 1][0] if i + 1 < len(headers) else end + 1
        if yr is None:
            for rr in range(begin + 2, after):
                dt = excel_date(ws.Cells(rr, 1 if kind == "order" else 7).Value2)
                if dt:
                    yr = dt.isocalendar().year
                    break
        # Existing plain 'KW NN' without any dated rows: infer the CURRENT ISO
        # year, as requested. Historical/future sections should carry a year.
        h[2] = yr if yr is not None else week_identity(date.today())[0]
    return rows, headers, end


def compare_existing(rec, existing):
    if not existing:
        return []
    problems = []
    if len(existing) > 1:
        problems.append(f"Excel 中该编号有 {len(existing)} 行")
    # 有误 invoice: amount in the original PDF can intentionally differ from
    # a manually entered Excel amount. Suppress ONLY amount comparisons;
    # keep duplicate-row detection, zero-value checks and identity checks.
    if rec.kind == "invoice" and rec.skip_order_invoice_amount_check:
        return problems
    for row in existing:
        for key in ("gross", "net"):
            if row[key] is not None and abs(row[key] - getattr(rec, key)) > Decimal("0.02"):
                problems.append(f"行 {row['row']} {key}：Excel={row[key]} PDF={getattr(rec,key)}")
    return problems


def normalize_duplicates(records):
    grouped = defaultdict(list)
    for rec in records:
        grouped[rec.key].append(rec)
    unique, issues = [], []
    for key, items in grouped.items():
        first = items[0]
        diff = [r for r in items[1:] if (r.document_date, r.gross, r.net, r.invoice, r.customer.casefold()) !=
                (first.document_date, first.gross, first.net, first.invoice, first.customer.casefold())]
        if diff:
            issues.append(f"{key} 存在多份内容冲突的 PDF：" + "; ".join(r.file for r in items))
        else:
            if first.kind == 'invoice' and len({r.skip_order_invoice_amount_check for r in items}) > 1:
                # Do not let one 有误-labelled copy disable validation of an
                # otherwise unlabelled invoice with the same financial identity.
                issues.append(f"{key} 存在相同发票但文件名有误标记不一致，请人工确认：" + "; ".join(r.file for r in items))
            else:
                unique.append(first)
    return unique, issues


def financial_checks(records, orders_existing, invoices_existing):
    """Detect zero and cross-document differences, including existing manual rows.

    Report a lower cumulative invoice amount as possible partial shipment, not
    as proof of bad data. Block zero-value new records or over-invoiced new
    invoices. Never correct amounts, never infer VAT, never replace manual rows.
    """
    warnings, notices = [], []
    block_orders, block_invoices = set(), set()
    pdf_order = {r.contract: r for r in records if r.kind == 'order'}
    pdf_invoices = {(r.contract, r.invoice): r for r in records if r.kind == 'invoice'}

    # Zero checks apply even to rows entered manually in Excel.
    for r in records:
        fields = [name for name, value in (('总金额', r.gross), ('Netto', r.net)) if value == 0]
        if fields:
            label = f'订单 {r.contract}' if r.kind == 'order' else f'发票 {r.invoice}（合同 {r.contract}）'
            warnings.append(f'【金额为0｜人工复核】PDF {label}：{",".join(fields)} 为 0；禁止自动新增：{r.file}')
            if r.kind == 'order':
                block_orders.add(r.contract)
            else:
                block_invoices.add((r.contract, r.invoice))
    for contract, rows in orders_existing.items():
        for row in rows:
            fields = [name for name, key in (('总金额', 'gross'), ('Netto', 'net')) if row[key] == 0]
            if fields:
                warnings.append(f'【金额为0｜人工复核】Excel 订单 合同 {contract} 第 {row["row"]} 行：{",".join(fields)} 为 0；保留人工原值')
                block_orders.add(contract)
    for (contract, invoice), rows in invoices_existing.items():
        for row in rows:
            fields = [name for name, key in (('发票金额', 'gross'), ('Netto', 'net')) if row[key] == 0]
            if fields:
                warnings.append(f'【金额为0｜人工复核】Excel 发货 发票 {invoice}/合同 {contract} 第 {row["row"]} 行：{",".join(fields)} 为 0；保留人工原值')
                block_invoices.add((contract, invoice))

    def order_gross(contract):
        # Compare PDF invoice against the actual PDF order whenever available;
        # Excel amounts are fallback only, and discrepancies against manual
        # Excel values are reported independently by compare_existing().
        order_pdf = pdf_order.get(contract)
        if order_pdf is not None:
            return order_pdf.gross
        existing_rows = orders_existing.get(contract, [])
        vals = {r['gross'] for r in existing_rows if r['gross'] is not None}
        if len(vals) > 1:
            warnings.append(f'【订单金额不确定】合同 {contract}：Excel 存在多个不同总金额 {sorted(vals)}，无法可靠比较发票')
            return None
        if vals:
            return next(iter(vals))  # Manual Excel has precedence.
        return None

    # Invoices named 有误 are exempt from BOTH order-vs-invoice amount comparisons
    # and PDF-vs-existing-Excel amounts (compare_existing handles the latter).
    # All other checks remain, including zero amounts, duplicate invoice IDs,
    # internal PDF totals validation and missing orders. Exclude those invoices
    # from same-contract cumulative invoice amount matching as well.
    exempt_invoice_keys = {key for key, rec in pdf_invoices.items()
                           if rec.skip_order_invoice_amount_check}
    for key in sorted(exempt_invoice_keys):
        notices.append(f'【有误发票｜跳过跨记录金额对比】合同 {key[0]} 发票 {key[1]}：已跳过与订单金额和已有 Excel 发货金额的对比；其他安全检查照常执行')

    # All NON-EXEMPT invoices per contract, once per (contract, invoice)
    # regardless of whether PDF and manually entered Excel rows both exist.
    by_contract = defaultdict(dict)
    for key, rows in invoices_existing.items():
        if key in exempt_invoice_keys:
            continue
        if len({row['gross'] for row in rows if row['gross'] is not None}) > 1:
            warnings.append(f'【发票金额不确定】发票 {key[1]}/合同 {key[0]} 在 Excel 有多个金额；已跳过该合同的累计比较')
            continue
        amounts = [row['gross'] for row in rows if row['gross'] is not None]
        if amounts:
            by_contract[key[0]][key] = amounts[0]
    for key, rec in pdf_invoices.items():
        if key in exempt_invoice_keys:
            continue  # Excluded only from order-vs-invoice amount matching.
        by_contract[key[0]][key] = rec.gross  # PDF is the amount source of truth.

    for contract, invoice_amounts in sorted(by_contract.items()):
        target = order_gross(contract)
        if target is None:
            # The ordinary matching check separately warns about missing orders.
            continue
        if target == 0 or contract in block_orders:
            for key in invoice_amounts:
                if key in pdf_invoices:
                    block_invoices.add(key)
            continue
        items = sorted(invoice_amounts.items())
        total = sum((v for _, v in items), Decimal('0.00'))
        diff = (total - target).quantize(Decimal('0.01'))
        if len(items) == 1:
            (key, inv_gross), = items
            if abs(diff) > Decimal('0.02'):
                warnings.append(f'【发票/订单金额不一致】合同 {contract} 发票 {key[1]}：订单总金额 €{target}，发票金额 €{inv_gross}，差额 €{diff:+}; 可能为部分发货/额外费用，需人工检查')
        elif abs(diff) <= Decimal('0.02'):
            notices.append(f'【分批开票，合计匹配】合同 {contract}：{len(items)} 张发票，累计 €{total} = 订单 €{target}')
        else:
            kind = '超出订单' if diff > 0 else '尚低于订单'
            warnings.append(f'【多张发票金额累计不一致】合同 {contract}：{len(items)} 张发票累计 €{total}，订单 €{target}，差额 €{diff:+}（{kind}；可能为分批开票，需核对）')
        # If recorded invoices already exceed the order, new invoice rows
        # require review; existing rows remain completely unchanged.
        if diff > Decimal('0.02'):
            for key in invoice_amounts:
                if key in pdf_invoices and key not in invoices_existing:
                    block_invoices.add(key)
            warnings.append(f'【暂停自动新增发票】合同 {contract} 累计发票金额超过订单，请人工复核后再录入')
    return warnings, notices, block_orders, block_invoices


def calculate_plan(records, orders_existing, invoices_existing):
    new_order, new_invoice, warnings, skipped = [], [], [], []
    financial_warnings, financial_notices, block_orders, block_invoices = financial_checks(
        records, orders_existing, invoices_existing)
    warnings.extend(financial_warnings)
    # Detect legacy Excel invoice rows where the contract column could contain
    # a delivery note number instead of the actual Ordernumber. DO NOT infer
    # a +1 relationship; block both new order and invoice pending review.
    invoice_to_old_contracts = defaultdict(set)
    for contract, number in invoices_existing:
        invoice_to_old_contracts[number].add(contract)
    invoice_conflicts = {}
    for rec in records:
        if rec.kind != "invoice":
            continue
        olds = invoice_to_old_contracts.get(rec.invoice, set())
        if olds and rec.contract not in olds:
            invoice_conflicts[rec.invoice] = (rec.contract, sorted(olds), rec.file)
    blocked_contracts = {contract for contract,old,src in invoice_conflicts.values()}
    for number,(contract,old,src) in sorted(invoice_conflicts.items()):
        warnings.append(f"疑似合同号/发货单号混用：发票 {number} 的 PDF Ordernumber={contract}，但 Excel 已登记合同号={old}；已冻结关联订单及发票，待核对：{src}")

    # Process orders BEFORE invoices; do not overwrite manually edited rows.
    for rec in sorted((r for r in records if r.kind == "order"), key=lambda r:(r.document_date,r.contract)):
        existing = orders_existing.get(rec.contract, [])
        problems = compare_existing(rec, existing)
        if problems:
            warnings.append(f"订单 {rec.contract}: " + "; ".join(problems) + f" | {rec.file}")
        elif rec.contract in blocked_contracts or rec.contract in block_orders:
            warnings.append(f"订单 {rec.contract} 涉及合同冲突或零金额，暂不自动新增：{rec.file}")
        elif existing:
            skipped.append(f"已存在：订单 {rec.contract}")
        else:
            new_order.append(rec)

    known_contracts = set(orders_existing) | {r.contract for r in new_order}
    invoice_contracts = defaultdict(set)
    for contract, number in invoices_existing:
        invoice_contracts[number].add(contract)
    for rec in records:
        if rec.kind == "invoice":
            invoice_contracts[rec.invoice].add(rec.contract)
    for rec in sorted((r for r in records if r.kind == "invoice"), key=lambda r:(r.document_date,r.contract,r.invoice)):
        if rec.invoice in invoice_conflicts:
            continue  # already reported with both numbers and source
        if len(invoice_contracts[rec.invoice]) > 1:
            warnings.append(f"发票号 {rec.invoice} 被不同合同引用 {sorted(invoice_contracts[rec.invoice])}，已跳过：{rec.file}")
            continue
        existing = invoices_existing.get((rec.contract, rec.invoice), [])
        problems = compare_existing(rec, existing)
        if problems:
            warnings.append(f"发票 {rec.invoice}/{rec.contract}: " + "; ".join(problems) + f" | {rec.file}")
        elif existing:
            skipped.append(f"已存在：发票 {rec.invoice}/{rec.contract}")
        elif (rec.contract, rec.invoice) in block_invoices or rec.contract in block_orders:
            warnings.append(f"发票 {rec.invoice}/{rec.contract} 涉及零金额/累计超额等高风险情况，暂不自动新增：{rec.file}")
        elif rec.contract not in known_contracts:
            warnings.append(f"发票 {rec.invoice} 的合同号 {rec.contract} 未在订单表/有效订单PDF中找到，暂不录入：{rec.file}")
        else:
            new_invoice.append(rec)
    return new_order, new_invoice, warnings, skipped, financial_notices

def get_groups(ws, kind):
    _, headers, end = sheet_snapshot(ws, kind)
    groups = []
    for i, (begin, week, year) in enumerate(headers):
        next_head = headers[i+1][0] if i+1 < len(headers) else end+1
        groups.append({"start":begin,"week":week,"year":year,"next":next_head})
    return groups, end


def add_week_header(ws, pos: int, kind: str, day: date, template_row=None):
    """Make a two-row, A:E merged KW heading, even in a brand-new workbook."""
    end_col = 'G' if kind == 'order' else 'L'
    if template_row is not None:
        ws.Range(f'A{template_row}:{end_col}{template_row+1}').Copy(
            Destination=ws.Range(f'A{pos}:{end_col}{pos+1}'))
        ws.Application.CutCopyMode = False
    else:
        head = ws.Range(f'A{pos}:E{pos+1}')
        head.Merge()
        head.Interior.Color = 0xE8E1D3
        head.Font.Bold = True
        head.Font.Color = 0x543B27
        head.Font.Size = 11
        head.HorizontalAlignment = -4131  # xlLeft
        head.VerticalAlignment = -4108  # xlCenter
        ws.Rows(f'{pos}:{pos+1}').RowHeight = 19
    ws.Cells(pos, 1).Value2 = week_label(day)


def ensure_week(ws, kind, d):
    wkyear, wk = week_identity(d)
    groups, end = get_groups(ws, kind)
    if not groups and end > 1:
        raise RuntimeError('现有工作表有数据但没有 KW 周次标题，无法安全自动分组；原记录未被修改')
    possible = [x for x in groups if x['week'] == wk and x['year'] == wkyear]
    if len(possible) > 1:
        raise RuntimeError(f'多个 {wkyear} KW {wk} 区块，无法定位')
    if possible:
        return possible[0]['next'], False
    # Insert before any later week-year group, otherwise append. Never
    # overwrite the existing historic groups / user-edited records.
    following = sorted((x for x in groups if (x['year'], x['week']) > (wkyear, wk)),
                       key=lambda x:(x['year'], x['week']))
    insertion = following[0]['start'] if following else end + 1
    template = groups[0]['start'] if following else (groups[-1]['start'] if groups else None)
    if insertion <= end:
        ws.Rows(f'{insertion}:{insertion+1}').Insert(Shift=XL_SHIFT_DOWN)
        if template is not None and template >= insertion:
            template += 2
    add_week_header(ws, insertion, kind, d, template)
    return insertion + 2, False


def append_rec(ws, rec):
    pos, _ = ensure_week(ws, rec.kind, rec.document_date)
    groups, end = get_groups(ws, rec.kind)
    # Sort newly inserted rows by the actual business date in the PDF.
    # Existing manual rows are NEVER globally sorted or replaced.
    for group in groups:
        if (group['year'], group['week']) == rec.week:
            for r in range(group['start'] + 2, group['next']):
                dt = excel_date(ws.Cells(r, 1 if rec.kind == 'order' else 7).Value2)
                if dt and dt > rec.document_date:
                    pos = r
                    break
            break
    if pos <= end:
        ws.Rows(pos).Insert(Shift=XL_SHIFT_DOWN)
    kind = rec.kind
    _, _, last = sheet_snapshot(ws, kind)
    cols = 7 if kind == 'order' else 12
    template = None
    for delta in range(1, last + 1):
        for rr in (pos-delta, pos+delta):
            if 2 <= rr <= last and rr != pos:
                first = ws.Cells(rr, 1).Value
                if first is None or WK_LABEL.match(str(first).strip()):
                    continue
                c = normalize_id(ws.Cells(rr, 2).Value)
                if re.fullmatch(r'\d{5,9}', c):
                    template = rr
                    break
        if template: break
    if template is not None:
        ws.Range(ws.Cells(template, 1), ws.Cells(template, cols)).Copy(
            Destination=ws.Range(ws.Cells(pos, 1), ws.Cells(pos, cols)))
        ws.Application.CutCopyMode = False
    else:
        ws.Range(ws.Cells(pos, 1), ws.Cells(pos, cols)).RowHeight = 19
    if kind == 'order':
        vals = [date_text(rec.document_date), int(rec.contract), float(rec.gross), float(rec.net),
                rec.customer, rec.salesperson, '']
    else:
        dat = date_text(rec.document_date)
        vals = [rec.customer, int(rec.contract), rec.invoice, float(rec.gross), float(rec.net),
                dat, dat, rec.payment_days, dat, '未更新', rec.salesperson, '']
    # CRITICAL: set text format FIRST, then set Value2 to actual Python str.
    # NumberFormat='dd.mm.yyyy' would only change display, not cell type.
    for col in ([1] if kind == 'order' else [6,7,9]):
        ws.Cells(pos, col).NumberFormat = '@'
    ws.Range(ws.Cells(pos, 1), ws.Cells(pos, cols)).Value2 = (tuple(vals),)
    return pos


def update_week_dir(root: Path, today: date, dry_run=False):
    current = []
    for p in root.iterdir():
        if not p.is_dir(): continue
        m = WK_DIR.fullmatch(p.name)
        if not m: continue
        dt = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        if dt.isocalendar()[:2] == today.isocalendar()[:2]:
            current.append(p)
    if len(current) != 1:
        return f"本周文件夹数量为 {len(current)}；已保留原样，请人工确认（周一新周可能需要先建文件夹）"
    dest = root / f"到{today:%d.%m.%Y}"
    if current[0] == dest:
        return "本周文件夹名称已是今日日期"
    if dest.exists():
        return f"目标文件夹已存在，未改名：{dest}"
    if not dry_run:
        current[0].rename(dest)
    return f"{'将重命名' if dry_run else '已重命名'}：{current[0].name} -> {dest.name}"


def write_report(stamp, root, workbook, scanned, new_order, new_invoice, issues, skipped, notices, applied, note, new_path="", ignored_cancelled=None, delivery_notes_scanned=0):
    REPORTS.mkdir(parents=True, exist_ok=True)
    ignored_cancelled = ignored_cancelled or []
    result = {
        "time": stamp.isoformat(timespec="seconds"), "version": VERSION, "mode": "APPLIED" if applied else "PREVIEW",
        "root": str(root), "workbook": str(workbook), "final_workbook": new_path,
        "pdf_scanned": scanned, "delivery_notes_scanned": delivery_notes_scanned, "orders_new": len(new_order), "invoices_new": len(new_invoice),
        "warnings": issues, "notices": notices, "already_recorded": skipped, "note": note,
        "cancelled_invoice_pdfs_ignored": len(ignored_cancelled),
        "cancelled_invoice_pdf_paths": ignored_cancelled,
        "proposed_orders": [dict(contract=r.contract,date=r.document_date.isoformat(),customer=r.customer,
                                 gross=str(r.gross),net=str(r.net),source=r.file) for r in new_order],
        "proposed_invoices": [dict(contract=r.contract,invoice=r.invoice,date=r.document_date.isoformat(),
                                   gross=str(r.gross),net=str(r.net),source=r.file) for r in new_invoice],
    }
    suffix = "applied" if applied else "preview"
    filename = REPORTS / f"{stamp:%Y%m%d_%H%M%S}_{suffix}.json"
    filename.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [f"业务周汇总自动化 v{VERSION} | {stamp:%Y-%m-%d %H:%M:%S}", f"模式：{'已执行' if applied else '只读预览'}",
             f"扫描订单/发票 PDF：{scanned}", f"检查发货单 PDF：{delivery_notes_scanned}", f"待新增订单：{len(new_order)}", f"待新增发货：{len(new_invoice)}",
             f"异常：{len(issues)}", f"重复跳过：{len(skipped)}",
             f"忽略作废发票：{len(ignored_cancelled)}", f"说明：{note}", ""]
    for label, recs in [("订单",new_order),("发货",new_invoice)]:
        lines.append(f"==== {label}明细 ====")
        for r in recs:
            lines.append(f"{r.document_date:%d.%m.%Y} | {r.contract} | {r.invoice} | {r.customer} | {r.gross} / Netto {r.net} | {r.file}")
    if issues:
        lines.extend(["", "==== 异常（需人工检查） ===="] + [f"[{i}] {x}" for i,x in enumerate(issues,1)])
    if notices:
        lines.extend(["", "==== 提示（无需作为错误处理） ===="] + [f"[{i}] {x}" for i,x in enumerate(notices, 1)])
    if skipped:
        lines.extend(["", "==== 已存在（前100条） ===="] + skipped[:100])
    if ignored_cancelled:
        lines.extend(["", "==== 已忽略作废发票（不作为异常） ===="] + ignored_cancelled[:100])
    txt_path = filename.with_suffix(".txt")
    txt_path.write_text("\n".join(lines), encoding="utf-8-sig")
    return txt_path, filename


def popup_warning(text):
    if os.name != "nt": return
    try:
        user32 = ctypes.windll.user32
        fn = getattr(user32, "MessageBoxTimeoutW", None)
        if fn:
            fn(None, text[:1200], "业务周汇总表 - 异常提醒", 0x30 | 0x1000, 0, 25000)
        else:
            user32.MessageBoxW(None, text[:1200], "业务周汇总表 - 异常提醒", 0x30 | 0x1000)
    except Exception:
        pass


def audit_db(stamp, report_path, applied, count_o, count_i, errors, parsed):
    STATE.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(STATE / "audit.sqlite3") as conn:
        conn.execute("CREATE TABLE IF NOT EXISTS runs (id INTEGER PRIMARY KEY, stamp TEXT, report TEXT, applied INTEGER, added_orders INTEGER, added_invoices INTEGER, warnings INTEGER)")
        conn.execute("INSERT INTO runs(stamp,report,applied,added_orders,added_invoices,warnings) VALUES(?,?,?,?,?,?)",
                     (stamp.isoformat(),str(report_path),int(applied),count_o,count_i,len(errors)))
        conn.execute("CREATE TABLE IF NOT EXISTS documents (sha256 TEXT, pdf_path TEXT, kind TEXT, business_key TEXT, seen_at TEXT, PRIMARY KEY(sha256,pdf_path))")
        for rec in parsed:
            conn.execute("INSERT INTO documents(sha256,pdf_path,kind,business_key,seen_at) VALUES(?,?,?,?,?) "
                         "ON CONFLICT(sha256,pdf_path) DO UPDATE SET seen_at=excluded.seen_at",
                         (rec.sha,rec.file,rec.kind,rec.key[1],stamp.isoformat()))


def safe_zip(path):
    with zipfile.ZipFile(path) as z:
        return z.testzip() is None and '[Content_Types].xml' in z.namelist()


def acquire_lock():
    STATE.mkdir(parents=True, exist_ok=True)
    try:
        fd = os.open(str(LOCK), os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        os.write(fd, f"pid={os.getpid()} time={datetime.now().isoformat()}".encode())
        os.close(fd)
    except FileExistsError:
        raise RuntimeError(f"已有任务锁：{LOCK}。如果确认程序已退出，可手动删除此文件。")


def reconcile_record_customers(records: list[Record]):
    """Apply PDF order customer to all matching invoices before Excel planning.

    A differing invoice-PDF customer is informational, not a write blocker;
    multiple conflicting order-PDF customers for one contract still block it.
    This is the legacy CLI path; the GUI uses gui_scanner instead.
    """
    from gui_customer import extract_order_customer
    names: dict[str, dict[str, str]] = {}
    for record in records:
        if record.kind != 'order':
            continue
        try:
            official = extract_order_customer(Path(record.file))
        except Exception:
            official = None
        if official:
            key = unicodedata.normalize('NFKC', official).casefold().strip()
            names.setdefault(record.contract, {})[key] = official
    conflicts = {contract for contract, entries in names.items() if len(entries) > 1}
    warnings = [f'【客户冲突】合同 {c} 对应的订单 PDF 包含不同客户名称；'
                '关联订单和发票均不自动写入，请人工检查：' + ' / '.join(sorted(names[c].values()))
                for c in sorted(conflicts)]
    resolved = []
    for record in records:
        if record.contract in conflicts:
            continue
        choices = names.get(record.contract, {})
        if choices:
            name = next(iter(choices.values()))
            # Only the GUI retains a visible informational note about a
            # different invoice addressee. A name difference is not a CLI
            # anomaly and must not inflate the error count or block a write.
            record = replace(record, customer=name)
        if record.customer == '未确定':
            warnings.append(f'【客户待核】合同 {record.contract} 的 PDF 未能确认客户名称，禁止自动写入 Excel：{record.file}')
            continue
        resolved.append(record)
    return resolved, warnings


def execute(config, apply=False):
    stamp = datetime.now()
    today = stamp.date()
    root = Path(os.path.expandvars(config['root_path'])).expanduser().resolve()
    if not root.is_dir():
        raise RuntimeError(f"目录不存在：{root}")
    acquire_lock()
    excel = None
    wb = None
    original = None
    preview_copy = None
    warnings = []
    try:
        docs, ignored_cancelled = discover(root, include_ignored=True)
        if not docs:
            raise RuntimeError("没有找到 周文件夹/客户/订单或发票 下的 PDF。为了安全，本次不修改 Excel 或文件夹。")
        parsed = []
        fallback_orders = []
        for path, kind, customer, folder in docs:
            try:
                parsed.append(extract_pdf(path, kind, customer, config.get('salesperson','Xie')))
            except Exception as e:
                warnings.append(f"PDF 无法自动识别：{path} | {e}")
                if kind == 'order':
                    try:
                        identity = pdf_order_identity(path)
                        if identity:
                            fallback_orders.append((identity[0], identity[1], customer, str(path)))
                    except Exception:
                        pass  # Original PDF parse warning is already recorded.
        parsed, name_warnings = reconcile_record_customers(parsed)
        warnings.extend(name_warnings)
        parsed, doubles = normalize_duplicates(parsed)
        warnings.extend(doubles)
        if not parsed:
            raise RuntimeError("所有 PDF 都无法通过解析与去重校验。为了安全，本次不修改 Excel 或文件夹；请检查 PDF 模板。")
        original = find_workbook(root)
        # Truly zero-byte placeholders are treated as missing content. An
        # invalid nonempty .xlsx still fails safely on Workbooks.Open.
        zero_placeholder = original is not None and original.stat().st_size == 0
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        # Preview a snapshot of an existing file; if missing/zero-byte,
        # create a NEW in-memory workbook, NEVER a new on-disk workbook.
        if original is None or zero_placeholder:
            wb = excel.Workbooks.Add()
        else:
            open_path = original
            if not apply:
                STATE.mkdir(parents=True, exist_ok=True)
                preview_copy = STATE / f'~preview_copy_{stamp:%Y%m%d_%H%M%S}_{os.getpid()}.xlsx'
                shutil.copy2(original, preview_copy)
                open_path = preview_copy
            wb = excel.Workbooks.Open(str(open_path), UpdateLinks=0, ReadOnly=not apply,
                                      IgnoreReadOnlyRecommended=True)
            if apply and wb.ReadOnly:
                raise RuntimeError('Excel 文件已被其他程序打开，无法取得独占写权限。请关闭 Excel 后重试。')
        was_initialized = prepare_workbook(wb, from_scratch=(original is None or zero_placeholder))
        orders_existing, _, _ = sheet_snapshot(wb.Worksheets('订单'), 'order')
        invoices_existing, _, _ = sheet_snapshot(wb.Worksheets('发货'), 'invoice')
        new_o, new_i, errors, skipped, notices = calculate_plan(parsed, orders_existing, invoices_existing)
        warnings.extend(errors)
        # Additional read-only month-end warning. A parsed invoice counts as
        # present even when excluded from auto-add for another business reason.
        dn_files = discover_delivery_notes(root)
        invoice_files = [(p, cust) for p, kind, cust, _ in docs if kind == 'invoice']
        month_warnings, month_notices, dn_count = missing_month_end_invoices(
            parsed, orders_existing, invoices_existing,
            dn_files, invoice_files, today,
            warning_days=int(config.get('month_end_missing_invoice_window_days', 5)),
            fallback_orders=fallback_orders)
        warnings.extend(month_warnings)
        notices.extend(month_notices)
        note = ('尚未修改 Excel；原汇总表缺失，预览以空白新表为基础（正式执行才会创建）'
                if original is None else
                '尚未修改 Excel；现有汇总表为空，预览按新建表结构计算（正式执行才会填充）'
                if was_initialized or zero_placeholder else '尚未修改 Excel')
        target = original or (root / f'{EXCEL_PREFIX}{today:%d.%m.%Y}.xlsx')
        if apply:
            target = root / f"{EXCEL_PREFIX}{today:%d.%m.%Y}.xlsx"
            if target != original and target.exists():
                raise RuntimeError(f"今日目标 Excel 已存在，避免覆盖：{target}")
            BACKUPS.mkdir(parents=True, exist_ok=True)
            backup = None
            if original is not None:
                backup = BACKUPS / f'{stamp:%Y%m%d_%H%M%S}_{original.name}'
                shutil.copy2(original, backup)
            # Insert documents in chronological order and recompute group positions for each insert.
            for rec in sorted(new_o, key=lambda x: (x.document_date,x.contract)):
                append_rec(wb.Worksheets('订单'), rec)
            for rec in sorted(new_i, key=lambda x: (x.document_date,x.contract,x.invoice)):
                append_rec(wb.Worksheets('发货'), rec)
            stage = root / f"~autofill_{stamp:%Y%m%d_%H%M%S}.xlsx"
            try:
                wb.SaveAs(str(stage), FileFormat=51)  # 51 is standard XLSX, preserves real Excel structures
                wb.Close(SaveChanges=False)
                wb = None
                if not safe_zip(stage):
                    raise RuntimeError("新 Excel 文件结构验证失败，原文件未修改")
                os.replace(stage, target)
                if original is not None and target != original:
                    original.unlink()
                note = (f'已保存 {target.name}；备份位于 {backup}' if backup
                        else f'已新建 {target.name}；原目录没有旧汇总表，无需备份')
            finally:
                if stage.exists():
                    try: stage.unlink()
                    except Exception: pass
        if config.get('rename_current_week_folder', True):
            try:
                msg = update_week_dir(root, today, dry_run=not apply)
                if '数量为' in msg or '已存在，未改名' in msg:
                    warnings.append(msg)
                note += '；' + msg
            except Exception as e:
                warnings.append(f"重命名本周文件夹失败：{e}")
        report, report_json = write_report(stamp,root,original or "（尚无汇总表）",len(docs),new_o,new_i,warnings,skipped,notices,
                                           apply,note,str(target),ignored_cancelled=ignored_cancelled, delivery_notes_scanned=dn_count)
        audit_db(stamp,report,apply,len(new_o),len(new_i),warnings,parsed)
        print(f"{'[已执行]' if apply else '[只读预览]'} 扫描 {len(docs)} PDFs；订单 +{len(new_o)}；发货 +{len(new_i)}；发货单 {dn_count}；忽略作废发票 {len(ignored_cancelled)}；异常 {len(warnings)}")
        print(f"报告：{report}")
        print(note)
        if warnings and apply and config.get('popup_on_errors',True):
            popup_warning(f"本次发现 {len(warnings)} 项异常。订单 +{len(new_o)}，发货 +{len(new_i)}。\n详情：{report}")
        return 0 if not warnings else 2
    finally:
        if wb is not None:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if excel is not None:
            try: excel.Quit()
            except Exception: pass
        if preview_copy is not None and preview_copy.exists():
            try: preview_copy.unlink()
            except Exception: pass
        try: LOCK.unlink()
        except FileNotFoundError: pass


def main():
    parser = argparse.ArgumentParser(description="业务周汇总表 PDF 自动补录（默认只读预览）")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--preview', action='store_true', help='只读预览，默认')
    mode.add_argument('--apply', action='store_true', help='写入（自动备份）')
    args = parser.parse_args()
    if args.apply:
        print('Safety: Excel writing is now permitted only after GUI preview and explicit confirmation. Open 10_START_GUI.bat.', file=sys.stderr)
        return 3
    if not CONFIG_FILE.exists():
        print("尚未配置，请先运行 02_CONFIGURE.bat 或 python configure.py", file=sys.stderr)
        return 1
    try:
        return execute(json.loads(CONFIG_FILE.read_text(encoding='utf-8-sig')),apply=args.apply)
    except Exception as exc:
        STATE.mkdir(parents=True,exist_ok=True)
        error_file = STATE / f"ERROR_{datetime.now():%Y%m%d_%H%M%S}.txt"
        error_file.write_text(traceback.format_exc(), encoding='utf-8')
        print(f"严重异常：{exc}\n错误详情：{error_file}",file=sys.stderr)
        try:
            config = json.loads(CONFIG_FILE.read_text(encoding='utf-8-sig'))
            if args.apply and config.get('popup_on_errors',True):
                popup_warning(f"自动汇总失败：{exc}\n错误详情：{error_file}")
        except Exception: pass
        return 1

if __name__ == '__main__':
    sys.exit(main())
