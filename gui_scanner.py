"""Universal *read-only* PDF discovery/classification for Hangcha GUI.

No Excel access, renaming, file modification, or API calls.  The original
v1.0.7 automation.py remains the source of the verified PDF field extraction.
"""
from __future__ import annotations

import csv
import hashlib
import os
import re
import unicodedata
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

import automation
from gui_customer import extract_order_customer, extract_invoice_customer, extract_delivery_recipient

VERSION = "1.5.34"
HINTS = ("自动识别", "订单", "发票", "发货单")
KIND_DISPLAY = {"order": "订单", "invoice": "发票", "delivery": "发货单", "other": "其他/未分类"}
DEFAULT_EXCLUDED = {"muster", "模板", ".venv", "venv", "__pycache__", "state", "backups", "backup", "运单", "shipping labels", "回收站", "$recycle.bin", "system volume information"}
# The only meaningful filename flags are the user-defined workflow markers
# "作废" (ignore) and "有误" (monetary-comparison exemption in automation.py).
# No name, document type, contract, date or customer may come from a filename.
CREDIT_NOTE_HEADING = re.compile(r"(?mi)^\s*(?:credit[\s_-]*note|credit[\s_-]*memo|gutschrift|贷项通知单|贷记通知单)\b", re.I)


@dataclass
class ScanRow:
    kind: str
    customer: str
    contract: str
    invoice: str
    day: str
    gross: str
    net: str
    status: str
    note: str
    source: str
    sha: str = ""
    customer_source: str = ""  # 订单 PDF / 发票 PDF / 订单合同关联 / 未确定
    customer_alias: str = ""   # Backward-compatible CSV field; never generated from paths.
    delivery_no: str = ""  # Delivery No. is distinct from Ordernumber/contract.

    @property
    def display_customer(self) -> str:
        """Never shorten a PDF name with an alias taken from a file or folder."""
        return self.customer


def norm(text: str) -> str:
    return unicodedata.normalize("NFKC", str(text)).strip().casefold()


def source_hint(kind: str) -> str:
    return {"自动识别": "", "订单": "order", "发票": "invoice", "发货单": "delivery"}.get(kind, "")


def content_hint(text: str) -> str:
    # Strong, pairwise field signatures; don't classify invoices as delivery notes
    # merely because they reference a delivery-note number.
    has_invoice = bool(re.search(r"\bInvoice\s+No\.?\b", text, re.I)) and bool(re.search(r"\bInvoice\s+date\b", text, re.I))
    has_order = bool(re.search(r"\bOrder\s+No\.?\b", text, re.I)) and "Orderdatum" in text
    if has_invoice and not has_order:
        return "invoice"
    if has_order and not has_invoice:
        return "order"
    if has_invoice and has_order:
        return "conflict"
    # Read the document's OWN title. Hangcha uses "Delivery order" with a
    # separate "Delivery No." field; neither number is the sales contract.
    # Invoice and sales-order signatures above still take precedence, so a
    # reference to a delivery note in an invoice cannot recategorize it.
    if re.search(r"(?mi)^\s*(?:delivery\s*(?:order|note)|lieferschein|发货单|送货单)\s*(?:no\.?|nr\.?|number|#|:|$)", text):
        return "delivery"
    return ""


def read_pdf_signature(path: Path) -> str:
    import pdfplumber
    if path.stat().st_size == 0:
        raise ValueError("PDF 是空文件")
    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            raise ValueError("PDF 没有页面")
        idx = list(dict.fromkeys([0, min(1,len(pdf.pages)-1), len(pdf.pages)-1]))
        # Hangcha template has overprinted, doubled characters in the logo,
        # customer address and document heading (e.g. DDeelliivveerryy).
        # Dedupe before classification, never compensate with a filename.
        return "\n".join((pdf.pages[i].dedupe_chars(tolerance=1.5).extract_text() or "") for i in idx)


def classify(path: Path, source_root: Path, type_hint: str, text: str) -> tuple[str, str]:
    """Only the PDF body can establish document type.

    Directory labels and filenames are traversal metadata, never business data.
    The configured type hint is deliberately ignored when content is absent or
    contradictory, so a misplaced document cannot be silently misclassified.
    """
    by_content = content_hint(text)
    if by_content == 'conflict':
        return 'other', 'PDF 正文同时出现订单与发票页眉，需要人工确认'
    if by_content:
        return by_content, '已由 PDF 正文确认'
    return 'other', 'PDF 正文未找到可靠的订单/发票/发货单页眉，待人工复核'


def iter_files(sources: list[dict], excluded: set[str], cancelled=None):
    """Yield (path, source_root, hint) without duplicates even for overlapping roots."""
    visited = set()
    for entry in sources:
        root = Path(entry["path"]).expanduser()
        hint = entry.get("hint", "自动识别")
        if not root.is_dir():
            yield (None, root, hint)
            continue
        for folder, dirs, files in os.walk(root, topdown=True, followlinks=False):
            if cancelled and cancelled.is_set():
                return
            dirs[:] = sorted(d for d in dirs if norm(d) not in excluded and not (Path(folder)/d).is_symlink()
                              and not getattr(Path(folder)/d, "is_junction", lambda:False)())
            for name in sorted(files):
                if cancelled and cancelled.is_set():
                    return
                if not name.lower().endswith(".pdf") or name.startswith("~$"):
                    continue
                p = Path(folder) / name
                key = os.path.normcase(str(p.resolve()))
                if key in visited:
                    continue
                visited.add(key)
                yield (p, root, hint)


def scan_one(path: Path, root: Path, hint: str, salesperson: str = "Xie") -> ScanRow:
    file_name = unicodedata.normalize('NFKC', path.name)
    try:
        # These TWO user-defined filename flags are workflow instructions,
        # never sources of business information. Void files are intentionally
        # ignored without parsing their contents.
        if '作废' in file_name:
            return ScanRow('作废文件', '未确定', '', '', '', '', '',
                           '忽略作废', '文件名标记「作废」，按原规则完全忽略', str(path))
        text = read_pdf_signature(path)
        customer = '未确定'
        customer_source = '未确定'
        if CREDIT_NOTE_HEADING.search(text):
            return ScanRow('贷项通知单', customer, '', '', '', '', '', '识别成功',
                           'PDF 正文识别为 Credit Note / Gutschrift；仅供单独归档', str(path))
        kind, classify_note = classify(path, root, hint, text)
        if kind == 'other':
            return ScanRow('其他/未分类', customer, '', '', '', '', '', '待复核', classify_note, str(path))
        if kind == 'order':
            try:
                name = extract_order_customer(path)
            except Exception:
                name = None
            if name:
                customer, customer_source = name, '订单 PDF'
        elif kind == 'invoice':
            try:
                name = extract_invoice_customer(path)
            except Exception:
                name = None
            if name:
                customer, customer_source = name, '发票 PDF'
        if kind == 'delivery':
            try:
                recipient = extract_delivery_recipient(path)
            except Exception:
                recipient = None
            if recipient:
                customer, customer_source = recipient, '发货单 PDF（收货公司）'
            # STRICT: link by explicit "Ordernumber", never by "Delivery No.".
            # Multiple different order numbers are ambiguous and never joined.
            refs = automation.delivery_order_refs(text)
            contract = next(iter(refs)) if len(refs) == 1 else ''
            delivery_num, departure_day = automation.delivery_header_fields(text)
            details = [f'【发货单号】{delivery_num}' if delivery_num else '【发货单号待核】未能从 Delivery No. 表头提取',
                       f'【关联合同】Ordernumber={contract}' if contract else
                       '【关联合同待核】未找到唯一的 Ordernumber（不使用 Delivery No.）',
                       '【收货公司】' + customer,
                       f'【实际发货日期】{departure_day}' if departure_day else '【实际发货日期待核】Date of departure 未识别',
                       '仅作发货证明及缺票核对；不会凭发货单新增 Excel 发货表记录']
            if len(refs) > 1:
                details.append('PDF 包含多个不同 Ordernumber（均已识别，需人工确认多合同关联）：' + ', '.join(sorted(refs)))
            status = ('识别成功' if len(refs) == 1 and recipient and delivery_num and departure_day else '待复核')
            if not recipient:
                details.append('【收货公司待核】无法从发货单 PDF 左上方地址块可靠提取')
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            return ScanRow('发货单', customer, contract, '', departure_day, '', '', status,
                           '; '.join(details), str(path), sha=digest,
                           customer_source=customer_source, delivery_no=delivery_num)
        try:
            record = automation.extract_pdf(path, kind, customer, salesperson,
                                            customer_selected_by_scanner=True)
            issues = []
            if record.gross == 0 or record.net == 0:
                issues.append('订单/发票金额为 0，必须人工确认')
            if record.gross < 0 or record.net < 0:
                issues.append('存在负金额，请人工确认')
            if kind == 'invoice' and record.skip_order_invoice_amount_check:
                issues.append('「有误」发票：后续金额对比豁免（非异常）')
            real_issues = [i for i in issues if '非异常' not in i]
            if customer == '未确定':
                issues.append(CUSTOMER_WAIT_NOTE)
                real_issues.append(CUSTOMER_WAIT_NOTE)
            note = '; '.join(issues) or classify_note
            if customer_source == '订单 PDF':
                note += '; 【客户来源】订单 PDF 左侧客户地址块'
            elif customer_source == '发票 PDF':
                note += '; 【客户来源】发票 PDF 客户地址块'
            return ScanRow(KIND_DISPLAY[kind], record.customer, record.contract, record.invoice,
                           record.document_date.strftime('%d.%m.%Y'), str(record.gross), str(record.net),
                           '待复核' if real_issues else '识别成功', note, str(path),
                           record.sha, customer_source, '')
        except Exception as exc:
            refs = automation.explicit_contract_refs(text)
            contract = next(iter(refs)) if len(refs) == 1 else ''
            return ScanRow(KIND_DISPLAY[kind], customer, contract, '', '', '', '', '待复核',
                           f'PDF 正文类型已识别，但字段提取不完整：{exc}', str(path),
                           customer_source=customer_source)
    except Exception as exc:
        return ScanRow('其他/未分类', '', '', '', '', '', '', '待复核',
                       f'无法读取 PDF：{exc}', str(path))


CUSTOMER_WAIT_NOTE = '【客户待核】PDF 未能可靠识别客户，需关联合同订单 PDF 或人工复核'


def reconcile_order_customers(rows: list[ScanRow]) -> None:
    """Use the order-PDF addressee as the authoritative customer for a contract.

    An invoice-PDF addressee is used only when no reliable order-PDF name is
    available. Differences between order and invoice addressees are recorded
    for transparency, *not* turned into a review/error by themselves. Multiple
    DIFFERENT order-PDF addressees for one contract remain ambiguous and must
    not be silently resolved. If no order identifies the customer, disagreeing
    invoice-PDF names likewise require review. All inputs originate from PDF
    contents and are linked only by the PDF's internal contract number.
    """
    orders: dict[str, dict[str, str]] = {}
    invoices: dict[str, dict[str, str]] = {}
    for row in rows:
        if not row.contract or row.status in {'忽略作废', '重复文件'}:
            continue
        contract = row.contract.strip()
        if row.kind == '订单' and row.customer_source == '订单 PDF':
            orders.setdefault(contract, {})[norm(row.customer)] = row.customer
        elif row.kind == '发票' and row.customer_source == '发票 PDF':
            invoices.setdefault(contract, {})[norm(row.customer)] = row.customer

    # Only conflicting authorities are ambiguous. A different invoice name
    # does not create a conflict if exactly one order-PDF name is established.
    ambiguous = {
        contract for contract in set(orders) | set(invoices)
        if len(orders.get(contract, {})) > 1
        or (not orders.get(contract) and len(invoices.get(contract, {})) > 1)
    }

    for row in rows:
        # Delivery rows display their *own recipient* even when the order
        # customer differs. Contract relationship does NOT imply same recipient.
        if row.status in {'忽略作废', '重复文件'} or row.kind in {'贷项通知单', '发货单'} or not row.contract:
            continue
        contract = row.contract.strip()
        if contract in ambiguous:
            row.status = '待复核'
            if '【客户冲突】' not in row.note:
                found = orders.get(contract) or invoices.get(contract) or {}
                row.note = (row.note + '; ' if row.note else '') + (
                    '【客户冲突】同合同多个订单 PDF 客户名称不一致：'
                    if orders.get(contract) else
                    '【客户冲突】无订单客户依据，发票 PDF 客户名称不一致：'
                ) + ' / '.join(sorted(found.values()))
            continue
        official_names = orders.get(contract) or invoices.get(contract) or {}
        if len(official_names) != 1:
            continue
        official = next(iter(official_names.values()))
        has_order = contract in orders
        if row.customer != official:
            # Retain the invoice's independently extracted PDF customer in the
            # human-readable note; the business record uses the order customer.
            if (has_order and row.kind == '发票' and
                    row.customer_source == '发票 PDF' and row.customer != '未确定'):
                row.note += ('; ' if row.note else '') + (
                    f'【客户名称差异（仅提示）】发票 PDF：{row.customer}；'
                    f'订单 PDF：{official}。按订单 PDF 客户名处理'
                )
            else:
                row.note += ('; ' if row.note else '') + '【客户来源】按 PDF 内合同号关联'
            row.customer = official
            row.customer_source = '订单合同关联' if has_order else '发票合同关联'
        if CUSTOMER_WAIT_NOTE in row.note:
            row.note = row.note.replace('; ' + CUSTOMER_WAIT_NOTE, '').replace(
                CUSTOMER_WAIT_NOTE + '; ', '').replace(CUSTOMER_WAIT_NOTE, '').strip(' ;')
            if row.status == '待复核' and not any(
                x in row.note for x in ('金额为 0', '负金额', '字段提取不完整', '客户冲突')
            ):
                row.status = '识别成功'
    for row in rows:
        if row.customer == '未确定' and row.kind in {'订单', '发票'} and row.status != '忽略作废':
            row.status = '待复核'
            if CUSTOMER_WAIT_NOTE not in row.note:
                row.note += '; ' + CUSTOMER_WAIT_NOTE


def scan_sources(sources: list[dict], salesperson="Xie", excluded_names=None,
                 cancelled=None, on_row: Callable[[ScanRow], None] | None = None) -> list[ScanRow]:
    excluded = set(DEFAULT_EXCLUDED)
    excluded |= {norm(name) for name in (excluded_names or []) if name.strip()}
    out = []
    for path, root, hint in iter_files(sources, excluded, cancelled):
        if path is None:
            r = ScanRow("目录异常", "", "", "", "", "", "", "待复核", "来源目录不存在或无法访问", str(root))
        else:
            r = scan_one(path, root, hint, salesperson)
        out.append(r)
        if on_row:
            on_row(r)
    if not (cancelled and cancelled.is_set()):
        reconcile_order_customers(out)
    return out


def flag_duplicate_content(rows: list[ScanRow]):
    """Mark second and later copies of identical PDF content in one run."""
    seen = {}
    for r in rows:
        if not r.sha or r.status == "忽略作废":
            continue
        if r.sha in seen:
            original = seen[r.sha]
            r.status = "重复文件"
            r.note = f"与另一目录中的 PDF 内容完全相同，作为重复副本跳过：{original}"
        else:
            seen[r.sha] = r.source



# The invoice status is derived only after a COMPLETE directory scan.  An
# invoice may have been manually recorded in Excel without a PDF in these
# directories, so the first-pass status MUST be described as provisional.
_INVOICE_STATUS_NOTE = "【开票状态】"


def mark_pending_invoice_status(rows: list[ScanRow], excel_invoice_contracts=None) -> int:
    """Label orders with no matching non-void invoice; never change PDF data.

    Called after a complete scan, and again after a successful read-only Excel
    business preview.  Exact contract numbers are used, not delivery-note IDs.
    Orders requiring field review and duplicate copies retain their status.
    """
    known = {str(x).strip() for x in (excel_invoice_contracts or ()) if str(x).strip()}
    unmatched_invoice_pdfs = False
    for r in rows:
        if r.kind != "发票" or r.status == "忽略作废":
            continue
        if r.contract:
            known.add(str(r.contract).strip())
        elif r.status != "重复文件":
            unmatched_invoice_pdfs = True

    verified_excel = excel_invoice_contracts is not None
    pending = 0
    for r in rows:
        if r.kind != "订单" or r.status not in {"识别成功", "待开票"} or not r.contract:
            continue
        # Rebuild the supplemental note rather than appending on every preview.
        base_note = r.note.split(_INVOICE_STATUS_NOTE, 1)[0].rstrip(" ;")
        if str(r.contract).strip() in known:
            r.status = "识别成功"
            if verified_excel:
                msg = "已核对：对应合同在有效发票 PDF 或 Excel 发货表中存在记录"
            else:
                msg = "扫描到相应有效发票 PDF（Excel 尚未核对）"
        else:
            r.status = "待开票"
            pending += 1
            if verified_excel:
                msg = "PDF 扫描范围及 Excel 发货表均未找到该合同的有效发票记录"
            else:
                msg = "仅 PDF 初筛未见有效发票；请点击「核对业务 / 预览 Excel」进一步确认"
            if unmatched_invoice_pdfs:
                msg += "；注意：存在无法提取合同号的发票，需人工排查"
        r.note = (base_note + "; " if base_note else "") + _INVOICE_STATUS_NOTE + msg
    return pending

def summarize(rows: list[ScanRow]) -> dict:
    return {
        "pdf": sum(r.kind != "目录异常" for r in rows),
        "orders": sum(r.kind == "订单" and r.status in {"识别成功", "待开票"} for r in rows),
        "pending": sum(r.kind == "订单" and r.status == "待开票" for r in rows),
        "invoices": sum(r.kind == "发票" and r.status == "识别成功" for r in rows),
        "delivery": sum(r.kind == "发货单" for r in rows),
        "review": sum(r.status == "待复核" for r in rows),
        "cancelled": sum(r.status == "忽略作废" for r in rows),
        "duplicates": sum(r.status == "重复文件" for r in rows),
    }


CSV_COLUMNS = ["kind", "customer", "contract", "invoice", "day", "gross", "net", "status", "note", "source", "sha", "customer_source", "customer_alias", "delivery_no"]

def export_csv(rows: list[ScanRow], target: str):
    with open(target, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for r in rows:
            values = asdict(r)
            # Protect users opening exported CSV in Excel from spreadsheet formulas
            # inserted through untrusted PDF filenames, folder names, etc.
            for field, value in values.items():
                if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
                    values[field] = "'" + value
            writer.writerow(values)
