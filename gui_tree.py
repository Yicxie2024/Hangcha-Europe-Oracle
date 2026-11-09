"""Pure, side-effect-free contract tree model for the Hangcha desktop GUI.

A parent describes a *contract*, never an individual PDF. Grouping requires
exact Hangcha internal contract numbers obtained from Order No. / Ordernumber.
Delivery note IDs, customer PO numbers, folder names and similar amounts must
never be used as substitutes. PDF files remain the underlying action units.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

from gui_scanner import ScanRow


@dataclass
class ContractNode:
    contract: str
    rows: list[ScanRow]
    primary: ScanRow
    state: str
    stage: str
    issues: int
    customer: str
    day: str
    gross: str
    net: str
    invoices: int


def amount(value: str) -> Decimal | None:
    try:
        raw = str(value).strip().replace('€', '').replace(' ', '')
        if ',' in raw and '.' in raw:
            raw = raw.replace('.', '').replace(',', '.') if raw.rfind(',') > raw.rfind('.') else raw.replace(',', '')
        elif ',' in raw:
            raw = raw.replace(',', '.')
        return Decimal(raw) if raw else None
    except (InvalidOperation, ValueError):
        return None


def _is_special(r: ScanRow) -> bool:
    """Non-business files never establish order/invoice relationships."""
    return r.status == '忽略作废' or r.kind in ('贷项通知单', '目录异常')


def _reliable_invoice(r: ScanRow) -> bool:
    return r.kind == '发票' and r.status in ('识别成功', '重复文件', '待开票') and not _is_special(r)


def _invoice_exempt(r: ScanRow) -> bool:
    return '有误' in Path(r.source).name


def _summary(contract: str, members: list[ScanRow], has_archived_invoice: bool = False) -> ContractNode:
    orders = [r for r in members if r.kind == '订单']
    primary = next((r for r in orders if r.status != '重复文件'), orders[0])
    invoices = [r for r in members if _reliable_invoice(r)]
    reviewed_files = [r for r in members if r.status == '待复核']
    issues = len(reviewed_files)
    extra_orders = max(0, len([r for r in orders if r.status != '重复文件']) - 1)
    issues += extra_orders
    duplicates = len([r for r in members if r.status == '重复文件'])
    # A PDF-only scan may have matching invoice files, but must not declare
    # that the contract is fully invoiced unless gross amounts can be compared
    # safely. Ignore duplicate copies and '有误' invoices for monetary sums.
    stage = '待开票'
    if invoices:
        stage = '已开票'
        if not reviewed_files and len([r for r in orders if r.status != '重复文件']) == 1:
            total_order = amount(primary.gross)
            normal_invoices = [r for r in invoices if r.status != '重复文件']
            if normal_invoices and len(normal_invoices) == len(invoices) and not any(_invoice_exempt(r) for r in normal_invoices):
                amounts = [amount(r.gross) for r in normal_invoices]
                if total_order is not None and total_order > 0 and amounts and all(v is not None and v >= 0 for v in amounts):
                    invoiced = sum(amounts, Decimal('0'))
                    if invoiced < total_order - Decimal('0.01'):
                        stage = '部分开票'
                    elif invoiced > total_order + Decimal('0.01'):
                        stage = '金额待核'
                        issues += 1
    elif any(r.kind == '发票' for r in members):
        stage = '票据待核'
    elif has_archived_invoice:
        stage = '已开票（归档）'
    elif primary.status == '识别成功':
        # Happens after an Excel preview confirms a shipping-sheet invoice
        # that has no PDF in the current scan roots.
        stage = '发货表已登记'
    state = stage
    if issues:
        state += f' · 需核{issues}'
    if duplicates:
        state += f' · 重复{duplicates}'
    return ContractNode(contract, members, primary, state, stage, issues,
                        primary.customer, primary.day, primary.gross, primary.net, len(invoices))


def make_tree(rows: Iterable[ScanRow], archived_paths: set[str] | None = None,
              path_key_fn=None) -> tuple[list[ContractNode], list[ScanRow]]:
    """Return order groups + unmatched documents without hiding any PDFs.

    All non-void, non-archived rows with a matching exact contract number are
    attached to that order; invoices without an order remain in unmatched.
    """
    archived = archived_paths or set()
    key_fn = path_key_fn or (lambda s: str(Path(s)).casefold())
    all_rows = list(rows)
    archived_invoice_contracts = {r.contract.strip() for r in all_rows
                                  if key_fn(r.source) in archived and _reliable_invoice(r) and r.contract.strip()}
    visible = [r for r in all_rows if key_fn(r.source) not in archived]
    orders_by_contract: dict[str, list[ScanRow]] = {}
    for r in visible:
        if r.kind == '订单' and r.contract and not _is_special(r):
            orders_by_contract.setdefault(r.contract.strip(), []).append(r)
    assigned: dict[str, list[ScanRow]] = {c: [] for c in orders_by_contract}
    unmatched = []
    for row in visible:
        contract = row.contract.strip()
        if contract in assigned and not _is_special(row) and row.kind in ('订单', '发票', '发货单', '其他/未分类'):
            assigned[contract].append(row)
        else:
            unmatched.append(row)
    groups = [_summary(contract, members, contract in archived_invoice_contracts)
              for contract, members in assigned.items()]
    return groups, unmatched


def matches_filter(row: ScanRow, selected: str, query: str) -> bool:
    if selected not in ('全部', row.kind, row.status):
        return False
    if query and query not in ' '.join((row.kind, row.customer, row.contract, row.invoice,
                                      row.delivery_no, row.day, row.gross, row.net,
                                      row.source, row.note)).casefold():
        return False
    return True
