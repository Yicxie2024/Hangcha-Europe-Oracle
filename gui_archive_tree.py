"""Pure, non-mutating tree adapter for the archived-PDF management view.

The active results and archive management both use contract-parent/PDF-child
hierarchies. Archived rows keep their individual identity so that restore,
search, checkbox selection and right-click can act on the correct PDF.
"""
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from gui_scanner import ScanRow
from gui_archive import ArchiveItem, path_key


@dataclass
class ArchivedContract:
    contract: str
    primary: ScanRow
    rows: list[ScanRow]
    customer: str
    day: str
    gross: str
    net: str
    state: str
    stage: str
    issues: int
    invoices: int


def archive_tree(rows: Iterable[ScanRow], archived_map: dict[str, ArchiveItem],
                 records: Iterable[ArchiveItem], cached_delivery_numbers: dict[str,str] | None = None
                 ) -> tuple[list[ArchivedContract], list[ScanRow], dict[str, ArchiveItem]]:
    """Build the archive tree from both scan results and persisted archive rows.

    Missing original files must remain visible and restorable. Never infer a
    contract from a delivery-note number or group a void/credit note into an
    unrelated order. Only orders with an exact nonempty contract form parents.
    """
    records = list(records)
    cached_delivery_numbers = cached_delivery_numbers or {}
    present = set()
    archived_rows: list[ScanRow] = []
    entry_by_path: dict[str, ArchiveItem] = {}
    for row in rows:
        item = archived_map.get(path_key(row.source))
        if item:
            archived_rows.append(row)
            present.add(item.path)
            entry_by_path[path_key(row.source)] = item
    for item in records:
        if item.path in present:
            continue
        placeholder = ScanRow(item.kind, '', item.contract, item.invoice,
                              '', '', '', '已归档',
                              ('原始 PDF 已缺失，归档记录受到保护；请重新定位或单独清除失效历史'
                               if not Path(item.path).is_file() else
                               '原始 PDF 未在本次扫描中按原指纹确认（可能已移动、被修改或未扫描）；恢复前需校验文件指纹'),
                              item.path, item.sha,
                              delivery_no=(cached_delivery_numbers.get(item.path, '') if item.kind == '发货单' else ''))
        archived_rows.append(placeholder)
        entry_by_path[path_key(placeholder.source)] = item

    orders: dict[str, list[ScanRow]] = {}
    for row in archived_rows:
        if row.kind == '订单' and row.contract.strip() and row.status != '忽略作废' \
                and entry_by_path[path_key(row.source)].mode != 'void':
            orders.setdefault(row.contract.strip(), []).append(row)
    members = {contract: [] for contract in orders}
    unmatched = []
    for row in archived_rows:
        contract = row.contract.strip()
        entry = entry_by_path[path_key(row.source)]
        if contract in members and entry.mode != 'void' and row.kind in {'订单', '发票', '发货单', '其他/未分类'}:
            members[contract].append(row)
        else:
            unmatched.append(row)
    grouped = []
    for contract, files in members.items():
        primary = next((r for r in files if r.kind == '订单'), orders[contract][0])
        invoice_count = sum(r.kind == '发票' for r in files)
        unverified = sum(r.status == '已归档' or
                         not Path(entry_by_path[path_key(r.source)].path).is_file() for r in files)
        state = f'已归档 · {len(files)} 份 PDF' + (f' · 来源待核{unverified}' if unverified else '')
        grouped.append(ArchivedContract(contract, primary, files, primary.customer,
                                         primary.day, primary.gross, primary.net,
                                         state, '已归档', unverified, invoice_count))
    return grouped, unmatched, entry_by_path
