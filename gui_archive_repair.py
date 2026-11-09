"""Read-only planning for historic leftover delivery notes.

No archive state changes until GUI explicitly shows a per-file confirmation and
ArchiveStore.archive rechecks hashes in one transaction.
"""
from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from gui_archive import ArchiveItem, path_key
from gui_scanner import ScanRow


def plan_historical_delivery_repair(rows: list[ScanRow], archived_records: list[ArchiveItem],
                                    archived_paths: set[str], is_reviewed):
    """Return (safe_groups, labelled_unresolved). Exact order refs only."""
    history = defaultdict(dict)
    for a in archived_records:
        if a.kind in {'订单','发票'} and a.mode in {'matched','pending'} and a.contract.strip():
            entry = history[a.contract.strip()].setdefault(a.group_id, {'mode':a.mode,'kinds':set()})
            entry['kinds'].add(a.kind)
    safe_contract_groups = {}
    ambiguous_history = set()
    for contract, archived_groups in history.items():
        groups = [gid for gid,detail in archived_groups.items()
                  if (detail['mode']=='matched' and {'订单','发票'} <= detail['kinds'])
                  or (detail['mode']=='pending' and '订单' in detail['kinds'])]
        if len(groups) == 1:
            safe_contract_groups[contract] = groups[0]
        elif len(groups) > 1:
            ambiguous_history.add(contract)

    # An active order of the same contract may represent a fresh contract cycle
    # or a conflict: never attach its slip to a historical group without review.
    active_orders = {r.contract.strip() for r in rows if r.kind == '订单'
                     and r.contract.strip() and path_key(r.source) not in archived_paths}
    matched = defaultdict(list)
    unresolved = []
    for row in rows:
        if row.kind != '发货单' or path_key(row.source) in archived_paths:
            continue
        label = f'{Path(row.source).name}（发货单号：{row.delivery_no or "未识别"}；Ordernumber：{row.contract or "未识别"}）'
        contract = row.contract.strip()
        if not contract:
            unresolved.append(('订单号未识别或不唯一',label)); continue
        if contract in active_orders:
            unresolved.append(('存在未归档订单，须先处理活动合同',label)); continue
        if contract in ambiguous_history:
            unresolved.append(('同合同存在多个历史归档组',label)); continue
        if contract not in safe_contract_groups:
            unresolved.append(('未找到唯一的已归档订单合同',label)); continue
        if row.status not in {'识别成功','重复文件'} and not (row.status == '待复核' and is_reviewed(row)):
            unresolved.append(('发货单待复核',label)); continue
        matched[contract].append(row)
    return [('delivery_followup', sorted(group, key=lambda r:r.source))
            for _,group in sorted(matched.items())], unresolved
