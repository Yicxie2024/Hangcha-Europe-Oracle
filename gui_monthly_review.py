"""Monthly invoice review decisions and stable, source-bound audit records.

Decisions are per YYYY-MM/invoice and do not alter source PDFs or weekly Excel.
An old decision becomes stale if the matched PDF or essential weekly row changes.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional
import hashlib
import json
import os
import tempfile

from gui_monthly_core import Invoice

@dataclass
class ReviewEntry:
    invoice: str
    state: str          # 纳入 / 待复核 / 排除 (scanner's unmodified result)
    reason: str
    file: str
    count: int = 0
    record: Optional[Invoice] = None
    can_confirm: bool = False
    fingerprint: str = ''
    decision: str = ''  # confirmed / removed / ''
    stale: bool = False

    @property
    def display_state(self) -> str:
        if self.stale:
            return '来源已变化'
        if self.decision == 'removed':
            return '已移除'
        if self.decision == 'confirmed':
            return '已确认'
        if self.state == '纳入':
            return '正常'
        if self.state == '排除':
            return '不符合条件'
        return '待复核'

    @property
    def included(self) -> bool:
        if self.stale or self.decision == 'removed':
            return False
        if self.decision == 'confirmed':
            return self.can_confirm and self.record is not None
        return self.state == '纳入' and self.record is not None


def fingerprint(invoice: str, state: str, reason: str,
                summary_details: str = '', record: Optional[Invoice] = None) -> str:
    """Fingerprint changes when core PDF or weekly invoice facts change."""
    data = [invoice, state, reason, summary_details,
            record.digest if record else '',
            record.day.isoformat() if record else '',
            str(record.net) if record else '',
            str(record.gross) if record else '']
    return hashlib.sha256('\0'.join(data).encode('utf-8')).hexdigest()


def load_decisions(path: str | Path, ym: str) -> dict:
    try:
        obj = json.loads(Path(path).read_text(encoding='utf-8'))
        if obj.get('schema') != 1:
            return {}
        return obj.get('months', {}).get(ym, {})
    except (OSError, ValueError, AttributeError, TypeError):
        return {}


def apply_decisions(entries: list[ReviewEntry], path: str | Path, ym: str) -> None:
    decisions = load_decisions(path, ym)
    for entry in entries:
        saved = decisions.get(entry.invoice)
        if not isinstance(saved, dict) or saved.get('action') not in ('confirmed', 'removed'):
            continue
        if saved.get('fingerprint') == entry.fingerprint:
            entry.decision = saved['action']
        else:
            entry.stale = True


def store_decision(path: str | Path, ym: str, entry: ReviewEntry, action: str) -> None:
    if action not in ('confirmed', 'removed', 'reset'):
        raise ValueError('无效复核操作')
    if action == 'confirmed' and (not entry.can_confirm or entry.record is None):
        raise ValueError('当前发票的编号、日期、金额或目的国尚未核实，不能强制纳入')
    p = Path(path)
    try:
        obj = json.loads(p.read_text(encoding='utf-8'))
        if obj.get('schema') != 1:
            obj = {'schema': 1, 'months': {}}
    except (OSError, ValueError, AttributeError, TypeError):
        obj = {'schema': 1, 'months': {}}
    months = obj.setdefault('months', {})
    decisions = months.setdefault(ym, {})
    if action == 'reset':
        decisions.pop(entry.invoice, None)
    else:
        decisions[entry.invoice] = {'action': action, 'fingerprint': entry.fingerprint,
                                    'at': datetime.now(timezone.utc).isoformat(),
                                    'pdf': entry.file,
                                    'recognized_items': len(entry.record.items) if entry.record else 0}
    p.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix='monthly_review_', suffix='.tmp', dir=p.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(obj, f, indent=2, ensure_ascii=False)
            f.flush(); os.fsync(f.fileno())
        os.replace(tmp, p)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)
    entry.decision = '' if action == 'reset' else action
    entry.stale = False


def selected_invoices(entries: list[ReviewEntry]) -> list[Invoice]:
    return sorted([x.record for x in entries if x.included and x.record is not None],
                  key=lambda r: (r.day, r.invoice))


def outcome_rows(entries: list[ReviewEntry]):
    from gui_monthly_core import Outcome
    return [Outcome(e.file, e.invoice, e.display_state, e.reason, e.count) for e in entries]
