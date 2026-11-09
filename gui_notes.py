"""Local, per-employee, reversible Excel remark drafts.

No PDF / Excel mutation happens here. The GUI collects these drafts; the user
must review and explicitly approve the separate business write plan.
"""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

MAX_NOTE_LENGTH = 500


@dataclass(frozen=True)
class NoteDraft:
    kind: str          # 'order' or 'invoice' (Excel sheet, NOT PDF type)
    contract: str
    invoice: str
    text: str
    updated_at: str = ''

    def __post_init__(self):
        if self.kind not in ('order', 'invoice'):
            raise ValueError('备注目标只能是订单或发货')
        if not self.contract.strip().isdigit():
            raise ValueError('备注缺少有效合同号')
        if self.kind == 'invoice' and not self.invoice.strip():
            raise ValueError('发货备注必须指定发票号')
        if self.kind == 'order' and self.invoice:
            raise ValueError('订单备注不应指定发票号')
        if not self.text.strip() or len(self.text) > MAX_NOTE_LENGTH:
            raise ValueError(f'备注须为 1–{MAX_NOTE_LENGTH} 个字符')


class NoteStore:
    def __init__(self, db_path: Path, profile: str):
        self.db_path = Path(db_path)
        self.profile = profile
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as db:
            db.execute('''CREATE TABLE IF NOT EXISTS note_drafts (
                profile TEXT NOT NULL, kind TEXT NOT NULL, contract TEXT NOT NULL,
                invoice TEXT NOT NULL DEFAULT '', note_text TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(profile,kind,contract,invoice)
            )''')

    @contextmanager
    def _connect(self):
        connection = sqlite3.connect(str(self.db_path), timeout=10)
        try:
            with connection:
                yield connection
        finally:
            connection.close()

    def list_drafts(self) -> list[NoteDraft]:
        with self._connect() as db:
            items = db.execute('SELECT kind,contract,invoice,note_text,updated_at FROM note_drafts '
                               'WHERE profile=? ORDER BY updated_at,kind,contract,invoice',
                               (self.profile,)).fetchall()
        return [NoteDraft(*item) for item in items]

    def get(self, kind: str, contract: str, invoice: str = '') -> str:
        with self._connect() as db:
            item = db.execute('SELECT note_text FROM note_drafts WHERE profile=? AND kind=? '
                              'AND contract=? AND invoice=?',
                              (self.profile, kind, contract, invoice)).fetchone()
        return item[0] if item else ''

    def save(self, kind: str, contract: str, invoice: str, text: str) -> None:
        text = text.strip()
        draft = NoteDraft(kind, contract, invoice, text)
        stamp = datetime.now(timezone.utc).isoformat(timespec='microseconds')
        with self._connect() as db:
            db.execute('''INSERT INTO note_drafts(profile,kind,contract,invoice,note_text,updated_at)
                VALUES(?,?,?,?,?,?) ON CONFLICT(profile,kind,contract,invoice)
                DO UPDATE SET note_text=excluded.note_text, updated_at=excluded.updated_at''',
                (self.profile, draft.kind, draft.contract, draft.invoice, draft.text, stamp))

    def discard(self, kind: str, contract: str, invoice: str) -> None:
        with self._connect() as db:
            db.execute('DELETE FROM note_drafts WHERE profile=? AND kind=? AND contract=? AND invoice=?',
                       (self.profile, kind, contract, invoice))

    def discard_if_unchanged(self, drafts: list[NoteDraft]) -> int:
        """Never erase edits made after preview / while the write was underway."""
        removed = 0
        with self._connect() as db:
            for draft in drafts:
                cur = db.execute('DELETE FROM note_drafts WHERE profile=? AND kind=? AND contract=? '
                                 'AND invoice=? AND note_text=? AND updated_at=?',
                                 (self.profile, draft.kind, draft.contract, draft.invoice,
                                  draft.text, draft.updated_at))
                removed += cur.rowcount
        return removed
