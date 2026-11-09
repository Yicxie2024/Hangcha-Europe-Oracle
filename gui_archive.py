"""Per-employee reversible PDF archive and manual-review records.

Only a user-local SQLite database is modified. Original PDFs and Excel are never
moved, renamed, edited or deleted by this module.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import json
import hashlib
import os
import sqlite3
import uuid
from contextlib import contextmanager

from gui_scanner import ScanRow


def file_sha256(path: str | Path) -> str:
    with open(path, 'rb') as handle:
        h = hashlib.sha256()
        for block in iter(lambda: handle.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def path_key(path: str | Path) -> str:
    return os.path.normcase(os.path.abspath(str(path)))


@dataclass(frozen=True)
class ArchiveItem:
    path: str
    sha: str
    kind: str
    contract: str
    invoice: str
    group_id: str
    mode: str
    reason: str
    archived_at: str


class ArchiveStore:
    def __init__(self, db_path: str | Path, profile_id: str):
        self.path = Path(db_path)
        self.profile_id = str(profile_id).strip()
        if not self.profile_id:
            raise ValueError('必须设置员工资料 ID')
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as db:
            db.executescript('''
            CREATE TABLE IF NOT EXISTS archived (
                profile TEXT NOT NULL, path TEXT NOT NULL, sha TEXT NOT NULL,
                kind TEXT NOT NULL, contract TEXT NOT NULL, invoice TEXT NOT NULL,
                group_id TEXT NOT NULL, mode TEXT NOT NULL, reason TEXT NOT NULL,
                archived_at TEXT NOT NULL, PRIMARY KEY(profile, path)
            );
            CREATE INDEX IF NOT EXISTS archived_hash ON archived(profile, sha);
            CREATE INDEX IF NOT EXISTS archived_group ON archived(profile, group_id);
            CREATE TABLE IF NOT EXISTS reviewed (
                profile TEXT NOT NULL, path TEXT NOT NULL, sha TEXT NOT NULL,
                reason TEXT NOT NULL, reviewed_at TEXT NOT NULL,
                PRIMARY KEY(profile, path)
            );
            CREATE TABLE IF NOT EXISTS audit (
                id INTEGER PRIMARY KEY AUTOINCREMENT, profile TEXT NOT NULL,
                action TEXT NOT NULL, paths TEXT NOT NULL, event_at TEXT NOT NULL
            );
            ''')

    @contextmanager
    def _connect(self):
        db = sqlite3.connect(self.path, timeout=15)
        db.row_factory = sqlite3.Row
        try:
            yield db
            db.commit()
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @staticmethod
    def _stamp():
        return datetime.now(timezone.utc).isoformat(timespec='seconds')

    def _archived(self) -> list[ArchiveItem]:
        with self._connect() as db:
            records = db.execute('SELECT path, sha, kind, contract, invoice, group_id, mode, reason, archived_at '
                                 'FROM archived WHERE profile=? ORDER BY archived_at DESC, path',
                                 (self.profile_id,)).fetchall()
        return [ArchiveItem(**dict(item)) for item in records]

    def all_archived(self) -> list[ArchiveItem]:
        return self._archived()

    def archived_for(self, rows: list[ScanRow]) -> dict[str, ArchiveItem]:
        """Only retain an archive for the exact PDF bytes reviewed originally.

        File moves can be resolved by unique content SHA, but we never blindly
        hide a same-path file that has been replaced with different bytes.
        """
        stored = self._archived()
        by_path = {x.path: x for x in stored}
        by_hash: dict[str, list[ArchiveItem]] = {}
        for item in stored:
            by_hash.setdefault(item.sha, []).append(item)
        mapped = {}
        for row in rows:
            if row.kind == '目录异常':
                continue
            p = Path(row.source)
            if not p.is_file():
                continue
            key = path_key(p)
            direct = by_path.get(key)
            sha = row.sha or file_sha256(p)
            if direct and direct.sha == sha:
                mapped[key] = direct
            elif not direct:
                candidates = by_hash.get(sha, ())
                if len({r.group_id for r in candidates}) == 1 and candidates:
                    # Same bytes, renamed or moved. Original state is preserved.
                    mapped[key] = candidates[0]
        return mapped

    def is_reviewed(self, row: ScanRow) -> bool:
        if row.status != '待复核' or not Path(row.source).is_file():
            return False
        sha = row.sha or file_sha256(row.source)
        with self._connect() as db:
            res = db.execute('SELECT sha FROM reviewed WHERE profile=? AND path=?',
                             (self.profile_id, path_key(row.source))).fetchone()
        return res is not None and res['sha'] == sha

    def mark_reviewed(self, row: ScanRow, reason: str) -> None:
        if row.status != '待复核' or not Path(row.source).is_file():
            raise ValueError('只有现存的待复核 PDF 可以标记人工已复核')
        if not reason.strip():
            raise ValueError('请填写人工复核结论')
        sha = row.sha or file_sha256(row.source)
        with self._connect() as db:
            db.execute('INSERT OR REPLACE INTO reviewed VALUES (?,?,?,?,?)',
                       (self.profile_id, path_key(row.source), sha, reason.strip(), self._stamp()))
            db.execute('INSERT INTO audit(profile,action,paths,event_at) VALUES(?,?,?,?)',
                       (self.profile_id, 'reviewed', str(row.source), self._stamp()))

    def archive(self, groups: list[tuple[str, list[ScanRow]]], reason: str = '') -> list[str]:
        """Archive validated groups in ONE transaction. GUI must ask first.

        Group modes: matched, delivery_followup, void, reviewed, pending, credit. Relation determination
        is performed by planning outside this persistence module.
        """
        if not groups:
            return []
        checked = []
        seen = set()
        for mode, rows in groups:
            if mode not in {'matched', 'delivery_followup', 'void', 'reviewed', 'pending', 'credit'}:
                raise ValueError(f'不支持的归档模式: {mode}')
            if mode in {'void', 'reviewed', 'credit'} and len(rows) != 1:
                raise ValueError('单独归档模式只能对应一个 PDF')
            if mode == 'matched':
                if not (any(r.kind == '订单' for r in rows) and any(r.kind == '发票' for r in rows)):
                    raise ValueError('关联合同归档必须同时包含订单和发票')
                contracts = {r.contract.strip() for r in rows}
                if len(contracts) != 1 or '' in contracts or any(
                        r.kind not in {'订单', '发票', '发货单'} for r in rows):
                    raise ValueError('关联合同归档不允许混合合同号或无关文件类型')
                if any(r.kind == '发货单' and r.status not in {'识别成功', '重复文件'}
                       and not (r.status == '待复核' and self.is_reviewed(r)) for r in rows):
                    raise ValueError('关联发货单尚未可靠识别或人工复核，禁止整组归档')
            if mode == 'reviewed' and (rows[0].status != '待复核' or not self.is_reviewed(rows[0])):
                raise ValueError('单独归档待复核文件前必须先标记人工已复核')
            if mode == 'pending':
                orders = [r for r in rows if r.kind == '订单' and r.status == '待开票']
                if len(orders) != 1 or any(r.contract.strip() != orders[0].contract.strip() for r in rows) or \
                        any(r.kind not in {'订单', '发货单'} for r in rows) or not orders[0].contract.strip():
                    raise ValueError('待开票归档必须包含一个待开票订单及同合同号的可靠发货单')
                if any(r.kind == '发货单' and r.status not in {'识别成功', '重复文件'}
                       and not (r.status == '待复核' and self.is_reviewed(r)) for r in rows):
                    raise ValueError('待开票合同的关联发货单尚未可靠识别或人工复核')
            if mode == 'delivery_followup':
                if not rows or any(r.kind != '发货单' or not r.contract.strip() or
                                   r.contract.strip() != rows[0].contract.strip() for r in rows):
                    raise ValueError('补充归档只接受同一合同号的发货单 PDF')
                if any(r.status not in {'识别成功', '重复文件'} and
                       not (r.status == '待复核' and self.is_reviewed(r)) for r in rows):
                    raise ValueError('补充归档的发货单未经可靠识别或人工复核')
            if mode == 'credit' and (len(rows) != 1 or rows[0].kind != '贷项通知单' or rows[0].status not in {'识别成功', '重复文件'}):
                raise ValueError('贷项通知单必须已经成功识别；待复核文件请先人工复核')
            sha_rows = []
            for row in rows:
                key = path_key(row.source)
                if key in seen:
                    raise ValueError('同一 PDF 被重复放进归档计划')
                seen.add(key)
                p = Path(row.source)
                if not p.is_file():
                    raise FileNotFoundError(f'原始 PDF 已不存在: {p}')
                actual_sha = file_sha256(p)
                if row.sha and row.sha != actual_sha:
                    raise RuntimeError(f'PDF 在扫描后被修改，禁止归档，请重新扫描: {p}')
                sha_rows.append((row, key, actual_sha))
            checked.append((mode, sha_rows))
        paths = []
        with self._connect() as db:
            db.execute('BEGIN IMMEDIATE')
            for mode, rows in checked:
                if mode == 'delivery_followup':
                    # Attach legacy leftover delivery notes to ONE existing archived
                    # matched order+invoice group, so restore/undo stays atomic.
                    contract = rows[0][0].contract.strip()
                    archived = db.execute(
                        'SELECT group_id,kind,mode FROM archived WHERE profile=? AND contract=? '
                        'AND mode IN (?,?)',
                        (self.profile_id, contract, 'matched', 'pending')).fetchall()
                    groups_by_id = {}
                    for existing in archived:
                        entry = groups_by_id.setdefault(existing['group_id'], {'mode':existing['mode'],'kinds':set()})
                        entry['kinds'].add(existing['kind'])
                    eligible = [gid for gid, details in groups_by_id.items()
                                if (details['mode']=='matched' and {'订单','发票'} <= details['kinds'])
                                or (details['mode']=='pending' and '订单' in details['kinds'])]
                    if len(eligible) != 1:
                        raise ValueError(f'合同 {contract} 无唯一有效的历史归档订单组；禁止自动关联发货单')
                    group_id = eligible[0]
                else:
                    group_id = str(uuid.uuid4())
                for row, key, sha in rows:
                    existing = db.execute('SELECT sha FROM archived WHERE profile=? AND path=?',
                                          (self.profile_id, key)).fetchone()
                    if existing is not None:
                        raise ValueError(f'PDF 已有归档记录，禁止覆盖或重复归档: {key}')
                    db.execute('INSERT INTO archived VALUES (?,?,?,?,?,?,?,?,?,?)',
                               (self.profile_id, key, sha, row.kind, row.contract, row.invoice,
                                group_id, mode, reason.strip(), self._stamp()))
                    paths.append(key)
            db.execute('INSERT INTO audit(profile,action,paths,event_at) VALUES(?,?,?,?)',
                       (self.profile_id, 'archive', '\n'.join(paths), self._stamp()))
        return paths

    def source_state(self, item: ArchiveItem) -> str:
        """Check exact original bytes, not just a matching filename.

        Returns 'ok', 'missing' or 'changed'. Never mutates the PDF.
        A directory or unreadable file cannot be treated as a valid source.
        """
        p = Path(item.path)
        if not p.is_file():
            return 'missing'
        try:
            return 'ok' if file_sha256(p) == item.sha else 'changed'
        except OSError:
            return 'missing'

    def check_restore(self, paths: list[str]) -> tuple[list[str], list[tuple[ArchiveItem, str]]]:
        """Expand original linked groups and validate ALL members before restore."""
        full = self.expand_restore(paths)
        stored = {item.path: item for item in self._archived()}
        bad = []
        for p in full:
            item = stored.get(p)
            if item is None:
                raise RuntimeError('归档列表在检查期间发生变化，请刷新再试')
            state = self.source_state(item)
            if state != 'ok':
                bad.append((item, state))
        return full, bad

    def relink_missing(self, original_path: str, replacement_path: str | Path) -> ArchiveItem:
        """Rebind ONE archived record by exact SHA-256; leave its group intact.

        Never copy or restore bytes, and never accept a different PDF even if its
        filename/invoice number looks the same. Review history follows exact SHA.
        """
        key = path_key(original_path)
        target = path_key(replacement_path)
        candidate = Path(replacement_path)
        if not candidate.is_file() or candidate.suffix.casefold() != '.pdf':
            raise FileNotFoundError('选择的文件不存在，或不是 PDF')
        with self._connect() as db:
            db.execute('BEGIN IMMEDIATE')
            old = db.execute('SELECT * FROM archived WHERE profile=? AND path=?',
                             (self.profile_id, key)).fetchone()
            if old is None:
                raise KeyError('找不到对应的归档历史；请刷新归档管理')
            item = ArchiveItem(**{field:old[field] for field in ArchiveItem.__dataclass_fields__})
            if self.source_state(item) == 'ok':
                raise ValueError('原始 PDF 仍然有效，不需要重新定位')
            actual = file_sha256(candidate)
            if actual != item.sha:
                raise ValueError('SHA-256 文件指纹不同：不是归档时的原始 PDF，已拒绝重新关联')
            if target != key:
                collision = db.execute('SELECT 1 FROM archived WHERE profile=? AND path=?',
                                       (self.profile_id, target)).fetchone()
                if collision:
                    raise ValueError('目标 PDF 已存在另一条归档记录，不能覆盖')
                db.execute('UPDATE archived SET path=? WHERE profile=? AND path=? AND sha=?',
                           (target, self.profile_id, key, item.sha))
                old_review = db.execute('SELECT sha,reason,reviewed_at FROM reviewed '
                                        'WHERE profile=? AND path=?', (self.profile_id,key)).fetchone()
                if old_review is not None and old_review['sha'] == item.sha:
                    db.execute('INSERT OR REPLACE INTO reviewed VALUES(?,?,?,?,?)',
                               (self.profile_id,target,old_review['sha'],
                                old_review['reason'],old_review['reviewed_at']))
                    db.execute('DELETE FROM reviewed WHERE profile=? AND path=? AND sha=?',
                               (self.profile_id,key,item.sha))
            db.execute('INSERT INTO audit(profile,action,paths,event_at) VALUES(?,?,?,?)',
                       (self.profile_id,'relink_archive',
                        json.dumps({'from':key,'to':target,'sha':item.sha},ensure_ascii=False),self._stamp()))
        return ArchiveItem(target, item.sha, item.kind, item.contract, item.invoice,
                           item.group_id, item.mode, item.reason, item.archived_at)

    def clear_invalid(self, paths: list[str]) -> list[ArchiveItem]:
        """Explicitly clear ONLY selected stale archive records (never PDF files).

        Crucially does not expand the linked archive group. Other group members
        remain archived, and full metadata is retained in the durable audit log.
        """
        unique = list(dict.fromkeys(path_key(p) for p in paths))
        if not unique:
            return []
        deleted = []
        with self._connect() as db:
            db.execute('BEGIN IMMEDIATE')
            for key in unique:
                record = db.execute('SELECT * FROM archived WHERE profile=? AND path=?',
                                    (self.profile_id,key)).fetchone()
                if record is None:
                    raise KeyError(f'归档记录不存在，清理已取消：{key}')
                item = ArchiveItem(**{field:record[field] for field in ArchiveItem.__dataclass_fields__})
                if self.source_state(item) == 'ok':
                    raise ValueError(f'源 PDF 仍存在且指纹一致，禁止清除有效归档记录：{key}')
                deleted.append(item)
            for item in deleted:
                db.execute('DELETE FROM archived WHERE profile=? AND path=? AND sha=?',
                           (self.profile_id,item.path,item.sha))
            db.execute('INSERT INTO audit(profile,action,paths,event_at) VALUES(?,?,?,?)',
                       (self.profile_id,'clear_invalid_archive',
                        json.dumps([item.__dict__ for item in deleted],ensure_ascii=False),self._stamp()))
        return deleted

    def check_undo_last_archive(self, paths: list[str]):
        """Validate ONLY PDFs added in the last archive action, not old group members.

        Full contract restoration is deliberately a separate user action.
        """
        expected = {path_key(p) for p in paths}
        if not expected:
            return [], []
        with self._connect() as db:
            last = db.execute('SELECT action,paths FROM audit WHERE profile=? ORDER BY id DESC LIMIT 1',
                              (self.profile_id,)).fetchone()
            if not last or last['action'] != 'archive' or set(last['paths'].splitlines()) != expected:
                raise RuntimeError('归档记录已发生其他操作，本次撤销不可安全执行；请使用「归档管理」恢复')
        entries = {r.path:r for r in self._archived()}
        if not expected <= set(entries):
            raise RuntimeError('部分上次归档记录已不存在；请使用「归档管理」核查')
        invalid = [(entries[p],state) for p in expected
                   if (state:=self.source_state(entries[p])) != 'ok']
        return sorted(expected),invalid

    def undo_last_archive(self, paths: list[str]):
        """Undo exactly the files inserted by the latest archive operation.

        Crucially, historical order/invoice members sharing the attached
        delivery note's group_id are not affected by this Undo action.
        """
        expected, invalid = self.check_undo_last_archive(paths)
        if invalid:
            raise RuntimeError('上次归档文件缺失或发生变化，撤销已阻止')
        if not expected:
            return []
        with self._connect() as db:
            db.execute('BEGIN IMMEDIATE')
            last = db.execute('SELECT action,paths FROM audit WHERE profile=? ORDER BY id DESC LIMIT 1',
                              (self.profile_id,)).fetchone()
            if not last or last['action'] != 'archive' or set(last['paths'].splitlines()) != set(expected):
                raise RuntimeError('撤销期间归档记录发生变化，已停止')
            for p in expected:
                record = db.execute('SELECT sha FROM archived WHERE profile=? AND path=?',
                                    (self.profile_id,p)).fetchone()
                if record is None or file_sha256(p) != record['sha']:
                    raise RuntimeError(f'撤销期间 PDF 变化，已停止：{p}')
            db.executemany('DELETE FROM archived WHERE profile=? AND path=?',
                           [(self.profile_id,p) for p in expected])
            db.execute('INSERT INTO audit(profile,action,paths,event_at) VALUES(?,?,?,?)',
                       (self.profile_id,'undo_archive','\n'.join(expected),self._stamp()))
        return expected

    def expand_restore(self, paths: list[str]) -> list[str]:
        records = {x.path: x for x in self._archived()}
        groups = {records[path_key(p)].group_id for p in paths if path_key(p) in records}
        return sorted({r.path for r in records.values() if r.group_id in groups})

    def restore(self, paths: list[str]) -> list[str]:
        # The store is authoritative: callers (including Undo) may not bypass
        # missing/changed-file protection by skipping the GUI's preview.
        with self._connect() as db:
            db.execute('BEGIN IMMEDIATE')
            full = self.expand_restore(paths)
            if not full:
                return []
            stored = {item.path: item for item in self._archived()}
            invalid = []
            for p in full:
                state = self.source_state(stored[p])
                if state != 'ok':
                    invalid.append((p, state))
            if invalid:
                detail = '\n'.join(f'{state}: {p}' for p,state in invalid[:12])
                raise RuntimeError('归档组中存在源文件缺失或内容变化，已阻止整组恢复。\n'
                                   '请先重新定位 PDF 或单独清除失效归档记录：\n' + detail)
            db.executemany('DELETE FROM archived WHERE profile=? AND path=?',
                           [(self.profile_id, path) for path in full])
            db.execute('INSERT INTO audit(profile,action,paths,event_at) VALUES(?,?,?,?)',
                       (self.profile_id, 'restore', '\n'.join(full), self._stamp()))
        return full


def plan_archive(clicked: list[ScanRow], all_rows: list[ScanRow], archived_paths: set[str],
                 is_reviewed, archived_records: list[ArchiveItem] | None = None
                 ) -> tuple[list[tuple[str, list[ScanRow]]], list[str]]:
    """Safely plan contract-wide archival including reliable delivery notes.

    Only the PDF body's explicit, unique Ordernumber (stored in ScanRow.contract)
    can link a delivery note. A delivery-note number, filename or customer never
    establishes a contract relationship. Ambiguous/unreviewed slips are NOT
    silently archived. Existing archived matched or pending order groups may accept legacy
    leftover delivery slips after the user's explicit approval.

    The returned plan is read-only. GUI displays its complete scope and requests
    confirmation before ArchiveStore writes to SQLite.
    """
    indexed = {path_key(row.source): row for row in all_rows}
    historical = list(archived_records or ())
    seen: set[str] = set()
    groups: list[tuple[str, list[ScanRow]]] = []
    blocked: list[str] = []

    # Extend one uniquely identifiable historical order group, whether it
    # already had an invoice (matched) or was explicitly archived pending.
    archived_by_contract: dict[str, dict[str, dict]] = {}
    for item in historical:
        if item.mode in {'matched','pending'} and item.kind in {'订单','发票'} and item.contract.strip():
            entry = archived_by_contract.setdefault(item.contract.strip(), {}).setdefault(
                item.group_id, {'mode':item.mode, 'kinds':set()})
            entry['kinds'].add(item.kind)

    def old_archived_order_group(contract: str) -> bool:
        found = [gid for gid, details in archived_by_contract.get(contract, {}).items()
                 if (details['mode']=='matched' and {'订单','发票'} <= details['kinds'])
                 or (details['mode']=='pending' and '订单' in details['kinds'])]
        return len(found) == 1

    def delivery_safe(r: ScanRow) -> bool:
        return r.status in {'识别成功', '重复文件'} or (
            r.status == '待复核' and is_reviewed(r))

    for row in clicked:
        key = path_key(row.source)
        if key in seen or key in archived_paths:
            continue
        if row.status == '忽略作废':
            groups.append(('void', [row])); seen.add(key)
            continue
        if row.kind not in {'订单', '发票', '发货单', '贷项通知单'}:
            blocked.append(f'{Path(row.source).name}：此类文件尚不支持普通归档')
            continue
        if row.kind == '贷项通知单' and row.status in {'识别成功', '重复文件'}:
            groups.append(('credit', [row])); seen.add(key)
            continue
        if row.status == '待复核' and row.kind != '发货单':
            if not is_reviewed(row):
                blocked.append(f'{Path(row.source).name}：请先打开 PDF 查看并标记「已复核」')
            else:
                groups.append(('reviewed', [row])); seen.add(key)
            continue
        # Delivery notes with no reliable unique Ordernumber are never
        # contract-archived. An explicit manual review permits ONE-file archive.
        if row.kind == '发货单' and not row.contract.strip():
            if row.status == '待复核' and is_reviewed(row):
                groups.append(('reviewed', [row])); seen.add(key)
            else:
                blocked.append(f'{Path(row.source).name}：发货单未识别唯一 Ordernumber，须人工复核；不会猜测合同号')
            continue
        if not row.contract.strip() or row.status not in {'识别成功', '待开票', '重复文件', '待复核'}:
            blocked.append(f'{Path(row.source).name}：合同号或识别状态不适合关联归档')
            continue
        contract = row.contract.strip()
        related = [r for r in indexed.values()
                   if r.contract.strip() == contract and r.kind in {'订单', '发票', '发货单'}
                   and r.status != '忽略作废' and path_key(r.source) not in archived_paths]
        # Any uncertain document with the same unique contract must be handled
        # explicitly. Never hide only the easy files and strand the other slip.
        pending_review = [r for r in related if r.status == '待复核'
                          and (r.kind != '发货单' or not is_reviewed(r))]
        if pending_review:
            blocked.append(f'合同 {contract}：存在 {len(pending_review)} 份关联 PDF 待复核，'
                           '请先完成复核再归档（发货单不会被遗漏）')
            continue
        order_invoice = [r for r in related if r.kind in {'订单', '发票'}]
        deliveries = [r for r in related if r.kind == '发货单' and delivery_safe(r)]
        have_order = any(r.kind == '订单' for r in order_invoice)
        have_invoice = any(r.kind == '发票' for r in order_invoice)
        if have_order and have_invoice:
            together = order_invoice + deliveries
            if any(path_key(r.source) in seen for r in together):
                continue
            together.sort(key=lambda r: ({'订单': 0, '发票': 1, '发货单': 2}.get(r.kind, 3),
                                         r.invoice, r.source))
            groups.append(('matched', together))
            seen.update(path_key(r.source) for r in together)
        elif row.kind == '发货单' and old_archived_order_group(contract) and not order_invoice:
            # The order and invoices were archived using an older version;
            # offer a safe one-click follow-up without restoring them first.
            if not deliveries:
                blocked.append(f'合同 {contract}：没有可核实的发货单可补充归档')
                continue
            if any(path_key(r.source) in seen for r in deliveries):
                continue
            deliveries.sort(key=lambda r: r.source)
            groups.append(('delivery_followup', deliveries))
            seen.update(path_key(r.source) for r in deliveries)
        elif row.kind == '订单' and row.status == '待开票' and not have_invoice:
            # The explicit pending-order double confirmation also covers the
            # shipment proof, preventing a stray delivery-only orphan.
            pending_members = [row] + deliveries
            if any(path_key(r.source) in seen for r in pending_members):
                continue
            groups.append(('pending', pending_members))
            seen.update(path_key(r.source) for r in pending_members)
        elif row.kind == '发货单' and row.status == '待复核' and is_reviewed(row):
            # A reviewed delivery note may be archived individually if it
            # cannot be grouped with an active order+invoice or legacy group.
            groups.append(('reviewed', [row])); seen.add(key)
        else:
            blocked.append(f'合同 {contract}：未同时找到活动的订单和发票，'
                           '也没有可安全补充的历史关联合同归档组')
    return groups, blocked
