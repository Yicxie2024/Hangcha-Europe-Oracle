"""Read-only incremental PDF scanning, with per-employee local SQLite cache.

The cache accelerates extraction only.  It never makes archival/write decisions:
ArchiveStore.archive/restore and the business write pipeline re-read and verify
PDF SHA-256 before committing.  A manual verification scan hashes archives.
"""
from __future__ import annotations

from dataclasses import asdict, fields
from pathlib import Path
from typing import Callable
import json
import sqlite3
import os

from gui_scanner import (ScanRow, DEFAULT_EXCLUDED, norm, iter_files,
                         scan_one, reconcile_order_customers)
from gui_archive import ArchiveItem, path_key, file_sha256

# Deliberately independent of the UI version.  Bump this whenever PDF field
# extraction or reconciliation semantics change; old cached rows are rejected.
CACHE_SCHEMA_VERSION = 'scan-parser-1.5.34.1'


def _signature(p: Path):
    st = p.stat()
    return (int(st.st_size), int(st.st_mtime_ns), int(st.st_ctime_ns))


class ScanCache:
    """Persist pre-reconciliation ScanRows, never final dynamic business status."""
    def __init__(self, filename: str | Path, profile: str):
        self.path = Path(filename)
        self.profile = str(profile)
        if not self.profile:
            raise ValueError('scan cache needs an employee profile')
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as db:
            db.executescript('''
                CREATE TABLE IF NOT EXISTS parsed_pdf (
                    profile TEXT NOT NULL, path TEXT NOT NULL, parser TEXT NOT NULL,
                    salesperson TEXT NOT NULL, size INTEGER NOT NULL,
                    mtime_ns INTEGER NOT NULL, ctime_ns INTEGER NOT NULL,
                    sha TEXT NOT NULL, row_json TEXT NOT NULL,
                    PRIMARY KEY (profile,path)
                );
                CREATE INDEX IF NOT EXISTS idx_scan_cache_profile ON parsed_pdf(profile);
            ''')

    def _connect(self):
        return sqlite3.connect(self.path, timeout=20)

    def load(self):
        with self._connect() as db:
            rows = db.execute('SELECT path,parser,salesperson,size,mtime_ns,ctime_ns,sha,row_json '
                              'FROM parsed_pdf WHERE profile=?', (self.profile,)).fetchall()
        return {row[0]: row[1:] for row in rows}

    def store(self, updates):
        if not updates:
            return
        with self._connect() as db:
            db.executemany('''INSERT INTO parsed_pdf
                   (profile,path,parser,salesperson,size,mtime_ns,ctime_ns,sha,row_json)
                   VALUES (?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(profile,path) DO UPDATE SET
                   parser=excluded.parser,salesperson=excluded.salesperson,
                   size=excluded.size,mtime_ns=excluded.mtime_ns,
                   ctime_ns=excluded.ctime_ns,sha=excluded.sha,row_json=excluded.row_json''',
                   [(self.profile, *row) for row in updates])


def incremental_scan(sources: list[dict], salesperson: str, excluded_names,
                     cancelled, on_row: Callable[[ScanRow], None] | None,
                     cache: ScanCache, archived_records: list[ArchiveItem],
                     mode: str = 'incremental'):
    """Return (rows, statistics, {path:(size,mtime_ns,ctime_ns)}).

    Only changes extraction workload; directory enumeration is still complete.
    In 'full' mode all ACTIVE PDFs are reparsed.  In 'verify' mode all archive
    files are SHA-256 checked (without changing their persisted archive state).
    """
    if mode not in {'incremental','full','verify'}:
        raise ValueError('不支持的扫描模式')
    # Corrupt/locked cache is a performance problem, not a business-data
    # reason to stop scanning; fall back to full PDF extraction safely.
    try:
        stored = cache.load()
        cache_unavailable = False
    except (sqlite3.DatabaseError, OSError):
        stored = {}
        cache_unavailable = True
    archived_by_path = {item.path: item for item in archived_records}
    excluded = set(DEFAULT_EXCLUDED) | {norm(s) for s in (excluded_names or []) if str(s).strip()}
    updates = []
    out = []
    metadata = {}
    stats = {'cached':0,'reparsed':0,'archived_cached':0,'archive_verified':0,
             'archived_changed':0,'directory_errors':0,
             'cache_unavailable':int(cache_unavailable)}
    allowed_row_fields = {f.name for f in fields(ScanRow)}

    for pdf, root, hint in iter_files(sources, excluded, cancelled):
        if cancelled and cancelled.is_set():
            break
        if pdf is None:
            row = ScanRow('目录异常','','','','','','','待复核','来源目录不存在或无法访问',str(root))
            stats['directory_errors'] += 1
        else:
            key = path_key(pdf)
            try:
                before = _signature(pdf)
            except OSError:
                # Keep the original PDF-reader exception rather than dropping the file.
                before = None
            entry = stored.get(key)
            archive = archived_by_path.get(key)
            valid_cache = False
            cached_row = None
            if entry and before is not None:
                parser, cached_salesperson, size, mtime, ctime, sha, payload = entry
                if (parser == CACHE_SCHEMA_VERSION and cached_salesperson == salesperson
                    and before == (size,mtime,ctime)):
                    try:
                        data = json.loads(payload)
                        cached_row = ScanRow(**{k:v for k,v in data.items() if k in allowed_row_fields})
                        valid_cache = (cached_row.sha == sha and bool(sha))
                    except (ValueError,TypeError):
                        pass
            # A cached row has to agree with the *archived file's original SHA*.
            # An altered archived file never inherits the archived status.
            if archive and (not valid_cache or cached_row.sha != archive.sha):
                valid_cache = False
            if mode == 'full' and not archive:
                valid_cache = False
            if archive and mode == 'verify':
                try:
                    current_sha = file_sha256(pdf)
                    stats['archive_verified'] += 1
                except OSError:
                    current_sha = ''
                if current_sha != archive.sha:
                    stats['archived_changed'] += 1
                    valid_cache = False
            if valid_cache:
                row = cached_row
                if archive:
                    stats['archived_cached'] += 1
                else:
                    stats['cached'] += 1
            else:
                row = scan_one(pdf, root, hint, salesperson)
                stats['reparsed'] += 1
                # Some malformed/unclassified files lack hashes; hash once to
                # safely distinguish unchanged cached failures as well.
                if not row.sha and pdf.is_file():
                    try:
                        row.sha = file_sha256(pdf)
                    except OSError:
                        pass
                if archive and row.sha != archive.sha:
                    stats['archived_changed'] += (0 if mode == 'verify' else 1)
                try:
                    after = _signature(pdf)
                except OSError:
                    after = None
                # A file edited during parsing must not enter the cache.
                if before is not None and after == before and row.sha:
                    updates.append((key,CACHE_SCHEMA_VERSION,salesperson,*after,row.sha,
                                    json.dumps(asdict(row),ensure_ascii=False)))
            if before is not None:
                metadata[key] = before
        out.append(row)
        if on_row:
            on_row(row)

    # Atomic cache flush before mutable reconciliation and duplicate/status
    # decoration, so repeated launches never re-use a stale '待开票' label.
    if not (cancelled and cancelled.is_set()):
        try:
            cache.store(updates)
        except (sqlite3.DatabaseError, OSError):
            stats['cache_unavailable'] = 1
        reconcile_order_customers(out)
    return out,stats,metadata


def changed_since_scan(metadata):
    """Cheap race-condition check; final archive/write operations hash bytes."""
    changed = []
    for name, signature in metadata.items():
        try:
            if _signature(Path(name)) != signature:
                changed.append(name)
        except OSError:
            changed.append(name)
    return changed
