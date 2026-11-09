"""Local per-employee Teams upload ledger; no cloud secrets are stored."""
from __future__ import annotations
import hashlib
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path


def sha_file(path: str | Path):
    h=hashlib.sha256()
    with open(path, 'rb') as f:
        for b in iter(lambda:f.read(1024*1024), b''):
            h.update(b)
    return h.hexdigest()


class TeamsLedger:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(self.path) as db:
            db.execute('''CREATE TABLE IF NOT EXISTS uploads(
                profile TEXT NOT NULL, source TEXT NOT NULL, sha TEXT NOT NULL,
                invoice TEXT NOT NULL, contract TEXT NOT NULL, drive_id TEXT NOT NULL,
                remote_id TEXT NOT NULL, url TEXT NOT NULL, name TEXT NOT NULL,
                status TEXT NOT NULL, uploaded_at TEXT NOT NULL,
                PRIMARY KEY (profile, source, sha, drive_id, remote_id))''')
            db.execute('CREATE INDEX IF NOT EXISTS upload_lookup ON uploads(profile,source,sha,status)')
            db.execute('''CREATE TABLE IF NOT EXISTS attempts(
                id INTEGER PRIMARY KEY AUTOINCREMENT, profile TEXT NOT NULL,
                source TEXT NOT NULL, invoice TEXT NOT NULL, contract TEXT NOT NULL,
                target TEXT NOT NULL, outcome TEXT NOT NULL, attempted_at TEXT NOT NULL)''')

    def has_uploaded(self, profile, source, sha):
        with sqlite3.connect(self.path) as db:
            r=db.execute("SELECT 1 FROM uploads WHERE profile=? AND source=? AND sha=? AND status='uploaded' LIMIT 1",
                         (profile,str(Path(source).resolve()),sha)).fetchone()
        return bool(r)

    def record(self, profile, row, remote, target):
        from pathlib import Path
        with sqlite3.connect(self.path) as db:
            db.execute('''INSERT OR IGNORE INTO uploads(profile,source,sha,invoice,contract,drive_id,
                remote_id,url,name,status,uploaded_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
                (profile,str(Path(row.source).resolve()),sha_file(row.source),row.invoice,row.contract,
                 target['drive_id'],remote['id'],remote.get('webUrl',''),remote.get('name',Path(row.source).name),
                 'uploaded',datetime.now(timezone.utc).isoformat(timespec='seconds')))

    def record_manual_batch(self, profile, rows, target_label, fingerprints):
        """Atomic user-confirmed uploads; does not pretend to be Graph verification."""
        import uuid
        rows=tuple(rows)
        fingerprints=tuple(fingerprints)
        if not rows or len(rows)!=len(fingerprints):
            raise ValueError('无有效发票待确认')
        values=[]
        for row, expected in zip(rows,fingerprints):
            current=sha_file(row.source)
            if current!=expected:
                raise ValueError(f'PDF 已改变，未记录完成：{Path(row.source).name}')
            values.append((row,current))
        now=datetime.now(timezone.utc).isoformat(timespec='seconds')
        with sqlite3.connect(self.path) as db:
            for row, current in values:
                resolved=str(Path(row.source).resolve())
                db.execute('''INSERT OR IGNORE INTO uploads(profile,source,sha,invoice,contract,drive_id,
                     remote_id,url,name,status,uploaded_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
                     (profile,resolved,current,row.invoice,row.contract,
                      'user-confirmed',str(uuid.uuid4()),'',Path(row.source).name,
                      'uploaded',now))
                db.execute('''INSERT INTO attempts(profile,source,invoice,contract,target,outcome,attempted_at)
                     VALUES(?,?,?,?,?,?,?)''',
                     (profile,resolved,row.invoice,row.contract,target_label,
                      '用户在 Teams 核实后人工标记上传成功（非 Graph 自动验证）',now))

    def record_manual(self, profile, row, target_label):
        return self.record_manual_batch(profile,(row,),target_label,(sha_file(row.source),))

    def list_history(self, profile, limit=200):
        with sqlite3.connect(self.path) as db:
            db.row_factory=sqlite3.Row
            r=db.execute('SELECT invoice,contract,name,url,uploaded_at,drive_id FROM uploads WHERE profile=? AND status="uploaded" ORDER BY uploaded_at DESC LIMIT ?',
                         (profile,limit)).fetchall()
        return [dict(x) for x in r]

    def log_attempt(self, profile, row, target, outcome):
        with sqlite3.connect(self.path) as db:
            db.execute('''INSERT INTO attempts(profile,source,invoice,contract,target,outcome,attempted_at)
                VALUES(?,?,?,?,?,?,?)''',
                (profile,str(Path(row.source).resolve()),row.invoice,row.contract,
                 target.get('label',''),outcome,datetime.now(timezone.utc).isoformat(timespec='seconds')))

    def list_attempts(self, profile, limit=300):
        with sqlite3.connect(self.path) as db:
            db.row_factory=sqlite3.Row
            r=db.execute('SELECT source,invoice,contract,target,outcome,attempted_at FROM attempts WHERE profile=? ORDER BY id DESC LIMIT ?',
                         (profile,limit)).fetchall()
        return [dict(x) for x in r]
