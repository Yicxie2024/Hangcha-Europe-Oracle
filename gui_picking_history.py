"""Per-user local audit trail for successful printing dispatch (not proof of paper output)."""
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import json
import os
import sqlite3

DB = Path(os.environ.get('LOCALAPPDATA',str(Path.home()))) / 'HangchaSalesAssistant' / 'picking_print_history.sqlite3'


def _connect():
    DB.parent.mkdir(parents=True,exist_ok=True)
    db=sqlite3.connect(DB,timeout=15)
    db.execute('''CREATE TABLE IF NOT EXISTS print_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        printed_at TEXT NOT NULL,
        order_number TEXT NOT NULL,
        pdf_path TEXT NOT NULL,
        picking_copies INTEGER NOT NULL,
        deliveries_json TEXT NOT NULL,
        inventory_path TEXT NOT NULL,
        result TEXT NOT NULL
    )''')
    return db


def count_submitted(order_number:str) -> int:
    if not order_number:return 0
    with _connect() as db:
        return int(db.execute('SELECT COUNT(*) FROM print_history WHERE order_number=? AND result=?',
                              (order_number,'submitted')).fetchone()[0])


def record_submitted(order_number:str,pdf_path:str,pick_copies:int,deliveries:list[tuple[str,int]],inventory_path:str):
    with _connect() as db:
        db.execute('''INSERT INTO print_history(printed_at,order_number,pdf_path,picking_copies,
                       deliveries_json,inventory_path,result) VALUES(?,?,?,?,?,?,?)''',
                   (datetime.now().isoformat(timespec='seconds'),order_number,pdf_path,int(pick_copies),
                    json.dumps(deliveries,ensure_ascii=False),inventory_path,'submitted'))
