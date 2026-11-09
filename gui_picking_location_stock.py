"""Optional manually confirmed per-rack stock for a specific read-only Excel snapshot.

The Hangcha stock workbook stores ONE aggregate MA quantity on many rows whose
Rack No. contains multiple locations. It does not contain rack-level balances.
Never infer a split or write these supplemental values back to the workbook.
Reconfirm when the source Excel changes; the fingerprint includes path/mtime/size.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import json
import os
import hashlib

from gui_inventory_core import StockRecord, number_or_none, readable_number
from gui_picking_core import split_rack_locations


STORE = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'HangchaSalesAssistant' / 'picking_rack_stock.json'


def source_fingerprint(path: str | Path) -> str:
    p = Path(path).expanduser().resolve()
    stat = p.stat()
    sha = hashlib.sha256()
    sha.update(f'{os.path.normcase(str(p))}|{stat.st_size}|{stat.st_mtime_ns}|'.encode('utf-8'))
    # A saved workbook may be replaced while retaining a similar size; the
    # exact content hash ensures old manually confirmed rack splits are never
    # silently applied to a different inventory workbook.
    with p.open('rb') as src:
        for chunk in iter(lambda: src.read(1024 * 1024), b''):
            sha.update(chunk)
    return sha.hexdigest()


def entry_key(record: StockRecord, location: str) -> str:
    return '|'.join((str(record.row), record.erp.strip().upper(),
                     record.part.strip().upper(), str(record.stock),
                     record.rack.strip().upper(), location.strip().upper()))


def validate_group(record: StockRecord, values: dict[str, Decimal | None]) -> str | None:
    """Only accept nonnegative factual input; never infer missing balances.

    Partial entries may sum below the aggregate, but *fully specified* splits
    must match the known workbook total. Unknown aggregate totals are not
    accepted as a basis for a verified group.
    """
    locations = split_rack_locations(record.rack)
    if len(locations) < 2:
        return '仅允许为包含多个库位的库存记录补录独立库存量'
    if record.stock is None or record.stock < 0:
        return '该行 Excel 库存总数无效，无法核对分库位库存量'
    total = Decimal('0')
    complete = True
    for loc in locations:
        value = values.get(loc)
        if value is None:
            complete = False
        elif not value.is_finite() or value < 0:
            return f'{loc} 的库存必须为非负数字'
        else:
            total += value
    if total > record.stock:
        return f'已填分库位库存合计 {readable_number(total)} 大于 Excel 当前总库存 {readable_number(record.stock)}'
    if complete and total != record.stock:
        return f'分库位库存合计 {readable_number(total)} 与 Excel 总库存 {readable_number(record.stock)} 不一致'
    return None


class LocationStockStore:
    """Small verified metadata file bound to ONE source snapshot.

    This is optional reference data only. An Excel change invalidates the
    whole metadata snapshot; no stock movement or automatic deduction.
    """
    def __init__(self, source: str | Path, store: str | Path | None = None):
        self.source = Path(source)
        self.path = Path(store) if store is not None else STORE
        self.fingerprint = source_fingerprint(self.source)
        self.values: dict[str, Decimal] = {}
        self._load()

    def _load(self):
        try:
            saved = json.loads(self.path.read_text(encoding='utf-8'))
            if not isinstance(saved, dict) or saved.get('fingerprint') != self.fingerprint:
                return
            values = saved.get('values', {})
            if not isinstance(values, dict):
                return
            for k,v in values.items():
                n = number_or_none(v)
                if n is not None and n.is_finite() and n >= 0:
                    self.values[str(k)] = n
        except (OSError, ValueError, TypeError):
            return

    def get(self, record: StockRecord, location: str) -> Decimal | None:
        return self.values.get(entry_key(record, location))

    def group(self, record: StockRecord) -> dict[str, Decimal | None]:
        return {loc: self.get(record, loc) for loc in split_rack_locations(record.rack)}

    def update_groups(self, updates: list[tuple[StockRecord, dict[str, Decimal | None]]]):
        """Validate ALL groups, persist atomically, or reject without change."""
        if source_fingerprint(self.source) != self.fingerprint:
            raise ValueError('库存 Excel 已变化，库位补录数据不能继续使用；请重新查询后重新核对')
        staged = dict(self.values)
        for record, group in updates:
            locations = split_rack_locations(record.rack)
            err = validate_group(record, group)
            if err:
                raise ValueError(f'Excel 第 {record.row} 行（{record.erp or record.part}）：{err}')
            for loc in locations:
                key = entry_key(record, loc)
                value = group.get(loc)
                if value is None:
                    staged.pop(key, None)
                else:
                    staged[key] = value
        self.path.parent.mkdir(parents=True, exist_ok=True)
        target = {'version': 1, 'fingerprint': self.fingerprint,
                  'values': {k:readable_number(v) for k,v in staged.items()}}
        temporary = self.path.with_name(self.path.name+'.tmp')
        temporary.write_text(json.dumps(target, ensure_ascii=False, indent=2), encoding='utf-8')
        temporary.replace(self.path)
        self.values = staged

    def pdf_values(self, results) -> dict[tuple[int,str], Decimal]:
        output = {}
        for result in results:
            if result.excluded:
                continue
            for record in result.records:
                locations = split_rack_locations(record.rack)
                if len(locations) < 2:
                    continue
                candidate = self.group(record)
                # Re-validate on EACH use; stale or tampered metadata must
                # never become an authoritative per-rack quantity.
                if validate_group(record, candidate) is not None:
                    continue
                for loc in locations:
                    amount = candidate.get(loc)
                    if amount is not None:
                        output[(record.row, loc)] = amount
        return output
