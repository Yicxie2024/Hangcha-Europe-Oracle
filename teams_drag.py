"""Offline safety rules for dragging local sales invoices into Microsoft Teams.

This module never uploads or writes to SharePoint. Windows drag/drop transfers
file paths to the target app, which is solely responsible for uploading.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Sequence

from teams_store import sha_file


class DragSelectionError(ValueError):
    pass


@dataclass(frozen=True)
class DragBatch:
    paths: tuple[str, ...]
    fingerprints: tuple[str, ...]
    prepared_on: date
    label: str

    @property
    def size(self) -> int:
        return len(self.paths)


def expected_teams_folder(today: date) -> str:
    return f'{today.month}月 / {today:%d.%m}'


def prepare_files(rows: Sequence, today: date, allowed_states: set[str] | None = None) -> DragBatch:
    """Inspect the exact set the user chose; no guesses, no silent exclusions."""
    if not rows:
        raise DragSelectionError('请先勾选至少一张发票。')
    allowed_states = allowed_states or {'待上传'}
    paths, hashes, names, identities = [], [], [], []
    for row, state in rows:
        p = Path(row.source)
        if row.kind != '发票' or state not in allowed_states:
            raise DragSelectionError(f'不可拖出（{state}）：{p.name}')
        if '作废' in p.name or p.suffix.lower() != '.pdf':
            raise DragSelectionError(f'不是有效销售发票：{p.name}')
        if not row.contract or not row.invoice:
            raise DragSelectionError(f'缺少合同号或发票号：{p.name}')
        if not p.is_file() or p.stat().st_size <= 0:
            raise DragSelectionError(f'源文件不存在或为空：{p}')
        try:
            fingerprint = sha_file(p)
        except OSError as exc:
            raise DragSelectionError(f'无法读取 PDF：{p.name}: {exc}') from exc
        paths.append(str(p.resolve()))
        hashes.append(fingerprint)
        names.append(p.name.casefold())
        identities.append((str(row.contract), str(row.invoice)))
    duplicates = [name for name,n in Counter(names).items() if n > 1]
    if duplicates:
        raise DragSelectionError('选中多张同名 PDF，请先分别核对：' + ', '.join(duplicates))
    if len(set(identities)) != len(identities):
        raise DragSelectionError('所选文件存在重复的合同号 + 发票号，请先核对再拖拽。')
    return DragBatch(tuple(paths), tuple(hashes), today, expected_teams_folder(today))


def validate_batch(batch: DragBatch, today: date) -> None:
    if batch.prepared_on != today:
        raise DragSelectionError('电脑日期已经变化，请重新准备拖拽，并在 Teams 中打开新的当天文件夹。')
    for path, fingerprint in zip(batch.paths, batch.fingerprints):
        try:
            if sha_file(path) != fingerprint:
                raise DragSelectionError(f'源 PDF 在准备拖拽后已修改：{Path(path).name}')
        except OSError as exc:
            raise DragSelectionError(f'源 PDF 无法访问：{Path(path).name}: {exc}') from exc
