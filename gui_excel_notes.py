"""Read only the existing 订单 sheet's G-column remarks from .xlsx.

This intentionally uses the Python standard library and never launches Excel,
changes cell formats, writes workbook bytes, or touches pending note drafts.
The order contract (B) is the key. Duplicated contracts with conflicting
remarks are flagged rather than silently choosing a row.
"""
from __future__ import annotations

import posixpath
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
P = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CELL = re.compile(r'^([A-Z]+)(\d+)$')
CONTRACT = re.compile(r'^\d{5,9}$')


def _text_parts(element) -> str:
    return ''.join(x.text or '' for x in element.iter(S + 't'))


def _cell_value(cell, strings: list[str]) -> str:
    if cell.get('t') == 'inlineStr':
        inline = cell.find(S + 'is')
        return _text_parts(inline) if inline is not None else ''
    v = cell.find(S + 'v')
    raw = v.text if v is not None and v.text is not None else ''
    if cell.get('t') == 's':
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return ''
    return raw


def _normalize_contract(raw: str) -> str:
    value = str(raw).strip()
    if value.endswith('.0') and value[:-2].isdigit():
        value = value[:-2]
    return value if CONTRACT.fullmatch(value) else ''


def _order_sheet_path(z: ZipFile) -> str | None:
    book = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    sheet_id = next((s.get(R + 'id') for s in book.findall('.//' + S + 'sheet')
                     if s.get('name') == '订单'), None)
    if not sheet_id:
        return None
    target = next((r.get('Target', '') for r in rels if r.get('Id') == sheet_id), '')
    if not target:
        return None
    path = (target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target))
    # Restrict the package lookup to internal worksheet XML parts.
    if not path.startswith('xl/worksheets/') or '..' in path.split('/'):
        raise ValueError('Excel 订单工作表路径无效')
    return path


def read_order_notes(workbook: str | Path) -> dict[str, str]:
    """Return {contract_no: original_order_note}. Missing file/sheet => {}.

    Blank remarks are not shown. If two order rows for the same contract
    contain different nonblank remarks, display a review notice instead of
    selecting one arbitrarily. This does not affect any write validations.
    """
    path = Path(workbook).expanduser()
    if not path.is_file():
        return {}
    with ZipFile(path, 'r') as z:
        sheet_path = _order_sheet_path(z)
        if sheet_path is None:
            return {}
        strings: list[str] = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            strings = [_text_parts(si) for si in root.findall(S + 'si')]
        sheet = ET.fromstring(z.read(sheet_path))
        remarks: dict[str, set[str]] = {}
        for row in sheet.findall('.//' + S + 'sheetData/' + S + 'row'):
            cells = {}
            for cell in row.findall(S + 'c'):
                m = CELL.match(cell.get('r', ''))
                if m and m.group(1) in ('B', 'G'):
                    cells[m.group(1)] = _cell_value(cell, strings)
            contract = _normalize_contract(cells.get('B', ''))
            if contract:
                remarks.setdefault(contract, set()).add(cells.get('G', '').strip())
        result = {}
        for contract, values in remarks.items():
            nonempty = {v for v in values if v}
            if len(nonempty) > 1:
                result[contract] = '⚠ 多行备注不一致，请核对 Excel'
            elif nonempty:
                result[contract] = next(iter(nonempty))
        return result


def brief_note(value: str, limit: int = 90) -> str:
    """Single-line table preview; complete note remains available in details."""
    one_line = ' '.join(str(value).split())
    return one_line if len(one_line) <= limit else one_line[:limit-1] + '…'
